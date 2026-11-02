Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports BTA_OSG

Namespace BTA_OSG.Benchmarks
    ''' <summary>
    ''' Measures the client-side cache and mirror work that runs on every workstation:
    ''' full XML save/load, the per-tick snapshot merge, and the local-time normalization
    ''' added in the 2026-11-02 hardening batch. SQL read time, network latency, and UI
    ''' rendering are deliberately excluded; they are environment-specific and measured
    ''' separately by the office.
    ''' </summary>
    Module Program

        Private Const AuditRows As Integer = 1000

        Sub Main(args As String())
            Dim sizes = If(args.Length > 0, args.Select(Function(a) Integer.Parse(a)).ToArray(), New Integer() {5000, 20000})
            Dim root = Path.Combine(Path.GetTempPath(), "bta_bench_" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(root)

            Try
                AppSettings.Instance.DatabaseSettings.UseSqlServer = False

                Console.WriteLine("# BTA OSG cache and mirror benchmark")
                Console.WriteLine()
                Console.WriteLine("- Date: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                Console.WriteLine("- OS: " & RuntimeInformation.OSDescription)
                Console.WriteLine("- Runtime: " & RuntimeInformation.FrameworkDescription)
                Console.WriteLine("- CPUs: " & Environment.ProcessorCount.ToString())
                Dim cpu = Environment.GetEnvironmentVariable("PROCESSOR_IDENTIFIER")
                If Not String.IsNullOrWhiteSpace(cpu) Then Console.WriteLine("- CPU: " & cpu)
                Console.WriteLine("- SQL: disabled (offline cache benchmark)")
                Console.WriteLine()

                For Each n In sizes
                    RunSize(n, root)
                Next
            Finally
                Try
                    Directory.Delete(root, True)
                Catch
                End Try
            End Try
        End Sub

        Private Sub RunSize(n As Integer, root As String)
            EmbeddedDB.DbPath = Path.Combine(root, "bta_osg_" & n.ToString() & ".xml")
            EmbeddedDB.Initialize()
            Populate(n)

            ' Warm up and create the first .bak generation before timing.
            EmbeddedDB.Save()
            EmbeddedDB.Save()

            Dim saveMs = Median(3, Sub() EmbeddedDB.Save())
            Dim fileMb = SizeMb(EmbeddedDB.DbPath)
            Dim bakMb = SizeMb(EmbeddedDB.DbPath & ".bak")

            EmbeddedDB.Initialize()
            Dim loadMs = Median(3, Sub() EmbeddedDB.Initialize())

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            Dim heapMb = GC.GetTotalMemory(True) / 1024.0 / 1024.0

            Dim quietSnapshot = BuildSnapshot()
            ApplyQuiet(quietSnapshot)
            Dim quietMs = Median(3, Sub() ApplyQuiet(quietSnapshot))

            Dim changedSnapshot = EmbeddedDB.DataSet.Tables("Documents").Copy()
            Dim changedValue As Boolean = False
            EmbeddedDB.ApplySnapshot("Documents", changedSnapshot, persist:=False)
            Dim changedMs = Median(3, Sub()
                                          changedValue = Not changedValue
                                          changedSnapshot.Rows(0)("Title") = If(changedValue, "Changed title", "Benchmark document 1")
                                          EmbeddedDB.ApplySnapshot("Documents", changedSnapshot, persist:=False)
                                      End Sub)

            Dim warmCopies = PrepareLogCopies()
            For Each pair In warmCopies
                DesktopDataCoordinator.NormalizeMirroredTimes(pair.Key, pair.Value)
            Next

            Dim normalizeTimes As New List(Of Double)()
            For k = 1 To 3
                Dim copies = PrepareLogCopies()
                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()
                Dim sw = Stopwatch.StartNew()
                For Each pair In copies
                    DesktopDataCoordinator.NormalizeMirroredTimes(pair.Key, pair.Value)
                Next
                sw.Stop()
                normalizeTimes.Add(sw.Elapsed.TotalMilliseconds)
            Next
            normalizeTimes.Sort()
            Dim normalizeMs = normalizeTimes(normalizeTimes.Count \ 2)

            Console.WriteLine("## " & n.ToString("N0") & " documents, " & (n * 2).ToString("N0") & " routing logs, " &
                              n.ToString("N0") & " movements, " & (n \ 2).ToString("N0") & " directives, " &
                              AuditRows.ToString("N0") & " audit rows")
            Console.WriteLine()
            Console.WriteLine("| Metric | Median |")
            Console.WriteLine("|---|---|")
            Console.WriteLine("| Cache save (full XML, temp + replace) | " & Fmt(saveMs) & " ms |")
            Console.WriteLine("| Cache load (schema + XML + identity resync) | " & Fmt(loadMs) & " ms |")
            Console.WriteLine("| Quiet mirror merge, all tables | " & Fmt(quietMs) & " ms |")
            Console.WriteLine("| One-row change merge, Documents | " & Fmt(changedMs) & " ms |")
            Console.WriteLine("| Local-time normalization, 4 log tables | " & Fmt(normalizeMs) & " ms |")
            Console.WriteLine("| Cache file | " & Fmt(fileMb) & " MB |")
            Console.WriteLine("| Previous generation (.bak) | " & Fmt(bakMb) & " MB |")
            Console.WriteLine("| Managed heap after load | " & Fmt(heapMb) & " MB |")
            Console.WriteLine()
        End Sub

        Private Sub Populate(n As Integer)
            Dim docs = EmbeddedDB.DataSet.Tables("Documents")
            Dim directives = EmbeddedDB.DataSet.Tables("Directives")
            Dim routing = EmbeddedDB.DataSet.Tables("RoutingLogs")
            Dim movements = EmbeddedDB.DataSet.Tables("Movements")
            Dim audit = EmbeddedDB.DataSet.Tables("AuditTrail")
            Dim heartbeat = EmbeddedDB.DataSet.Tables("Heartbeat")

            Dim baseTime = New DateTime(2026, 1, 1, 8, 0, 0)
            Dim rowVersion As Byte() = New Byte() {0, 0, 0, 0, 0, 0, 7, 209}
            Dim statuses = {"RECEIVED", "FOR_REVIEW", "FOR_REVISION", "APPROVED"}

            docs.BeginLoadData()
            Try
                For i = 1 To n
                    Dim r = docs.NewRow()
                    r("DocCode") = "COMM-2026-" & i.ToString("D5")
                    r("DocType") = "Regular Communication"
                    r("Title") = "Benchmark document " & i.ToString()
                    r("OriginatingOffice") = "Records Section"
                    r("DestinationOffice") = "Secretariat"
                    r("CabinetID") = "CAB-" & ((i Mod 20) + 1).ToString("D2")
                    r("ShelfNo") = "S-" & ((i Mod 8) + 1).ToString()
                    r("BoxCode") = "BOX-" & ((i Mod 50) + 1).ToString("D2")
                    r("GDriveURL") = ""
                    r("CurrentStatus") = statuses(i Mod statuses.Length)
                    r("AssignedStaff") = "Officer " & ((i Mod 12) + 1).ToString()
                    r("DateReceived") = baseTime.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:ss")
                    r("FlowDirection") = "INCOMING"
                    r("AssignedSection") = "Secretariat"
                    r("TargetDeadlineUTC") = baseTime.AddDays(3).ToString("yyyy-MM-dd HH:mm:ss")
                    r("RevisionPunchlist") = ""
                    r("LastActionTaken") = "Registered and routed to Secretariat"
                    r("ExternalControlNumber") = ""
                    r("RequesterGender") = ""
                    r("CreatedByUserID") = 1
                    r("ModifiedByUserID") = 1
                    r("IsNewOfflineRecord") = False
                    r("PendingSync") = False
                    r("RowVersion") = rowVersion
                    docs.Rows.Add(r)
                Next
            Finally
                docs.EndLoadData()
            End Try

            directives.BeginLoadData()
            Try
                For i = 1 To n \ 2
                    Dim r = directives.NewRow()
                    r("DocumentID") = ((i * 2 - 1) Mod n) + 1
                    r("SGDirective") = "For Immediate Action"
                    r("AssignedTo") = "Records Section"
                    r("Notes") = "Benchmark directive"
                    r("LogUser") = "Secretary-General"
                    r("Timestamp") = baseTime.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:ss")
                    r("DirectiveCode") = "IMMEDIATE_ACTION"
                    r("IssuedByUserID") = 1
                    r("PendingSync") = False
                    directives.Rows.Add(r)
                Next
            Finally
                directives.EndLoadData()
            End Try

            routing.BeginLoadData()
            Try
                For i = 1 To n * 2
                    Dim r = routing.NewRow()
                    r("DocumentID") = ((i - 1) Mod n) + 1
                    r("FromOffice") = "Records Section"
                    r("ToOffice") = "Secretariat"
                    r("RoutedBy") = "Officer " & ((i Mod 12) + 1).ToString()
                    r("ActionTaken") = "ROUTED"
                    r("Remarks") = "Benchmark route"
                    r("Timestamp") = baseTime.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:ss")
                    r("RoutedByUserID") = 1
                    r("PendingSync") = False
                    routing.Rows.Add(r)
                Next
            Finally
                routing.EndLoadData()
            End Try

            movements.BeginLoadData()
            Try
                For i = 1 To n
                    Dim r = movements.NewRow()
                    r("DocumentID") = i
                    r("FromLocation") = "CAB-01|S-1|BOX-01"
                    r("ToLocation") = "CAB-02|S-2|BOX-05"
                    r("MovedBy") = "Officer " & ((i Mod 12) + 1).ToString()
                    r("Reason") = "Benchmark transfer"
                    r("Timestamp") = baseTime.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:ss")
                    r("MovedByUserID") = 1
                    r("PendingSync") = False
                    movements.Rows.Add(r)
                Next
            Finally
                movements.EndLoadData()
            End Try

            audit.BeginLoadData()
            Try
                For i = 1 To Math.Min(n, AuditRows)
                    Dim r = audit.NewRow()
                    r("UserName") = "Officer " & ((i Mod 12) + 1).ToString()
                    r("ActionDescription") = "Benchmark audit event"
                    r("Timestamp") = baseTime.AddMinutes(i).ToString("yyyy-MM-dd HH:mm:ss")
                    r("ActionType") = "ROUTING_LOG_ADDED"
                    r("UserID") = 1
                    r("PendingSync") = False
                    audit.Rows.Add(r)
                Next
            Finally
                audit.EndLoadData()
            End Try

            Dim hb = heartbeat.NewRow()
            hb("MachineName") = "BENCHMARK"
            hb("LastSyncUTC") = baseTime.ToString("yyyy-MM-dd HH:mm:ss")
            hb("AppVersion") = "2.9"
            hb("PendingOutbox") = 0
            hb("LastError") = ""
            heartbeat.Rows.Add(hb)
        End Sub

        Private Function BuildSnapshot() As Dictionary(Of String, DataTable)
            Return New Dictionary(Of String, DataTable) From {
                {"Documents", EmbeddedDB.DataSet.Tables("Documents").Copy()},
                {"Directives", EmbeddedDB.DataSet.Tables("Directives").Copy()},
                {"RoutingLogs", EmbeddedDB.DataSet.Tables("RoutingLogs").Copy()},
                {"Movements", EmbeddedDB.DataSet.Tables("Movements").Copy()},
                {"AuditTrail", EmbeddedDB.DataSet.Tables("AuditTrail").Copy()},
                {"Heartbeat", EmbeddedDB.DataSet.Tables("Heartbeat").Copy()}
            }
        End Function

        Private Sub ApplyQuiet(snapshot As Dictionary(Of String, DataTable))
            EmbeddedDB.ApplySnapshots(snapshot, persist:=False)
        End Sub

        Private Function PrepareLogCopies() As Dictionary(Of String, DataTable)
            Dim copies As New Dictionary(Of String, DataTable)()
            For Each name In {"Directives", "RoutingLogs", "Movements", "AuditTrail"}
                copies(name) = EmbeddedDB.DataSet.Tables(name).Copy()
            Next
            Return copies
        End Function

        Private Function Median(iterations As Integer, action As Action) As Double
            Dim times As New List(Of Double)()
            For i = 1 To iterations
                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()
                Dim sw = Stopwatch.StartNew()
                action()
                sw.Stop()
                times.Add(sw.Elapsed.TotalMilliseconds)
            Next
            times.Sort()
            Return times(times.Count \ 2)
        End Function

        Private Function SizeMb(path As String) As Double
            If Not File.Exists(path) Then Return 0
            Return New FileInfo(path).Length / 1024.0 / 1024.0
        End Function

        Private Function Fmt(value As Double) As String
            Return value.ToString("N2")
        End Function

    End Module
End Namespace
