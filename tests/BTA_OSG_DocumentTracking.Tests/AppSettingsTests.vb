Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text.Json
Imports System.Windows.Forms
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports BTA_OSG

Namespace BTA_OSG.Tests
    <TestClass>
    Public Class AppSettingsTests
        <TestMethod>
        Public Sub AppSettings_MalformedLayer_DoesNotDiscardLaterLayers()
            ' A stray character in one shipped layer used to abort the whole load, so a seat lost
            ' the connection it was configured with and fell back to the built-in default server.
            Dim resources = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources")
            Dim firstPath = Path.Combine(resources, "appsettings.json")
            Dim lastPath = Path.Combine(resources, "appsettings.local.json")
            Dim firstExisting = If(File.Exists(firstPath), File.ReadAllText(firstPath), Nothing)
            Dim lastExisting = If(File.Exists(lastPath), File.ReadAllText(lastPath), Nothing)
            Try
                Directory.CreateDirectory(resources)
                File.WriteAllText(firstPath, "{ ""Database"": { ""ConnectionString"": ""unterminated")
                File.WriteAllText(lastPath, "{ ""Database"": { ""ConnectionString"": ""Server=SENTINEL-SEAT;Database=BTA_OSG_DB"" } }")

                AppSettings.Reload()

                Assert.IsTrue(AppSettings.Instance.IgnoredConfigLayers.Any(Function(s) s.Contains("appsettings.json")),
                              "the layer that failed to parse must be named")
                StringAssert.Contains(AppSettings.Instance.DatabaseSettings.ConnectionString, "SENTINEL-SEAT",
                                      "a valid later layer must still apply")
            Finally
                Restore(firstPath, firstExisting)
                Restore(lastPath, lastExisting)
                AppSettings.Reload()
            End Try
        End Sub

        Private Shared Sub Restore(path As String, content As String)
            If content Is Nothing Then
                If File.Exists(path) Then File.Delete(path)
            Else
                File.WriteAllText(path, content)
            End If
        End Sub

        <TestMethod>
        Public Sub AppSettingsRoot_Serialization_PreservesIsConfiguredFlag()
            Dim root As New AppSettingsRoot With {
                .IsConfigured = True,
                .Database = New DatabaseSettings With {
                    .ConnectionString = "Server=TEST-SQL;Database=BTA_OSG_DB;Integrated Security=true;",
                    .UseSqlServer = True
                },
                .Rfid = New RfidSettings With {
                    .LockoutThreshold = 5
                },
                .Session = New SessionSettings With {
                    .TimeoutMinutes = 20
                },
                .Portal = New PortalSettings With {
                    .PortalEnabled = True,
                    .BaseUrl = "http://localhost:8085"
                }
            }

            Dim options As New JsonSerializerOptions With {.WriteIndented = True}
            Dim json = JsonSerializer.Serialize(root, options)
            Assert.IsTrue(json.Contains("""IsConfigured"": true"), "Serialized JSON must include IsConfigured flag")

            Dim deserialized = JsonSerializer.Deserialize(Of AppSettingsRoot)(json, New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True})
            Assert.IsNotNull(deserialized)
            Assert.IsTrue(deserialized.IsConfigured, "Deserialized root must retain IsConfigured = True")
            Assert.AreEqual("Server=TEST-SQL;Database=BTA_OSG_DB;Integrated Security=true;", deserialized.Database.ConnectionString)
            Assert.AreEqual(20, deserialized.Session.TimeoutMinutes)
        End Sub

        <TestMethod>
        Public Sub AppSettingsRoot_UnconfiguredFlag_DeserializesToFalse()
            Dim json = "{""Database"": {""UseSqlServer"": false}}"
            Dim deserialized = JsonSerializer.Deserialize(Of AppSettingsRoot)(json, New JsonSerializerOptions With {.PropertyNameCaseInsensitive = True})
            Assert.IsNotNull(deserialized)
            Assert.IsFalse(deserialized.IsConfigured, "Root without IsConfigured property should default to False")
        End Sub

        <TestMethod>
        Public Sub AppSettings_Instance_IsAccessibleAndPopulated()
            Dim instance = AppSettings.Instance
            Assert.IsNotNull(instance, "AppSettings.Instance must not be null")
            Assert.IsNotNull(instance.DatabaseSettings, "DatabaseSettings must not be null")
            Assert.IsNotNull(instance.RfidSettings, "RfidSettings must not be null")
            Assert.IsNotNull(instance.PortalSettings, "PortalSettings must not be null")
        End Sub

        <TestMethod>
        Public Sub FormConnect_BuildDataSource_NamedInstanceNeverGetsPortSuffix()
            Assert.AreEqual("DESKTOP-SCMBIEM", FormConnect.BuildDataSource("DESKTOP-SCMBIEM", ""))
            Assert.AreEqual("DESKTOP-SCMBIEM,1433", FormConnect.BuildDataSource("DESKTOP-SCMBIEM", "1433"))
            ' A named instance resolves its own port through the SQL Browser service, so
            ' a stale port prefill must never be appended to it.
            Assert.AreEqual("OFFICE-PC\SQLEXPRESS", FormConnect.BuildDataSource("OFFICE-PC\SQLEXPRESS", "1433"))
        End Sub

        <TestMethod>
        Public Sub FormClaimAdmin_TakesNameCardAndDesk()
            ' The keyboard-wedge reader types the UID and ends with Enter, so the accept
            ' button has to be the claim itself or a tapped badge does nothing.
            Using dlg As New FormClaimAdmin()
                Dim boxes As New List(Of TextBox)()
                FindControls(dlg, boxes)
                Assert.AreEqual(3, boxes.Count, "full name, card UID, and desk are the three fields a claim needs")

                Dim buttons As New List(Of Button)()
                FindControls(dlg, buttons)
                Dim claim = buttons.Find(Function(b) b.Text.Contains("Claim"))
                Assert.IsNotNull(claim, "the claim button must exist")
                Assert.AreSame(dlg.AcceptButton, claim, "Enter submits the claim")
            End Using
        End Sub

        <TestMethod>
        Public Sub FormLogin_ListsEnrolledBadges()
            ' The list is the only visible credential input on the login window, so an enrolled
            ' officer has to appear in it with no setting to switch it on.
            EmbeddedDB.Initialize()
            EmbeddedDB.AddUser("LISTING001", "Listing Probe Officer", "Records Section", "Records Section")

            Using dlg As New FormLogin()
                Dim boxes As New List(Of TextBox)()
                FindControls(dlg, boxes)
                Assert.AreEqual(0, boxes.Count, "the login window takes no typed card ID")

                Dim combos As New List(Of ComboBox)()
                FindControls(dlg, combos)
                Assert.AreEqual(1, combos.Count, "the login window offers one badge list")
                Dim captions = combos(0).Items.Cast(Of Object)().Select(Function(item) item.ToString()).ToList()
                Assert.IsTrue(captions.Any(Function(c) c.Contains("Listing Probe Officer")), "an enrolled officer is listed")
            End Using
        End Sub

        <TestMethod>
        Public Sub FormConnect_PresentsConnectFirstSetup()
            Using dlg As New FormConnect()
                Dim links As New List(Of LinkLabel)()
                FindControls(dlg, links)
                Assert.IsNotNull(links.Find(Function(l) l.Text.Contains("demo")), "the demo fallback link must exist")
                Assert.IsNotNull(links.Find(Function(l) l.Text.Contains("other settings")), "the options link must exist")
                Dim masks As New List(Of TextBox)()
                FindControls(dlg, masks, Function(t) t.UseSystemPasswordChar)
                Assert.AreEqual(1, masks.Count, "exactly one masked shared-password box must exist")
            End Using
        End Sub

        Private Shared Sub FindControls(Of T As Control)(parent As Control, result As List(Of T), Optional match As Func(Of T, Boolean) = Nothing)
            For Each c As Control In parent.Controls
                Dim typed = TryCast(c, T)
                If typed IsNot Nothing AndAlso (match Is Nothing OrElse match(typed)) Then
                    result.Add(typed)
                End If
                If c.HasChildren Then
                    FindControls(c, result, match)
                End If
            Next
        End Sub
    End Class
End Namespace
