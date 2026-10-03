Option Explicit On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Text
Imports System.Text.Json
Imports System.Threading.Tasks
Imports System.Linq

Namespace BTA_OSG
    ''' <summary>
    ''' HTTP bridge client communicating with the external public intake portal.
    ''' Conforms to the isolation invariant: No portal outage or network failure
    ''' may ever crash, freeze, or block internal office document workflows.
    ''' </summary>
    Public Class PortalBridge
        Implements IPortalBridge

        Private ReadOnly _httpClient As HttpClient
        Private ReadOnly _settings As PortalSettings
        Private ReadOnly _auditService As AuditService

        Public Sub New(Optional settings As PortalSettings = Nothing, Optional httpClient As HttpClient = Nothing, Optional auditService As AuditService = Nothing)
            _settings = If(settings, AppSettings.Instance.PortalSettings)
            If _settings Is Nothing Then
                _settings = New PortalSettings()
            End If

            If httpClient IsNot Nothing Then
                _httpClient = httpClient
            Else
                _httpClient = New HttpClient()
                _httpClient.Timeout = TimeSpan.FromSeconds(Math.Max(3, _settings.TimeoutSeconds))
            End If

            _auditService = auditService
        End Sub

        Public Async Function FetchExternalQueueAsync() As Task(Of List(Of PortalSubmission)) Implements IPortalBridge.FetchExternalQueueAsync
            Dim list As New List(Of PortalSubmission)()

            If Not _settings.PortalEnabled OrElse String.IsNullOrWhiteSpace(_settings.BaseUrl) Then
                Return list
            End If

            Try
                Dim requestUrl = $"{_settings.BaseUrl.TrimEnd("/"c)}/api/external_queue.php"
                Using request = New HttpRequestMessage(HttpMethod.Get, requestUrl)
                    AttachHeaders(request)

                    Using response = Await _httpClient.SendAsync(request).ConfigureAwait(False)
                        If Not response.IsSuccessStatusCode Then
                            LogWarning($"Portal queue query failed with HTTP {CInt(response.StatusCode)}.")
                            Return list
                        End If

                        Dim json = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                        Using doc = JsonDocument.Parse(json)
                            Dim root = doc.RootElement
                            If root.TryGetProperty("success", Nothing) AndAlso root.GetProperty("success").GetBoolean() Then
                                If root.TryGetProperty("items", Nothing) Then
                                    For Each item In root.GetProperty("items").EnumerateArray()
                                        Dim subItem As New PortalSubmission()

                                        If item.TryGetProperty("control_number", Nothing) Then
                                            subItem.ControlNumber = item.GetProperty("control_number").GetString()
                                        End If
                                        If item.TryGetProperty("category", Nothing) Then
                                            subItem.Category = item.GetProperty("category").GetString()
                                        End If
                                        If item.TryGetProperty("document_title", Nothing) Then
                                            subItem.DocumentTitle = item.GetProperty("document_title").GetString()
                                        End If
                                        If item.TryGetProperty("requester_name", Nothing) Then
                                            subItem.RequesterName = item.GetProperty("requester_name").GetString()
                                        End If
                                        If item.TryGetProperty("requester_email", Nothing) Then
                                            subItem.RequesterEmail = item.GetProperty("requester_email").GetString()
                                        End If
                                        If item.TryGetProperty("requester_phone", Nothing) Then
                                            subItem.RequesterPhone = item.GetProperty("requester_phone").GetString()
                                        End If
                                        If item.TryGetProperty("requester_gender", Nothing) Then
                                            subItem.RequesterGender = item.GetProperty("requester_gender").GetString()
                                        End If
                                        If item.TryGetProperty("created_at", Nothing) Then
                                            Dim dateStr = item.GetProperty("created_at").GetString()
                                            Dim parsedDate As DateTime
                                            If DateTime.TryParse(dateStr, parsedDate) Then
                                                subItem.CreatedAt = parsedDate
                                            Else
                                                subItem.CreatedAt = DateTime.UtcNow
                                            End If
                                        End If

                                        list.Add(subItem)
                                    Next
                                End If
                            End If
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                LogWarning($"Exception fetching external queue from portal: {ex.Message}")
            End Try

            Return list
        End Function

        Public Async Function AcknowledgeImportAsync(controlNumber As String) As Task(Of Boolean) Implements IPortalBridge.AcknowledgeImportAsync
            If Not _settings.PortalEnabled OrElse String.IsNullOrWhiteSpace(_settings.BaseUrl) Then
                Return True
            End If

            If String.IsNullOrWhiteSpace(controlNumber) Then
                Return False
            End If

            Try
                Dim requestUrl = $"{_settings.BaseUrl.TrimEnd("/"c)}/api/import_ack.php"
                Dim payload = JsonSerializer.Serialize(New Dictionary(Of String, String) From {
                    {"control_number", controlNumber}
                })

                Using request = New HttpRequestMessage(HttpMethod.Post, requestUrl)
                    AttachHeaders(request)
                    request.Content = New StringContent(payload, Encoding.UTF8, "application/json")

                    Using response = Await _httpClient.SendAsync(request).ConfigureAwait(False)
                        If Not response.IsSuccessStatusCode Then
                            LogWarning($"Import ack failed for {controlNumber} with HTTP {CInt(response.StatusCode)}.")
                            Return False
                        End If

                        Dim json = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                        Using doc = JsonDocument.Parse(json)
                            Dim root = doc.RootElement
                            Return root.TryGetProperty("success", Nothing) AndAlso root.GetProperty("success").GetBoolean()
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                LogWarning($"Exception acknowledging import for {controlNumber}: {ex.Message}")
                Return False
            End Try
        End Function

        Public Async Function AcknowledgeBatchImportAsync(controlNumbers As IEnumerable(Of String)) As Task(Of Boolean) Implements IPortalBridge.AcknowledgeBatchImportAsync
            If Not _settings.PortalEnabled OrElse String.IsNullOrWhiteSpace(_settings.BaseUrl) Then
                Return True
            End If

            If controlNumbers Is Nothing OrElse Not controlNumbers.Any() Then
                Return True
            End If

            Try
                Dim requestUrl = $"{_settings.BaseUrl.TrimEnd("/"c)}/api/import_ack.php"
                Dim payload = JsonSerializer.Serialize(New Dictionary(Of String, Object) From {
                    {"control_numbers", controlNumbers.ToList()}
                })

                Using request = New HttpRequestMessage(HttpMethod.Post, requestUrl)
                    AttachHeaders(request)
                    request.Content = New StringContent(payload, Encoding.UTF8, "application/json")

                    Using response = Await _httpClient.SendAsync(request).ConfigureAwait(False)
                        If Not response.IsSuccessStatusCode Then
                            LogWarning($"Batch import ack failed with HTTP {CInt(response.StatusCode)}.")
                            Return False
                        End If

                        Dim json = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                        Using doc = JsonDocument.Parse(json)
                            Dim root = doc.RootElement
                            Return root.TryGetProperty("success", Nothing) AndAlso root.GetProperty("success").GetBoolean()
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                LogWarning($"Exception acknowledging batch import: {ex.Message}")
                Return False
            End Try
        End Function

        Public Async Function PushPublicStatusAsync(controlNumber As String, publicStatus As String) As Task(Of Boolean) Implements IPortalBridge.PushPublicStatusAsync
            If Not _settings.PortalEnabled OrElse String.IsNullOrWhiteSpace(_settings.BaseUrl) Then
                Return True
            End If

            If String.IsNullOrWhiteSpace(controlNumber) OrElse String.IsNullOrWhiteSpace(publicStatus) Then
                Return False
            End If

            Try
                Dim requestUrl = $"{_settings.BaseUrl.TrimEnd("/"c)}/api/status.php"
                Dim payload = JsonSerializer.Serialize(New Dictionary(Of String, String) From {
                    {"control_number", controlNumber},
                    {"public_status", publicStatus}
                })

                Using request = New HttpRequestMessage(HttpMethod.Post, requestUrl)
                    AttachHeaders(request)
                    request.Content = New StringContent(payload, Encoding.UTF8, "application/json")

                    Using response = Await _httpClient.SendAsync(request).ConfigureAwait(False)
                        If Not response.IsSuccessStatusCode Then
                            LogWarning($"Status push failed for {controlNumber} ({publicStatus}) with HTTP {CInt(response.StatusCode)}.")
                            Return False
                        End If

                        Dim json = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)
                        Using doc = JsonDocument.Parse(json)
                            Dim root = doc.RootElement
                            Return root.TryGetProperty("success", Nothing) AndAlso root.GetProperty("success").GetBoolean()
                        End Using
                    End Using
                End Using
            Catch ex As Exception
                LogWarning($"Exception pushing public status for {controlNumber}: {ex.Message}")
                Return False
            End Try
        End Function

        Private Sub AttachHeaders(request As HttpRequestMessage)
            request.Headers.Accept.Clear()
            request.Headers.Accept.Add(New MediaTypeWithQualityHeaderValue("application/json"))

            If Not String.IsNullOrEmpty(_settings.BridgeKey) Then
                request.Headers.Add("X-Bridge-Key", _settings.BridgeKey)
            End If

            request.Headers.TryAddWithoutValidation("User-Agent", "BTA-OSG-PortalBridge/2.1")
        End Sub

        Private Sub LogWarning(message As String)
            Trace.TraceWarning($"[PortalBridge] {message}")
            If _auditService IsNot Nothing Then
                Try
                    _auditService.LogEvent("PORTAL_BRIDGE_WARNING", "PortalBridge", Nothing, Nothing, Nothing, Nothing, False, message)
                Catch
                    ' Never throw from audit logging
                End Try
            End If
        End Sub
    End Class
End Namespace
