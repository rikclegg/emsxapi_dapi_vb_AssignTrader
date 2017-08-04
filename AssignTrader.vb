' Copyright 2017. Bloomberg Finance L.P.
'
' Permission Is hereby granted, free of charge, to any person obtaining a copy
' of this software And associated documentation files (the "Software"), to
' deal in the Software without restriction, including without limitation the
' rights to use, copy, modify, merge, publish, distribute, sublicense, And/Or
' sell copies of the Software, And to permit persons to whom the Software Is
' furnished to do so, subject to the following conditions:  The above
' copyright notice And this permission notice shall be included in all copies
' Or substantial portions of the Software.
'
' THE SOFTWARE Is PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS Or
' IMPLIED, INCLUDING BUT Not LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE And NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS Or COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES Or OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT Or OTHERWISE, ARISING
' FROM, OUT OF Or IN CONNECTION WITH THE SOFTWARE Or THE USE Or OTHER DEALINGS
' IN THE SOFTWARE.
'

Imports Bloomberglp.Blpapi

Namespace com.bloomberg.emsx.samples
    Module AssignTrader

        Private ReadOnly SESSION_STARTED As New Name("SessionStarted")
        Private ReadOnly SESSION_STARTUP_FAILURE As New Name("SessionStartupFailure")
        Private ReadOnly SERVICE_OPENED As New Name("ServiceOpened")
        Private ReadOnly SERVICE_OPEN_FAILURE As New Name("ServiceOpenFailure")
        Private ReadOnly ERROR_INFO As New Name("ErrorInfo")
        Private ReadOnly ASSIGN_TRADER As New Name("AssignTrader")

        Private d_service As String
        Private d_host As String
        Private d_port As Integer

        Private quit As Boolean = False

        Private requestID As CorrelationID

        Sub Main(ByVal args As String())

            System.Console.WriteLine("Bloomberg - EMSX API Example - AssignTrader")

            Dim example As AssignTrader = New AssignTrader()
            example.run()

            Do
            Loop Until quit

            'System.Console.ReadLine()

        End Sub

        Class AssignTrader

            Sub New()
                d_service = "//blp/emapisvc_beta"
                d_host = "localhost"
                d_port = 8194
            End Sub

            Friend Sub run()

                Dim d_sessionOptions As SessionOptions
                Dim session As Session

                d_sessionOptions = New SessionOptions()
                d_sessionOptions.ServerHost = d_host
                d_sessionOptions.ServerPort = d_port

                session = New Session(d_sessionOptions, New EventHandler(AddressOf processEvent))
                session.StartAsync()

            End Sub

            Public Sub processEvent(ByVal eventObj As [Event], ByVal session As Session)

                Try
                    Select Case eventObj.Type

                        Case [Event].EventType.SESSION_STATUS
                            processSessionEvent(eventObj, session)
                            Exit Select
                        Case [Event].EventType.SERVICE_STATUS
                            processServiceEvent(eventObj, session)
                            Exit Select
                        Case [Event].EventType.RESPONSE
                            processResponseEvent(eventObj, session)
                            Exit Select
                        Case Else
                            processMiscEvent(eventObj, session)
                            Exit Select
                    End Select
                Catch ex As Exception
                    System.Console.Error.WriteLine(ex)
                End Try

            End Sub

            Private Sub processSessionEvent(eventObj As [Event], session As Session)

                System.Console.WriteLine("Processing " + eventObj.Type.ToString)

                For Each msg As Message In eventObj
                    If msg.MessageType.Equals(SESSION_STARTED) Then
                        System.Console.WriteLine("Session started...")
                        session.OpenServiceAsync(d_service)
                    ElseIf msg.MessageType.Equals(SESSION_STARTUP_FAILURE) Then
                        System.Console.Error.WriteLine("Error: Session startup failed")
                    End If
                Next msg

            End Sub

            Private Sub processServiceEvent(eventObj As [Event], session As Session)

                System.Console.WriteLine("Processing " + eventObj.Type.ToString)

                For Each msg As Message In eventObj
                    If msg.MessageType.Equals(SERVICE_OPENED) Then
                        System.Console.WriteLine("Service opened...")

                        Dim service As Service = session.GetService(d_service)
                        Dim request As Request = service.CreateRequest("AssignTrader")

                        'request.Set("EMSX_REQUEST_SEQ", 1)

                        ' The fields below are mandatory
                        request.Append("EMSX_SEQUENCE", 3998997)

                        request.Set("EMSX_ASSIGNEE_TRADER_UUID", 12109783)

                        System.Console.WriteLine("Request: " + request.ToString)

                        requestID = New CorrelationID()

                        ' Submit the request
                        Try
                            session.SendRequest(request, requestID)
                        Catch ex As Exception
                            System.Console.Error.WriteLine("Failed to send the request: " + ex.Message)
                        End Try

                    ElseIf msg.MessageType.Equals(SERVICE_OPEN_FAILURE) Then
                        System.Console.Error.WriteLine("Error: Service failed to open")
                    End If

                Next msg

            End Sub

            Private Sub processResponseEvent(eventObj As [Event], session As Session)

                System.Console.WriteLine("Processing " + eventObj.Type.ToString)

                For Each msg As Message In eventObj

                    System.Console.WriteLine("MESSAGE: " + msg.ToString)
                    System.Console.WriteLine("CORRELATION ID: " + msg.CorrelationID.ToString)

                    If msg.CorrelationID Is requestID Then

                        System.Console.WriteLine("Message Type: " + msg.MessageType.ToString)

                        If msg.MessageType.Equals(ERROR_INFO) Then
                            Dim errorCode As Integer = msg.GetElementAsInt32("ERROR_CODE")
                            Dim errorMessage As String = msg.GetElementAsString("ERROR_MESSAGE")
                            System.Console.WriteLine("ERROR CODE: " + errorCode.ToString + Chr(9) + "ERROR MESSAGE: " + errorMessage)
                        ElseIf msg.MessageType.Equals(ASSIGN_TRADER) Then

                            Dim success As Boolean = msg.GetElementAsBool("EMSX_ALL_SUCCESS")

                            If success Then

                                System.Console.WriteLine("All orders successfully assigned")

                                Dim successful As Element = msg.GetElement("EMSX_ASSIGN_TRADER_SUCCESSFUL_ORDERS")
                                Dim numValues As Integer = successful.NumValues

                                If numValues > 0 Then System.Console.WriteLine("Successful assignments:-")

                                For i As Integer = 0 To numValues - 1
                                    Dim order As Element = successful.GetValueAsElement(i)
                                    System.Console.WriteLine(order.GetElement("EMSX_SEQUENCE").GetValueAsInt32().ToString)
                                Next i
                            Else
                                System.Console.WriteLine("One or more failed assignments...\n")

                                If msg.HasElement("EMSX_ASSIGN_TRADER_SUCCESSFUL_ORDERS") Then
                                    Dim successful As Element = msg.GetElement("EMSX_ASSIGN_TRADER_SUCCESSFUL_ORDERS")
                                    Dim numValues As Integer = successful.NumValues

                                    If numValues > 0 Then System.Console.WriteLine("Successful assignments:-")

                                    For i As Integer = 0 To numValues - 1
                                        Dim order As Element = successful.GetValueAsElement(i)
                                        System.Console.WriteLine(order.GetElement("EMSX_SEQUENCE").GetValueAsInt32().ToString)
                                    Next i
                                End If

                                If msg.HasElement("EMSX_ASSIGN_TRADER_FAILED_ORDERS") Then

                                    Dim failed As Element = msg.GetElement("EMSX_ASSIGN_TRADER_FAILED_ORDERS")

                                    Dim numValues As Integer = failed.NumValues

                                    If numValues > 0 Then System.Console.WriteLine("Failed assignments:-")

                                    For i As Integer = 0 To numValues - 1

                                        Dim order As Element = failed.GetValueAsElement(i)
                                        System.Console.WriteLine(order.GetElement("EMSX_SEQUENCE").GetValueAsInt32().ToString)
                                    Next i
                                End If
                            End If
                        End If

                        quit = True
                        session.Stop()
                    End If
                Next
            End Sub

            Private Sub processMiscEvent(eventObj As [Event], session As Session)

                System.Console.WriteLine("Processing " + eventObj.Type.ToString)

                For Each msg As Message In eventObj
                    System.Console.WriteLine("MESSAGE: " + msg.ToString)
                Next msg

            End Sub

        End Class

    End Module

End Namespace
