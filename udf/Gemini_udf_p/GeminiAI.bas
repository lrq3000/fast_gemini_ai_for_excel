Attribute VB_Name = "GeminiAI"
Option Explicit

'==============================================================
'  GLOBAL STATE
'==============================================================
Public gGeminiRequests  As Collection      'alive asynchronous requests
Private gTimerScheduled As Boolean         'true while a poll is queued

'--------------------------------------------------------------
'  MANUAL START FOR THE POLLING LOOP
'--------------------------------------------------------------
' Call this once (e.g. from the Macros dialog or a button) 
' after you open the workbook and click "Enable Content".
Public Sub StartGeminiPoller()
    If Not gTimerScheduled Then
        ScheduleNextPoll
    End If
End Sub

'==============================================================
'  PUBLIC UDF — called from the worksheet
'==============================================================
Public Function Gemini_udf_p( _
        prompt As String, _
        api_key As String, _
        Optional model As String = "gemini-2.5-flash-preview-05-20", _
        Optional word_count As Long = 0, _
        Optional maxDelayMs As Long = 500, _
        Optional retries As Integer = 2, _
        Optional server_url As String = "", _
        Optional ByVal asynchronous As Boolean = True) As Variant
    ' This function allows users to interact with the Gemini API (or an OpenAI-compatible API)
    ' from an Excel worksheet. It supports both asynchronous (default) and synchronous execution modes.
    '
    ' Arguments:
    '   prompt (String): The user's query or prompt for the AI model.
    '   api_key (String): Your API key for accessing the Gemini API or the OpenAI-compatible API.
    '   model (String, Optional): The name of the AI model to use (default: "gemini-2.5-flash-preview-05-20").
    '   word_count (Long, Optional): The desired maximum word count for the AI's response (default: 0, no limit).
    '   maxDelayMs (Long, Optional): Maximum delay in milliseconds for retries in asynchronous mode (default: 500).
    '   retries (Integer, Optional): Number of retries for failed requests in asynchronous mode (default: 2).
    '   server_url (String, Optional): Custom API endpoint URL for OpenAI-compatible servers (default: "").
    '   asynchronous (Boolean, Optional): If False, the function executes synchronously, waiting for the
    '                                    AI response before returning. If True (default), it executes
    '                                    asynchronously, returning "Pending..." immediately and updating
    '                                    the cell later.

    If Len(api_key) = 0 Then
        Gemini_udf_p = "Error: API key missing"
        Exit Function
    End If

    Dim req As New cGeminiRequest

    If asynchronous Then
        ' Asynchronous execution (default behavior): Add the request to the global collection
        ' and return "Pending..." immediately. The result will be committed later by the poller.
        If gGeminiRequests Is Nothing Then Set gGeminiRequests = New Collection
        req.Launch prompt, api_key, model, word_count, Application.Caller, retries, maxDelayMs, server_url, asynchronous
        gGeminiRequests.Add req
        Gemini_udf_p = "Pending..." ' Temporary placeholder
    Else
        ' Synchronous execution: Wait for the response before returning.
        ' The request is not added to the global collection for polling.
        req.Launch prompt, api_key, model, word_count, Application.Caller, retries, maxDelayMs, server_url, asynchronous
        
        ' Wait for the request to complete
        Do While Not req.IsDone
            DoEvents ' Allow other processes to run while waiting
        Loop
        
        ' In synchronous mode, the UDF must return its value directly.
        ' It is crucial to avoid writing to Application.Caller.Value (which req.CommitResult would do)
        ' or reading from Application.Caller.Value within a synchronous UDF's calculation,
        ' as this creates a circular reference in Excel.
        ' Therefore, req.CommitResult is not called here, and Application.Caller.Value is not accessed.
        ' In synchronous mode, the UDF must return its value directly.
        ' It is crucial to avoid writing to Application.Caller.Value (which req.CommitResult would do)
        ' or reading from Application.Caller.Value within a synchronous UDF's calculation,
        ' as this creates a circular reference in Excel.
        ' Therefore, req.CommitResult is not called here, and Application.Caller.Value is not accessed.
        ' Instead, the UDF's return value is set directly from the request's raw response text.
        ' This resolves the circular reference issue and provides the correct synchronous return.
        ' Use the new ExtractedError and ExtractedText properties for robust content and error extraction.
        If req.ExtractedError <> "No specific error message found in JSON response." Then
            Gemini_udf_p = "Error: " & req.ExtractedError
        Else
            Gemini_udf_p = req.ExtractedText
        End If
    End If
End Function

'--------------------------------------------------------------
'  POLLING LOOP
'--------------------------------------------------------------
Private Sub ScheduleNextPoll()
    Application.OnTime Now + TimeSerial(0, 0, 1), "PollAsyncRequests"
    gTimerScheduled = True
End Sub

Public Sub PollAsyncRequests()
    DoEvents                                'allow MSXML to advance readyState
    On Error Resume Next
    Application.EnableEvents = False

    Dim i As Long
    For i = gGeminiRequests.Count To 1 Step -1
        Dim req As cGeminiRequest
        Set req = gGeminiRequests(i)

        If req.IsDone Then
            If req.NeedsRetry Then
                req.RetryRequest            'retry on failure
            Else
                req.CommitResult            'write final answer into the cell
                gGeminiRequests.Remove i
            End If
        End If
    Next i

    Application.EnableEvents = True

    If gGeminiRequests.Count > 0 Then
        ScheduleNextPoll                    'still have requests → keep polling
    Else
        gTimerScheduled = False             'no more requests → stop polling
    End If
End Sub

