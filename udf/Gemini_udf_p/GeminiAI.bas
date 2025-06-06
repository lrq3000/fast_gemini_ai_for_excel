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
        prompt     As String, _
        api_key    As String, _
        Optional model As String = "gemini-2.5-flash-preview-05-20", _
        Optional word_count As Long = 0) As Variant
    
    If Len(api_key) = 0 Then
        Gemini_udf_p = "Error: API key missing"
        Exit Function
    End If
    
    If gGeminiRequests Is Nothing Then Set gGeminiRequests = New Collection
    
    Dim req As New cGeminiRequest
    req.Launch prompt, api_key, model, word_count, Application.Caller
    gGeminiRequests.Add req
    
    Gemini_udf_p = "Pending..."              'temporary placeholder
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
            req.CommitResult                'write final answer into the cell
            gGeminiRequests.Remove i
        End If
    Next i
    
    Application.EnableEvents = True
    
    If gGeminiRequests.Count > 0 Then
        ScheduleNextPoll                    'still have requests → keep polling
    Else
        gTimerScheduled = False             'no more requests → stop polling
    End If
End Sub

'==============================================================
'  JSON HELPERS  (unchanged from before)
'==============================================================
Public Function ExtractContent(jsonString As String) As String
    Dim p1&, p2&, s$
    p1 = InStr(jsonString, """text"": """) + 9
    If p1 < 10 Then ExtractContent = "Error: parse failure": Exit Function
    p2 = InStr(p1, jsonString, """")
    s = Mid$(jsonString, p1, p2 - p1)
    s = Replace(s, "\""", """")
    s = Replace(s, "\n", vbLf)
    If Left$(Trim$(s), 1) = "=" Then s = "'" & s
    ExtractContent = Trim$(s)
End Function

Public Function ExtractError(jsonString As String) As String
    Dim p1&, p2&, s$
    p1 = InStr(jsonString, """message"": """) + 12
    If p1 < 13 Then ExtractError = "unknown error": Exit Function
    p2 = InStr(p1, jsonString, """")
    s = Mid$(jsonString, p1, p2 - p1)
    s = Replace(s, "\""", """")
    ExtractError = Trim$(s)
End Function
