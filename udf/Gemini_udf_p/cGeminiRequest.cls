VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cGeminiRequest"
Option Explicit

Private http As Object ' late-bound → no reference needed
Private m_ParsedJson As Object ' Stores the parsed JSON response
Private retryCount As Integer
Private maxRetries As Integer
Private maxDelayMs As Long
Public TargetCell As Range
Private Const OPENAI_COMPATIBLE_API_PATH As String = "/v1/chat/completions" ' Define the relative path for OpenAI-compatible chat completions

'------------------  SEND THE REQUEST  ------------------------
Public Sub Launch( _
        ByVal userPrompt As String, _
        ByVal key As String, _
        ByVal modelName As String, _
        ByVal word_count As Long, _
        ByVal tgt As Range, _
        Optional ByVal retries As Integer = 2, _
        Optional ByVal delayMs As Long = 500, _
        Optional ByVal server_url As String = "", _
        Optional ByVal asynchronous As Boolean = True)
    ' Initializes and sends an HTTP request to the Gemini API or an OpenAI-compatible API.
    '
    ' Arguments:
    '   userPrompt (String): The prompt text to send to the AI model.
    '   key (String): The API key for authentication.
    '   modelName (String): The name of the AI model to use.
    '   word_count (Long): The desired maximum word count for the response.
    '   tgt (Range): The Excel cell where the result will be committed.
    '   retries (Integer, Optional): The number of times to retry the request on failure (default: 2).
    '   delayMs (Long, Optional): The maximum delay in milliseconds before retrying (default: 500).
    '   server_url (String, Optional): A custom server URL for OpenAI-compatible APIs (default: "").

    Set TargetCell = tgt
    maxRetries = retries
    maxDelayMs = delayMs
    retryCount = 0

    Dim prompt$, url$, body$, api_key_header As String
    prompt = userPrompt
    If word_count > 0 Then _
        prompt = prompt & ". Provide response in maximum " & word_count & " words"
    ' Clean prompt for JSON formatting
    prompt = Replace(prompt, """", "'")
    prompt = Replace(prompt, vbLf, " ")

    ' Conditional logic to determine URL and request body based on server_url
    If server_url = "" Then
        ' Gemini API configuration
        url = "https://generativelanguage.googleapis.com/v1beta/models/" & _
              modelName & ":generateContent?key=" & key
        body = "{""contents"":{""parts"":[{""text"":""" & prompt & """}]}," & _
               """generationConfig"":{""temperature"":0.5}}"
        api_key_header = "" ' API key is part of the URL for Gemini
    Else
        ' OpenAI-compatible API configuration
        url = server_url & OPENAI_COMPATIBLE_API_PATH
        ' Construct OpenAI-compatible request body
        body = "{""model"":""" & modelName & """," & _
               """messages"":[{""role"":""user"",""content"":""" & prompt & """}]," & _
               """temperature"":0.0" ' Default temperature for OpenAI-compatible
        If word_count > 0 Then
            body = body & ",""max_tokens"":" & word_count * 5 ' Approximate max_tokens based on word_count (5 chars per word)
        End If
        body = body & "}"

        ' API key handling for OpenAI-compatible servers
        If Len(key) > 0 Then
            api_key_header = "Bearer " & key
        Else
            api_key_header = "" ' No API key for local Ollama instances or if not provided
        End If
    End If

    ' Craft the HTTP request
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, asynchronous                 ' True = async
    http.setRequestHeader "Content-Type", "application/json"
    ' Add Authorization header for OpenAI-compatible servers if an API key is provided
    If api_key_header <> "" Then
        http.setRequestHeader "Authorization", api_key_header
    End If
    http.Send body

    ' For synchronous requests, wait for completion and parse immediately
    If Not asynchronous Then
        Do While Not http.readyState = 4
            DoEvents ' Yield control to allow other events to process
        Loop
        ' Parse the JSON response immediately after synchronous request completes
        Call ParseHttpResponse
    End If
End Sub

'------------------  STATUS QUERIES  --------------------------
Public Property Get IsDone() As Boolean
    On Error Resume Next
    IsDone = (http.readyState = 4)
End Property

Public Sub CommitResult()
    ' Commits the result of the HTTP request to the target cell.
    ' This method handles successful responses by parsing JSON and extracting content,
    ' and handles errors by extracting error messages and providing retry information.

    Dim resultText As String

    ' Call the shared JSON parsing logic.
    ' This ensures m_ParsedJson is populated for asynchronous requests,
    ' or set to Nothing if parsing fails.
    Call ParseHttpResponse

    ' Check if the HTTP request was successful (HTTP Status 200 OK)
    If http.Status = 200 Then
        ' Extract the relevant text content from the parsed JSON.
        ' This leverages the ExtractedText property, which handles cases where m_ParsedJson is Nothing.
        resultText = Me.ExtractedText
    Else
        ' If the HTTP request failed, extract the error message from the response.
        ' This uses the ExtractedError property, which handles cases where m_ParsedJson is Nothing.
        resultText = "Error: " & Me.ExtractedError
        ' Append the original formula to allow the user to easily retry the request.
        resultText = resultText & vbLf & vbLf & "Retry with: " & TargetCell.Formula
    End If

    ' Temporarily disable events to prevent triggering other macros during cell update.
    Application.EnableEvents = False
    ' Commit the extracted result (either content or error message) to the target cell.
    TargetCell.Value = resultText
    ' Re-enable events.
    Application.EnableEvents = True
End Sub

Public Property Get NeedsRetry() As Boolean
    If http.Status = 200 Then
        NeedsRetry = False
    ElseIf retryCount < maxRetries Then
        NeedsRetry = True
    Else
        NeedsRetry = False
    End If
End Property

Private Sub ParseHttpResponse()
    ' Parses the HTTP response text as JSON and populates m_ParsedJson.
    ' This method is called for both synchronous (in Launch) and asynchronous (in CommitResult)
    ' requests to ensure consistent JSON parsing.
    On Error GoTo Err_ParseJson

    ' Only attempt to parse if the HTTP request was successful (HTTP Status 200 OK).
    ' This prevents attempting to parse non-JSON error responses as JSON.
    If http.Status = 200 Then
        Set m_ParsedJson = JsonConverter.ParseJson(http.responseText)
    Else
        ' If HTTP status is not 200, ensure m_ParsedJson is Nothing,
        ' as there's no valid JSON response to parse for content.
        Set m_ParsedJson = Nothing
    End If

    On Error GoTo 0 ' Disable error handling after successful parsing setup
    Exit Sub

Err_ParseJson:
    ' Handle JSON parsing errors.
    ' If parsing fails, set m_ParsedJson to Nothing to indicate no valid JSON was parsed.
    ' The ExtractedText and ExtractedError properties are designed to handle m_ParsedJson being Nothing.
    Set m_ParsedJson = Nothing
    ' Optionally, log the error for debugging purposes.
    Debug.Print "JSON parsing failed: " & Err.Description & " for response: " & http.responseText
    On Error GoTo 0 ' Ensure error handling is reset
End Sub

Public Sub RetryRequest()
    retryCount = retryCount + 1

    Dim delay As Single
    Randomize
    delay = Rnd() * maxDelayMs * retryCount / 1000

    Dim t As Single: t = Timer
    Do While Timer - t < delay
        DoEvents
    Loop

    ' Re-issue the same request
    http.Open "POST", http.ResponseURL, True
    http.setRequestHeader "Content-Type", "application/json"
    http.Send
End Sub

'------------------  RESPONSE ACCESS  -------------------------
Public Property Get ResponseText() As String
    ' Exposes the raw HTTP response text from the private http object.
    ' This is used for synchronous calls to return the response directly.
    ResponseText = http.responseText
End Property

Public Property Get ExtractedText() As String
    ' Extracts the "text" field from the parsed JSON response.
    ' The expected JSON structure is {"candidates": [{"content": {"parts": [{"text": "..."}]}}]}.
    ' Implements robust error handling to navigate the JSON structure.
    Dim candidates As Object
    Dim content As Object
    Dim parts As Object
    Dim part As Object

    On Error GoTo Err_ExtractedText

    If Not m_ParsedJson Is Nothing Then
        If m_ParsedJson.Exists("candidates") Then
            Set candidates = m_ParsedJson("candidates")
            If candidates.Count > 0 Then
                ' Access the first candidate (VBA collections are 1-based)
                If candidates(1).Exists("content") Then
                    Set content = candidates(1)("content")
                    If content.Exists("parts") Then
                        Set parts = content("parts")
                        If parts.Count > 0 Then
                            ' Access the first part
                            Set part = parts(1)
                            If part.Exists("text") Then
                                ExtractedText = part("text")
                            Else
                                ExtractedText = "Error: 'text' field not found in JSON response."
                            End If
                        Else
                            ExtractedText = "Error: 'parts' array is empty in JSON response."
                        End If
                    Else
                        ExtractedText = "Error: 'content' object does not contain 'parts' in JSON response."
                    End If
                Else
                    ExtractedText = "Error: First candidate does not contain 'content' in JSON response."
                End If
            Else
                ExtractedText = "Error: 'candidates' array is empty in JSON response."
            End If
        Else
            ExtractedText = "Error: 'candidates' field not found in JSON response."
        End If
    Else
        ExtractedText = "Error: JSON response not parsed."
    End If

    Exit Property

Err_ExtractedText:
    ExtractedText = "Error navigating JSON structure for text extraction: " & Err.Description
    On Error GoTo 0 ' Reset error handling
End Property

Public Property Get ExtractedError() As String
    ' Extracts an error message from the parsed JSON response.
    ' This property checks for a common "error" field in the JSON.
    ' It replaces the previous ExtractError logic.
    Dim errorObject As Object

    On Error GoTo Err_ExtractedError

    If Not m_ParsedJson Is Nothing Then
        If m_ParsedJson.Exists("error") Then
            Set errorObject = m_ParsedJson("error")
            If errorObject.Exists("message") Then
                ExtractedError = errorObject("message")
            Else
                ExtractedError = "Unknown error: 'error' object found but no 'message' field."
            End If
        Else
            ExtractedError = "No specific error message found in JSON response."
        End If
    Else
        ExtractedError = "Error: JSON response not parsed, cannot extract error."
    End If

    Exit Property

Err_ExtractedError:
    ExtractedError = "Error navigating JSON structure for error extraction: " & Err.Description
    On Error GoTo 0 ' Reset error handling
End Property

