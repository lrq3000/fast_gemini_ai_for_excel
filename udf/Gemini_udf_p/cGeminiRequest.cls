VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cGeminiRequest"
Option Explicit

Private http As Object 'late-bound → no reference needed
Private retryCount As Integer
Private maxRetries As Integer
Private maxDelayMs As Long
Public  TargetCell As Range
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
        Optional ByVal server_url As String = "") ' Added optional server_url parameter

    Set TargetCell = tgt
    maxRetries = retries
    maxDelayMs = delayMs
    retryCount = 0

    Dim prompt$, url$, body$, api_key_header As String
    prompt = userPrompt
    If word_count > 0 Then _
        prompt = prompt & ". Provide response in maximum " & word_count & " words"
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

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, True                 ' True = async
    http.setRequestHeader "Content-Type", "application/json"
    ' Add Authorization header for OpenAI-compatible servers if an API key is provided
    If api_key_header <> "" Then
        http.setRequestHeader "Authorization", api_key_header
    End If
    http.Send body
End Sub

'------------------  STATUS QUERIES  --------------------------
Public Property Get IsDone() As Boolean
    On Error Resume Next
    IsDone = (http.readyState = 4)
End Property

Public Sub CommitResult()
    Dim txt$
    Dim jsonResponse As Object ' For parsing JSON response

    If http.Status = 200 Then ' Success! HTTP OK, server returned content
        ' Attempt to parse as OpenAI-compatible response first
        On Error Resume Next ' Enable error handling for JSON parsing
        Set jsonResponse = JsonConverter.ParseJson(http.responseText)
        On Error GoTo 0 ' Disable error handling

        If Not jsonResponse Is Nothing And jsonResponse.Exists("choices") Then
            ' OpenAI-compatible response: content is in choices[0].message.content
            If jsonResponse("choices")(1).Exists("message") Then
                If jsonResponse("choices")(1)("message").Exists("content") Then
                    txt = jsonResponse("choices")(1)("message")("content")
                End If
            End If
        Else
            ' Fallback to Gemini response parsing if not OpenAI-compatible
            txt = ExtractContent(http.responseText)
        End If
    Else ' Failure! HTTP error or server returned an error, typically rate limit hit
        txt = "Error: " & ExtractError(http.responseText) ' extract error message
        txt = txt & vbLf & vbLf & "Retry with: " & TargetCell.Formula ' append original formula so that the user can retry
    End If

    Application.EnableEvents = False
    TargetCell.Value = txt
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
