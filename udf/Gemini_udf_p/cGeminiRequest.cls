VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "cGeminiRequest"
Option Explicit

Private http As Object 'late-bound → no reference needed
Public  TargetCell As Range

'------------------  SEND THE REQUEST  ------------------------
Public Sub Launch( _
        ByVal userPrompt As String, _
        ByVal key As String, _
        ByVal modelName As String, _
        ByVal word_count As Long, _
        ByVal tgt As Range)
    
    Set TargetCell = tgt
    
    Dim prompt$, url$, body$
    prompt = userPrompt
    If word_count > 0 Then _
        prompt = prompt & ". Provide response in maximum " & word_count & " words"
    prompt = Replace(prompt, """", "'")
    prompt = Replace(prompt, vbLf, " ")
    
    url = "https://generativelanguage.googleapis.com/v1beta/models/" & _
          modelName & ":generateContent?key=" & key
    body = "{""contents"":{""parts"":[{""text"":""" & prompt & """}]}," & _
           """generationConfig"":{""temperature"":0.5}}"
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", url, True                 'True = async
    http.setRequestHeader "Content-Type", "application/json"
    http.send body
End Sub

'------------------  STATUS QUERIES  --------------------------
Public Property Get IsDone() As Boolean
    On Error Resume Next
    IsDone = (http.readyState = 4)
End Property

Public Sub CommitResult()
    Dim txt$
    If http.Status = 200 Then
        txt = ExtractContent(http.responseText)
    Else
        txt = "Error: " & ExtractError(http.responseText)
    End If
    
    Application.EnableEvents = False
    TargetCell.Value = txt
    Application.EnableEvents = True
End Sub
