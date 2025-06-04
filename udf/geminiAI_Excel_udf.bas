Attribute VB_Name = "geminiAI_Excel_udf"
Option Explicit

' Custom Excel UDF for interacting with Google Gemini API
Function Gemini_udf(text As String, api_key As String, Optional Model As String = "gemini-2.5-flash-preview-05-20", Optional word_count As Long = 0) As String
    Dim request As Object
    Dim response As String
    Dim API, DisplayText, error_result As String
    Dim status_code As Long

    ' API Info
    API = "https://generativelanguage.googleapis.com/v1beta/models/" & Model & ":generateContent?key=" & api_key

    ' API Key Check
    If api_key = "" Then
        MsgBox "Error: API key is blank!"
        Exit Function
    End If

    ' Append word count instruction if applicable
    If word_count > 0 Then
        text = text & ". Provide response in maximum " & word_count & " words"
    End If

    ' Clean text input to make it JSON-safe
    text = Replace(text, Chr(34), Chr(39))
    text = Replace(text, vbLf, " ")

    ' Create an HTTP request object
    Set request = CreateObject("MSXML2.XMLHTTP")
    With request
        .Open "POST", API, False
        .setRequestHeader "Content-Type", "application/json"
        .send "{""contents"":{""parts"":[{""text"":""" & text & """}]},""generationConfig"":{""temperature"":0.5}}"
        status_code = .Status
        response = .responseText
    End With

    ' Extract content
    If status_code = 200 Then
        DisplayText = ExtractContent(response)
    Else
        DisplayText = "Error : " & ExtractError(response)
    End If

    ' Clean up the object
    Set request = Nothing

    ' Return result
    Gemini_udf = DisplayText
End Function

' Helper: Extracts response text from Gemini JSON
Function ExtractContent(jsonString As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim TextValue As String
    Dim Content As String

    startPos = InStr(jsonString, """text"": """) + Len("""text"": """)
    endPos = InStr(startPos, jsonString, """") ' Find the position of the next double quote character
    TextValue = Mid(jsonString, startPos, endPos - startPos)
    Content = Trim(Replace(TextValue, "\""", Chr(34)))

    ' Fix for Excel formulas as response
    If Left(Trim(Content), 1) = "=" Then
        Content = "'" & Content
    End If

    ' Clean up common escape characters
    Content = Replace(Content, vbCrLf, "")
    Content = Replace(Content, vbLf, "")
    Content = Replace(Content, vbCr, "")
    Content = Replace(Content, "\n", vbNewLine)

    ' Trim trailing quote if present
    If Right(Content, 1) = """" Then
        Content = Left(Content, Len(Content) - 1)
    End If

    ExtractContent = Content
End Function

' Helper: Extracts error message from failed JSON API response
Function ExtractError(jsonString As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim TextValue As String
    Dim Content As String

    startPos = InStr(jsonString, """message"": """) + Len("""message"": """)
    endPos = InStr(startPos, jsonString, """") ' Find the position of the next double quote character
    TextValue = Mid(jsonString, startPos, endPos - startPos)
    Content = Trim(Replace(TextValue, "\""", Chr(34)))

    ' Fix for Excel formulas as response
    If Left(Trim(Content), 1) = "=" Then
        Content = "'" & Content
    End If

    ' Clean up formatting
    Content = Replace(Content, vbCrLf, "")
    Content = Replace(Content, vbLf, "")
    Content = Replace(Content, vbCr, "")

    ' Trim trailing quote if present
    If Right(Content, 1) = """" Then
        Content = Left(Content, Len(Content) - 1)
    End If

    ExtractError = Content
End Function
