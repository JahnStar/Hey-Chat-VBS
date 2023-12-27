' Developed by Halil Emre Yildiz (Github: @JahnStar)
Dim url, default_apiKey, apiKey
url = "https://api.openai.com/v1/chat/completions"
apiKey = "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"

Dim requestBody, message
message = InputBox("You:", "Hey ChatVBS v0.1", "Hello, can you assist me?")

Dim request, responseContent
Set request = CreateObject("Microsoft.XMLHTTP")

run = True
firstPrompt = "You are a helpful AI assistant named JAHNVIS, developed by Halil Emre YILDIZ and running in a VBS script."
Do While run 
    requestBody = "{""model"": ""gpt-3.5-turbo"", ""messages"": [{""role"": ""system"", ""content"": """ & firstPrompt & """}, {""role"": ""user"", ""content"": """ & message & """}], ""temperature"": 0.7}"

    request.Open "POST", url, False
    request.setRequestHeader "Content-Type", "application/json"
    request.setRequestHeader "Authorization", "Bearer " & apiKey
    request.send requestBody

    If request.Status = 200 Then
        responseContent = ParseJSON(request.responseText, "content")
        message = InputBox("JAHNVIS: " & responseContent & vbCrLf & vbCrLf & "You:", "Hey ChatVBS v0.1")
        If IsEmpty(message) Then
            run = False
        End If
    Else
        MsgBox "Request failed with status: " & request.Status
        WScript.Quit 
    End If
Loop

Function ParseJSON(jsonString, key)
    Dim startPos, endPos, keyPos, valueStartPos, valueEndPos
    Dim keyValue, valueStr

    ' Replace escaped characters
    jsonString = Replace(jsonString, "\""", "'")
    jsonString = Replace(jsonString, "\\", "\")

    startPos = InStr(jsonString, """" & key & """") ' Start position of the switch
    keyPos = InStr(startPos, jsonString, ":") ' Position of the ":" character of the key

    If keyPos > 0 Then
        valueStartPos = InStr(keyPos, jsonString, """") + 1 ' Start position of the value
        valueEndPos = InStr(valueStartPos, jsonString, """") ' End position of value

        valueStr = Mid(jsonString, valueStartPos, valueEndPos - valueStartPos) ' Value string
        ParseJSON = valueStr ' Return value
    Else
        ParseJSON = "" ' Return empty string if key not found
    End If
End Function

' JSON request body sent to the API:
' {
'   "model": "gpt-3.5-turbo",
'   "messages": [
'     {
'       "role": "system",
'       "content": "You are a helpful assistant."
'     },
'     {
'       "role": "user",
'       "content": "Hello!"
'     },
'    "temperature": 0.7
'   ]
' }

' JSON response body received from the API:
' {
'   "id": "chatcmpl-123",
'   "object": "chat.completion",
'   "created": 1677652288,
'   "model": "gpt-3.5-turbo-0613",
'   "system_fingerprint": "fp_44709d6fcb",
'   "choices": [{
'     "index": 0,
'     "message": {
'       "role": "assistant",
'       "content": "\n\nHello there, how may I assist you today?",
'     },
'     "logprobs": null,
'     "finish_reason": "stop"
'   }],
'   "usage": {
'     "prompt_tokens": 9,
'     "completion_tokens": 12,
'     "total_tokens": 21
'   }
' }
