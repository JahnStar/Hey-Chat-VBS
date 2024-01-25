'***************************************************************************************************
' Author: Halil Emre Yildiz
' GitHub: @JahnStar
'***************************************************************************************************
Dim fso, CurrentDirectory, iniPath, file, endpoint, apiKey, requestBody
Set fso = CreateObject("Scripting.FileSystemObject")
CurrentDirectory = fso.GetParentFolderName(WScript.ScriptFullName)
iniPath = fso.BuildPath(CurrentDirectory, "request.ini")

Set file = fso.OpenTextFile(iniPath, 1) ' 1 means ForReading
endpoint = file.ReadLine
apiKey = file.ReadLine
Do Until file.AtEndOfStream
    requestBody = requestBody & file.ReadLine
Loop

If MsgBox(requestBody, vbYesNo, "Request?" ) = vbNo Then
    WScript.Quit
End If

Dim request, responseContent
Set request = CreateObject("Microsoft.XMLHTTP")

request.Open "POST", endpoint, False
request.setRequestHeader "Content-Type", "application/json"
request.setRequestHeader "Authorization", "Bearer " & apiKey
request.send requestBody

MsgBox request.responseText, vbInformation, "Response"

' Request example:
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
' Response example:
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