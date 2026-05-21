Attribute VB_Name = "OllamaTest"
Option Explicit

Public Sub Ollama_Hello()
    MsgBox "Import works!", vbInformation, "Ollama AI Test"
End Sub

Public Sub Ollama_CheckOllama()
    On Error GoTo NotRunning
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "http://localhost:11434/api/tags", False
    http.SetTimeouts 5000, 5000, 5000, 5000
    http.Send
    If http.Status = 200 Then
        MsgBox "Ollama is running!", vbInformation, "Ollama AI"
    Else
        MsgBox "Ollama responded with status: " & http.Status, vbExclamation, "Ollama AI"
    End If
    Exit Sub
NotRunning:
    MsgBox "Ollama is not running at localhost:11434", vbExclamation, "Ollama AI"
End Sub
