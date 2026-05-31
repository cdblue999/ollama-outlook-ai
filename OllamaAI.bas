' Copyright (C) 2026 ZMS
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <https://www.gnu.org/licenses/>.
Attribute VB_Name = "OllamaAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================
' Ollama Outlook AI - Local AI Email Assistant for Outlook 2013
' Version 1.1.1
'
' Single-file VBA module. No class modules, no external dependencies.
'
' Requirements:
'   - Ollama running at http://localhost:11434
'   - At least one model pulled (e.g., llama3.2:3b)
'   - Outlook 2013 or later
'
' All processing is fully local. No data leaves your machine.
'====================================================================

'====================================================================
' Constants
'====================================================================
Private Const APP_NAME As String = "OllamaAI"
Private Const APP_VERSION As String = "1.1.1"
Private Const OLLAMA_BASE_URL As String = "http://localhost:11434"
Private Const OLLAMA_DEFAULT_MODEL As String = "llama3.2:3b"
Private Const REQUEST_TIMEOUT_SECS As Long = 120
Private Const REG_PATH_ROOT As String = "HKEY_CURRENT_USER\Software\OllamaAI"

'====================================================================
' Initialization - Run this once after import or to refresh toolbars
'====================================================================
Public Sub Ollama_Initialize()
    On Error Resume Next
    
    ' Create default settings if first run
    Dim testVal As String
    testVal = GetRegSetting("Model", "")
    If testVal = "" Then
        SaveRegSetting "Model", OLLAMA_DEFAULT_MODEL
        SaveRegSetting "Prompt", "You are a helpful email assistant. Analyze the email content and respond professionally. Keep your response concise and well-structured."
        SaveRegSetting "Timeout", CStr(REQUEST_TIMEOUT_SECS)
    End If
    
    ' Add toolbar to all open compose windows
    AddToolbarToAllInspectors
End Sub

'====================================================================
' Add toolbar to all currently open inspectors
'====================================================================
Private Sub AddToolbarToAllInspectors()
    On Error Resume Next
    Dim insp As Outlook.Inspector
    Dim i As Long
    For i = 1 To Outlook.Application.Inspectors.Count
        Set insp = Outlook.Application.Inspectors.Item(i)
        If Not insp Is Nothing Then
            If insp.CurrentItem.Class = olMail Then
                EnsureToolbar insp
            End If
        End If
    Next i
End Sub

'====================================================================
' Create toolbar on an inspector if it doesn't exist
'====================================================================
Private Sub EnsureToolbar(ByVal Inspector As Outlook.Inspector)
    On Error Resume Next
    Dim cmdBar As Office.CommandBar
    Dim btn As Office.CommandBarButton
    
    ' Check if toolbar already exists
    Set cmdBar = Inspector.CommandBars("Ollama AI")
    If Not cmdBar Is Nothing Then Exit Sub
    
    ' Create new visible toolbar at top of compose window
    Set cmdBar = Inspector.CommandBars.Add("Ollama AI", msoBarTop, False, True)
    If cmdBar Is Nothing Then Exit Sub
    
    cmdBar.Visible = True
    
    ' Add "Process with AI" button
    Set btn = cmdBar.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Process with AI"
        .Tag = "OllamaAI_Process"
        .OnAction = "'OllamaAI.Ollama_ProcessWithAI'"
        .TooltipText = "Send email to Ollama AI for processing"
        .Style = msoButtonCaption
    End With
    
    cmdBar.Controls.Add Type:=msoControlSeparator
    
    ' Add "Settings" button
    Set btn = cmdBar.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Settings"
        .Tag = "OllamaAI_Settings"
        .OnAction = "'OllamaAI.Ollama_ShowConfigurationForm'"
        .TooltipText = "Configure Ollama AI settings"
        .Style = msoButtonCaption
    End With
    
    ' Add "About" button
    Set btn = cmdBar.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "About"
        .Tag = "OllamaAI_About"
        .OnAction = "'OllamaAI.Ollama_ShowAboutForm'"
        .TooltipText = "About Ollama Outlook AI"
        .Style = msoButtonCaption
    End With
End Sub

'====================================================================
' Optional: Call from ThisOutlookSession to auto-add toolbar to new windows
' Paste in ThisOutlookSession:
'   Private Sub Application_NewInspector(ByVal Inspector As Outlook.Inspector)
'       OllamaAI.Ollama_OnNewInspector Inspector
'   End Sub
'====================================================================
Public Sub Ollama_OnNewInspector(ByVal Inspector As Outlook.Inspector)
    On Error Resume Next
    If Inspector.CurrentItem.Class = olMail Then
        EnsureToolbar Inspector
    End If
End Sub

'====================================================================
' Configuration Form
'====================================================================
Public Sub Ollama_ShowConfigurationForm()
    On Error GoTo ConfigFormError
    
    Dim currentModel As String
    Dim currentPrompt As String
    Dim currentTimeout As String
    Dim psScript As String
    
    currentModel = GetRegSetting("Model", OLLAMA_DEFAULT_MODEL)
    currentPrompt = GetRegSetting("Prompt", "")
    currentTimeout = GetRegSetting("Timeout", CStr(REQUEST_TIMEOUT_SECS))
    
    ' Write PowerShell script to temp file and execute it
    Dim psCode As String
    psCode = "$models = @(); "
    psCode = psCode & "try { "
    psCode = psCode & "$resp = Invoke-RestMethod -Uri 'http://localhost:11434/api/tags' -TimeoutSec 5 -ErrorAction Stop; "
    psCode = psCode & "$models = $resp.models | ForEach-Object { $_.name }; "
    psCode = psCode & "} catch { $models = @('" & EscapeForPS(currentModel) & "') }; "
    psCode = psCode & "if ($models.Count -eq 0) { $models = @('" & EscapeForPS(currentModel) & "') }; "
    psCode = psCode & "$dm = '" & EscapeForPS(currentModel) & "'; "
    psCode = psCode & "$dp = '" & EscapeForPS(currentPrompt) & "'; "
    psCode = psCode & "$dt = " & currentTimeout & "; "
    psCode = psCode & "Add-Type -AssemblyName System.Windows.Forms; "
    psCode = psCode & "Add-Type -AssemblyName System.Drawing; "
    psCode = psCode & "$form = New-Object System.Windows.Forms.Form; "
    psCode = psCode & "$form.Text = 'Ollama AI - Settings'; "
    psCode = psCode & "$form.Size = New-Object System.Drawing.Size(520,420); "
    psCode = psCode & "$form.StartPosition = 'CenterScreen'; "
    psCode = psCode & "$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $pid).MainModule.FileName); "
    psCode = psCode & "$form.FormBorderStyle = 'FixedDialog'; "
    psCode = psCode & "$form.MaximizeBox = $false; "
    psCode = psCode & "$lblModel = New-Object System.Windows.Forms.Label; "
    psCode = psCode & "$lblModel.Text = 'AI Model:'; "
    psCode = psCode & "$lblModel.Location = New-Object System.Drawing.Point(15,20); "
    psCode = psCode & "$lblModel.Size = New-Object System.Drawing.Size(100,25); "
    psCode = psCode & "$form.Controls.Add($lblModel); "
    psCode = psCode & "$cmbModel = New-Object System.Windows.Forms.ComboBox; "
    psCode = psCode & "$cmbModel.Location = New-Object System.Drawing.Point(120,18); "
    psCode = psCode & "$cmbModel.Size = New-Object System.Drawing.Size(370,25); "
    psCode = psCode & "$cmbModel.DropDownStyle = 'DropDownList'; "
    psCode = psCode & "foreach ($m in $models) { [void]$cmbModel.Items.Add($m); if ($m -eq $dm) { $cmbModel.SelectedItem = $m } }; "
    psCode = psCode & "if ($cmbModel.SelectedIndex -eq -1 -and $cmbModel.Items.Count -gt 0) { $cmbModel.SelectedIndex = 0 }; "
    psCode = psCode & "$form.Controls.Add($cmbModel); "
    psCode = psCode & "$lblPrompt = New-Object System.Windows.Forms.Label; "
    psCode = psCode & "$lblPrompt.Text = 'System Prompt:'; "
    psCode = psCode & "$lblPrompt.Location = New-Object System.Drawing.Point(15,55); "
    psCode = psCode & "$lblPrompt.Size = New-Object System.Drawing.Size(100,25); "
    psCode = psCode & "$form.Controls.Add($lblPrompt); "
    psCode = psCode & "$txtPrompt = New-Object System.Windows.Forms.TextBox; "
    psCode = psCode & "$txtPrompt.Location = New-Object System.Drawing.Point(15,80); "
    psCode = psCode & "$txtPrompt.Size = New-Object System.Drawing.Size(475,120); "
    psCode = psCode & "$txtPrompt.Multiline = $true; "
    psCode = psCode & "$txtPrompt.ScrollBars = 'Vertical'; "
    psCode = psCode & "$txtPrompt.Text = $dp; "
    psCode = psCode & "$form.Controls.Add($txtPrompt); "
    psCode = psCode & "$lblTimeout = New-Object System.Windows.Forms.Label; "
    psCode = psCode & "$lblTimeout.Text = 'Timeout (seconds):'; "
    psCode = psCode & "$lblTimeout.Location = New-Object System.Drawing.Point(15,215); "
    psCode = psCode & "$lblTimeout.Size = New-Object System.Drawing.Size(120,25); "
    psCode = psCode & "$form.Controls.Add($lblTimeout); "
    psCode = psCode & "$nudTimeout = New-Object System.Windows.Forms.NumericUpDown; "
    psCode = psCode & "$nudTimeout.Location = New-Object System.Drawing.Point(140,213); "
    psCode = psCode & "$nudTimeout.Size = New-Object System.Drawing.Size(80,25); "
    psCode = psCode & "$nudTimeout.Minimum = 10; "
    psCode = psCode & "$nudTimeout.Maximum = 600; "
    psCode = psCode & "$nudTimeout.Value = $dt; "
    psCode = psCode & "$form.Controls.Add($nudTimeout); "
    psCode = psCode & "$lblInfo = New-Object System.Windows.Forms.Label; "
    psCode = psCode & "$lblInfo.Text = 'Settings are saved to Windows Registry (HKCU\Software\OllamaAI).'; "
    psCode = psCode & "$lblInfo.Location = New-Object System.Drawing.Point(15,255); "
    psCode = psCode & "$lblInfo.Size = New-Object System.Drawing.Size(475,25); "
    psCode = psCode & "$lblInfo.ForeColor = 'Gray'; "
    psCode = psCode & "$form.Controls.Add($lblInfo); "
    psCode = psCode & "$btnOK = New-Object System.Windows.Forms.Button; "
    psCode = psCode & "$btnOK.Text = 'Save'; "
    psCode = psCode & "$btnOK.Location = New-Object System.Drawing.Point(160,300); "
    psCode = psCode & "$btnOK.Size = New-Object System.Drawing.Size(90,30); "
    psCode = psCode & "$btnOK.Add_Click({ $form.Tag = $cmbModel.SelectedItem + '|' + $txtPrompt.Text + '|' + $nudTimeout.Value; $form.DialogResult = 'OK'; $form.Close() }); "
    psCode = psCode & "$form.Controls.Add($btnOK); "
    psCode = psCode & "$btnCancel = New-Object System.Windows.Forms.Button; "
    psCode = psCode & "$btnCancel.Text = 'Cancel'; "
    psCode = psCode & "$btnCancel.Location = New-Object System.Drawing.Point(270,300); "
    psCode = psCode & "$btnCancel.Size = New-Object System.Drawing.Size(90,30); "
    psCode = psCode & "$btnCancel.Add_Click({ $form.DialogResult = 'Cancel'; $form.Close() }); "
    psCode = psCode & "$form.Controls.Add($btnCancel); "
    psCode = psCode & "$result = $form.ShowDialog(); "
    psCode = psCode & "if ($result -eq 'OK') { "
    psCode = psCode & "$parts = $form.Tag -split '\|'; "
    psCode = psCode & "Write-Output $parts[0]; "
    psCode = psCode & "Write-Output $parts[1]; "
    psCode = psCode & "Write-Output ([int]$parts[2]); "
    psCode = psCode & "} else { Write-Output 'CANCELLED' }"
    
    Dim result As String
    result = RunPowerShell(psCode)
    
    If result <> "CANCELLED" And result <> "" Then
        Dim lines() As String
        lines = Split(result, vbCrLf)
        
        If UBound(lines) >= 0 Then
            Dim modelVal As String
            modelVal = Trim(lines(0))
            If modelVal <> "" Then SaveRegSetting "Model", modelVal
        End If
        
        If UBound(lines) >= 1 Then
            Dim promptVal As String
            promptVal = Trim(lines(1))
            If promptVal <> "" Then SaveRegSetting "Prompt", promptVal
        End If
        
        If UBound(lines) >= 2 Then
            Dim timeoutVal As String
            timeoutVal = Trim(lines(2))
            If timeoutVal <> "" Then SaveRegSetting "Timeout", timeoutVal
        End If
        
        MsgBox "Settings saved successfully.", vbInformation, APP_NAME
    End If
    
    Exit Sub
    
ConfigFormError:
    MsgBox "Could not open configuration form: " & Err.Description, vbExclamation, APP_NAME
End Sub

'====================================================================
' About Form
'====================================================================
Public Sub Ollama_ShowAboutForm()
    Dim msg As String
    msg = APP_NAME & " v" & APP_VERSION & vbCrLf & vbCrLf
    msg = msg & "Local AI Email Assistant for Outlook" & vbCrLf
    msg = msg & "Powered by Ollama" & vbCrLf & vbCrLf
    msg = msg & "All processing is fully local." & vbCrLf
    msg = msg & "No data leaves your machine."
    MsgBox msg, vbInformation, "About " & APP_NAME
End Sub

'====================================================================
' Health Check - Is Ollama Running?
'====================================================================
Public Function Ollama_IsOllamaRunning() As Boolean
    On Error GoTo NotRunning
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", OLLAMA_BASE_URL & "/api/tags", False
    http.SetTimeouts 5000, 5000, 5000, 5000
    http.SetOption 0, "Ollama-Outlook-AI/1.0"
    http.Send
    Ollama_IsOllamaRunning = (http.Status = 200)
    Exit Function
NotRunning:
    Ollama_IsOllamaRunning = False
End Function

'====================================================================
' Refresh AI Models List (fetches from Ollama)
'====================================================================
Public Sub Ollama_RefreshModelsList()
    On Error GoTo RefreshError
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", OLLAMA_BASE_URL & "/api/tags", False
    http.SetTimeouts 5000, 5000, 5000, 5000
    http.SetOption 0, "Ollama-Outlook-AI/1.0"
    http.Send
    
    If http.Status = 200 Then
        Dim json As String
        json = http.ResponseText
        Dim models As String
        models = ExtractModelNames(json)
        If models <> "" Then
            SaveRegSetting "AvailableModels", models
        Else
            SaveRegSetting "AvailableModels", OLLAMA_DEFAULT_MODEL
        End If
    Else
        SaveRegSetting "AvailableModels", OLLAMA_DEFAULT_MODEL
    End If
    Exit Sub
    
RefreshError:
    SaveRegSetting "AvailableModels", OLLAMA_DEFAULT_MODEL
End Sub

'====================================================================
' Extract model names from Ollama API JSON response
'====================================================================
Private Function ExtractModelNames(ByVal json As String) As String
    On Error Resume Next
    Dim result As String
    result = ""
    
    Dim modelsStart As Long
    modelsStart = InStr(json, """models"":")
    If modelsStart = 0 Then
        ExtractModelNames = OLLAMA_DEFAULT_MODEL
        Exit Function
    End If
    
    Dim searchPos As Long
    Dim nameStart As Long
    Dim nameEnd As Long
    Dim modelName As String
    searchPos = modelsStart
    
    Do
        nameStart = InStr(searchPos, json, """name"":""")
        If nameStart = 0 Or nameStart > modelsStart + 5000 Then Exit Do
        nameStart = nameStart + 8
        nameEnd = InStr(nameStart, json, """")
        If nameEnd = 0 Then Exit Do
        modelName = Mid(json, nameStart, nameEnd - nameStart)
        If result <> "" Then result = result & ","
        result = result & modelName
        searchPos = nameEnd + 1
    Loop
    
    If result = "" Then result = OLLAMA_DEFAULT_MODEL
    ExtractModelNames = result
End Function

'====================================================================
' Main AI Processing - Entry point from toolbar button
'====================================================================
Public Sub Ollama_ProcessWithAI()
    On Error GoTo ProcessError
    
    ' First, ensure toolbar exists on the active inspector
    Dim insp As Outlook.Inspector
    Set insp = GetActiveInspector()
    If Not insp Is Nothing Then
        If insp.CurrentItem.Class = olMail Then
            EnsureToolbar insp
        End If
    End If
    
    If Not Ollama_IsOllamaRunning() Then
        Dim msgOff As String
        msgOff = "Ollama is not running." & vbCrLf & vbCrLf & "Please start Ollama and try again."
        MsgBox msgOff, vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Set insp = GetActiveInspector()
    If insp Is Nothing Then
        Dim msgNoInsp As String
        msgNoInsp = "No email composition window detected." & vbCrLf & "Please open a new email or reply to an existing one."
        MsgBox msgNoInsp, vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim mail As Outlook.MailItem
    Set mail = insp.CurrentItem
    If mail Is Nothing Then
        MsgBox "Could not access the email content.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim subject As String
    Dim body As String
    subject = mail.Subject
    body = mail.Body
    
    If body = "" Then
        Dim msgEmpty As String
        msgEmpty = "The email body is empty." & vbCrLf & "Please write some content in your email first."
        MsgBox msgEmpty, vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim modelName As String
    Dim systemPrompt As String
    Dim timeoutSecs As Long
    
    modelName = GetRegSetting("Model", OLLAMA_DEFAULT_MODEL)
    systemPrompt = GetRegSetting("Prompt", "You are a helpful email assistant. Analyze the email content and respond professionally. Keep your response concise and well-structured.")
    timeoutSecs = CLng(GetRegSetting("Timeout", CStr(REQUEST_TIMEOUT_SECS)))
    
    Dim userContent As String
    userContent = "Email Subject: " & subject & vbCrLf & vbCrLf & "Email Body:" & vbCrLf & body
    
    Dim response As String
    response = Ollama_ProcessRequest(systemPrompt, userContent, modelName, timeoutSecs)
    
    If response <> "" Then
        Ollama_InsertResponseIntoEmail response, insp
    Else
        Dim msgNoResp As String
        msgNoResp = "No response received from Ollama." & vbCrLf & "Check that your model is downloaded and try again."
        MsgBox msgNoResp, vbExclamation, APP_NAME
    End If
    
    Exit Sub
    
ProcessError:
    MsgBox "Error processing email: " & Err.Description, vbCritical, APP_NAME
End Sub

'====================================================================
' Insert AI Response into Email
'====================================================================
Public Sub Ollama_InsertResponseIntoEmail(ByVal responseText As String, ByVal insp As Outlook.Inspector)
    On Error Resume Next
    
    If insp Is Nothing Then Exit Sub
    
    Dim wordDoc As Object
    Set wordDoc = insp.WordEditor
    
    If wordDoc Is Nothing Then
        Dim mail As Outlook.MailItem
        Set mail = insp.CurrentItem
        If Not mail Is Nothing Then
            mail.Body = mail.Body & vbCrLf & vbCrLf & responseText
        End If
        Exit Sub
    End If
    
    Dim selection As Object
    Set selection = wordDoc.Application.Selection
    If Not selection Is Nothing Then
        selection.Text = responseText
    End If
End Sub

'====================================================================
' API Communication - Send request to Ollama
'====================================================================
Private Function Ollama_ProcessRequest(ByVal systemPrompt As String, ByVal userContent As String, ByVal modelName As String, ByVal timeoutSecs As Long) As String
    On Error GoTo RequestError
    
    Dim payload As String
    payload = "{"
    payload = payload & """model"":""" & EscapeJson(modelName) & ""","
    payload = payload & """messages"":["
    payload = payload & "{""role"":""system"",""content"":""" & EscapeJson(systemPrompt) & """},"
    payload = payload & "{""role"":""user"",""content"":""" & EscapeJson(userContent) & """}"
    payload = payload & "],"
    payload = payload & """stream"":false"
    payload = payload & "}"
    
    Dim url As String
    url = OLLAMA_BASE_URL & "/v1/chat/completions"
    
    Dim response As String
    response = Ollama_SendHttpRequest(url, "POST", payload, timeoutSecs)
    
    If response = "" Then
        Ollama_ProcessRequest = ""
        Exit Function
    End If
    
    Dim content As String
    content = ExtractResponseContent(response)
    Ollama_ProcessRequest = content
    Exit Function
    
RequestError:
    Ollama_ProcessRequest = ""
End Function

'====================================================================
' HTTP Request using WinHttp
'====================================================================
Private Function Ollama_SendHttpRequest(ByVal url As String, ByVal method As String, ByVal payload As String, ByVal timeoutSecs As Long) As String
    On Error GoTo HttpError
    
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    Dim timeoutMs As Long
    timeoutMs = timeoutSecs * 1000
    
    http.Open method, url, False
    http.SetTimeouts 10000, 15000, timeoutMs, timeoutMs
    http.SetOption 0, "Ollama-Outlook-AI/1.0"
    http.SetRequestHeader "Content-Type", "application/json"
    
    If method = "POST" Then
        http.Send payload
    Else
        http.Send
    End If
    
    If http.Status = 200 Then
        Ollama_SendHttpRequest = http.ResponseText
    Else
        Ollama_SendHttpRequest = ""
    End If
    
    Exit Function
    
HttpError:
    Ollama_SendHttpRequest = ""
End Function

'====================================================================
' Extract response content from Ollama API JSON
'====================================================================
Private Function ExtractResponseContent(ByVal json As String) As String
    On Error Resume Next
    
    Dim contentMarker As String
    contentMarker = """content"":"""
    
    Dim msgStart As Long
    msgStart = InStr(json, """message"":{")
    If msgStart = 0 Then
        msgStart = InStr(json, contentMarker)
        If msgStart = 0 Then
            ExtractResponseContent = ""
            Exit Function
        End If
    Else
        msgStart = InStr(msgStart, json, contentMarker)
        If msgStart = 0 Then
            ExtractResponseContent = ""
            Exit Function
        End If
    End If
    
    Dim contentStart As Long
    contentStart = msgStart + Len(contentMarker)
    
    Dim contentEnd As Long
    contentEnd = contentStart
    
    Do While contentEnd <= Len(json)
        Dim ch As String
        ch = Mid(json, contentEnd, 1)
        If ch = "\" Then
            contentEnd = contentEnd + 2
        ElseIf ch = """" Then
            Exit Do
        Else
            contentEnd = contentEnd + 1
        End If
    Loop
    
    If contentEnd > contentStart Then
        ExtractResponseContent = Mid(json, contentStart, contentEnd - contentStart)
        ExtractResponseContent = Replace(ExtractResponseContent, "\\n", vbCrLf)
        ExtractResponseContent = Replace(ExtractResponseContent, "\\t", vbTab)
        ExtractResponseContent = Replace(ExtractResponseContent, "\\r", vbCr)
        ExtractResponseContent = Replace(ExtractResponseContent, "\""", """")
        ExtractResponseContent = Replace(ExtractResponseContent, "\\", "\")
    Else
        ExtractResponseContent = ""
    End If
End Function

'====================================================================
' Escape strings for JSON
'====================================================================
Private Function EscapeJson(ByVal text As String) As String
    Dim result As String
    result = text
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\\n")
    result = Replace(result, vbCr, "\\r")
    result = Replace(result, vbLf, "\\n")
    result = Replace(result, vbTab, "\\t")
    EscapeJson = result
End Function

'====================================================================
' Escape strings for PowerShell single-quoted strings
'====================================================================
Private Function EscapeForPS(ByVal text As String) As String
    Dim result As String
    result = Replace(text, "'", "''")
    EscapeForPS = result
End Function

'====================================================================
' Registry Helpers
'====================================================================
Private Function GetRegSetting(ByVal key As String, ByVal defaultValue As String) As String
    On Error Resume Next
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    Dim regPath As String
    regPath = REG_PATH_ROOT & "\" & key
    Dim value As Variant
    value = ws.RegRead(regPath)
    If Err.Number = 0 Then
        GetRegSetting = CStr(value)
    Else
        GetRegSetting = defaultValue
    End If
End Function

Private Sub SaveRegSetting(ByVal key As String, ByVal value As String)
    On Error Resume Next
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    Dim regPath As String
    regPath = REG_PATH_ROOT & "\" & key
    ws.RegWrite REG_PATH_ROOT, "", "REG_SZ"
    ws.RegWrite regPath, value, "REG_SZ"
End Sub

'====================================================================
' Run PowerShell script and capture output
'====================================================================
Private Function RunPowerShell(ByVal script As String) As String
    On Error GoTo PSError
    
    Dim tempFile As String
    Dim tempDir As String
    tempDir = Environ("TEMP")
    tempFile = tempDir & "\ollama_config.ps1"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    Set file = fso.CreateTextFile(tempFile, True)
    file.Write script
    file.Close
    
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    Dim exec As Object
    Dim cmd As String
    
    cmd = "powershell.exe -ExecutionPolicy Bypass -File """ & tempFile & """"
    Set exec = ws.Exec(cmd)
    
    Dim startTime As Double
    startTime = Timer
    Do While exec.Status = 0
        DoEvents
        If Timer > startTime + 60 Then
            exec.Terminate
            Exit Do
        End If
    Loop
    
    Dim output As String
    If exec.Status <> 0 Then
        output = exec.StdOut.ReadAll()
    End If
    
    On Error Resume Next
    fso.DeleteFile tempFile
    On Error GoTo 0
    
    RunPowerShell = output
    Exit Function
    
PSError:
    RunPowerShell = ""
End Function

'====================================================================
' Get Active Inspector
'====================================================================
Private Function GetActiveInspector() As Outlook.Inspector
    On Error Resume Next
    Set GetActiveInspector = Outlook.Application.ActiveInspector
End Function

'====================================================================
' Utility: Null-safe string conversion
'====================================================================
Private Function NzStr(ByVal value As Variant) As String
    If IsNull(value) Or IsMissing(value) Then
        NzStr = ""
    Else
        NzStr = CStr(value)
    End If
End Function

'====================================================================
' End of Module
'====================================================================
