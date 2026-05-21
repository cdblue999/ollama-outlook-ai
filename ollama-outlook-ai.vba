Attribute VB_Name = "OllamaAI"
'====================================================================
' Ollama Outlook AI - Local AI Email Assistant for Outlook 2013
' Version 1.0.0
'
' Recoded from key-assist-outlook (paulchi-intel) to use Ollama's
' OpenAI-compatible local API endpoint instead of Intel ExpertGPT.
'
' Requirements:
'   - Ollama running at http://localhost:11434
'   - At least one model pulled (e.g., llama3.2:3b)
'   - Outlook 2013 or later
'
' All processing is fully local. No data leaves your machine.
'====================================================================

Option Explicit

'====================================================================
' Constants
'====================================================================
Private Const APP_NAME As String = "OllamaAI"
Private Const APP_VERSION As String = "1.0.0"
Private Const OLLAMA_BASE_URL As String = "http://localhost:11434"
Private Const OLLAMA_DEFAULT_MODEL As String = "llama3.2:3b"
Private Const REQUEST_TIMEOUT_SECS As Long = 120
Private Const REG_PATH_ROOT As String = "HKEY_CURRENT_USER\Software\OllamaAI"

'====================================================================
' Module-level variables
'====================================================================
Private m_Inspectors As Outlook.Inspectors
Private m_OllamaHandler As COllamaAIHandler

'====================================================================
' Windows API declarations
'====================================================================
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As LongPtr, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, ByVal lpOperation As String, _
        ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

'====================================================================
' CLASS: COllamaAIHandler - Handles Outlook inspector events
'====================================================================
' VBA class defined as a separate module-level implementation
' using the WithEvents pattern embedded in this module.

Private WithEvents m_InspectorEvents As Outlook.Inspectors

Private Sub Class_Initialize()
    ' No-op: initialization happens through SetInspectors
End Sub

Public Sub SetInspectors(ByVal inspectors As Outlook.Inspectors)
    Set m_InspectorEvents = inspectors
End Sub

Private Sub m_InspectorEvents_NewInspector(ByVal Inspector As Outlook.Inspector)
    On Error GoTo HandlerError
    If Inspector.CurrentItem.Class = olMail Then
        AddToolbarButton Inspector
    End If
    Exit Sub
HandlerError:
    ' Silently handle errors during inspector creation
End Sub

Private Sub AddToolbarButton(ByVal Inspector As Outlook.Inspector)
    On Error Resume Next
    Dim cmdBar As Office.CommandBar
    Dim cmdBarPopup As Office.CommandBarPopup
    Dim btn As Office.CommandBarButton
    Dim existingBtn As Office.CommandBarButton
    
    ' Get the command bar for the inspector
    Set cmdBar = Inspector.CommandBars("Menu Bar")
    If cmdBar Is Nothing Then Exit Sub
    
    ' Check if button already exists
    Set existingBtn = Nothing
    On Error Resume Next
    Set existingBtn = cmdBar.FindControl(Tag:="OllamaAI")
    On Error GoTo 0
    If Not existingBtn Is Nothing Then Exit Sub
    
    ' Add popup button to the toolbar
    Set cmdBarPopup = cmdBar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    If cmdBarPopup Is Nothing Then Exit Sub
    
    cmdBarPopup.Caption = "Ollama AI"
    cmdBarPopup.Tag = "OllamaAI"
    cmdBarPopup.TooltipText = "Ollama AI Email Assistant"
    
    ' Add "Process with AI" button
    Set btn = cmdBarPopup.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Process with AI"
        .Tag = "OllamaAI_Process"
        .OnAction = "'" & APP_NAME & ".Ollama_ProcessWithAI'"
        .TooltipText = "Send email to Ollama AI for processing"
        .Style = msoButtonCaption
    End With
    
    ' Add separator
    cmdBarPopup.Controls.Add Type:=msoControlSeparator
    
    ' Add "Settings" button
    Set btn = cmdBarPopup.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Settings..."
        .Tag = "OllamaAI_Settings"
        .OnAction = "'" & APP_NAME & ".Ollama_ShowConfigurationForm'"
        .TooltipText = "Configure Ollama AI settings"
        .Style = msoButtonCaption
    End With
    
    ' Add "About" button
    Set btn = cmdBarPopup.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "About..."
        .Tag = "OllamaAI_About"
        .OnAction = "'" & APP_NAME & ".Ollama_ShowAboutForm'"
        .TooltipText = "About Ollama Outlook AI"
        .Style = msoButtonCaption
    End With
End Sub

'====================================================================
' Outlook Startup - Called automatically
'====================================================================
Private Sub Application_Startup()
    Ollama_Initialize
End Sub

'====================================================================
' Initialization
'====================================================================
Public Sub Ollama_Initialize()
    On Error Resume Next
    Dim olApp As Outlook.Application
    Set olApp = Outlook.Application
    
    Set m_Inspectors = olApp.Inspectors
    
    ' Create handler class and wire up events
    Dim handler As COllamaAIHandler
    Set handler = New COllamaAIHandler
    handler.SetInspectors m_Inspectors
    Set m_OllamaHandler = handler
    
    ' Ensure default settings exist
    Dim testVal As String
    testVal = GetRegSetting("Model", "")
    If testVal = "" Then
        SaveRegSetting "Model", OLLAMA_DEFAULT_MODEL
        SaveRegSetting "Prompt", "You are a helpful email assistant. Analyze the email content and respond professionally. Keep your response concise and well-structured."
        SaveRegSetting "Timeout", CStr(REQUEST_TIMEOUT_SECS)
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
    
    ' Load current settings
    currentModel = GetRegSetting("Model", OLLAMA_DEFAULT_MODEL)
    currentPrompt = GetRegSetting("Prompt", "")
    currentTimeout = GetRegSetting("Timeout", CStr(REQUEST_TIMEOUT_SECS))
    
    ' Create PowerShell script for the configuration form
    psScript = "$models = @(); " & vbCrLf
    psScript = psScript & "try { " & vbCrLf
    psScript = psScript & "  $resp = Invoke-RestMethod -Uri 'http://localhost:11434/api/tags' -TimeoutSec 5 -ErrorAction Stop; " & vbCrLf
    psScript = psScript & "  $models = $resp.models | ForEach-Object { $_.name }; " & vbCrLf
    psScript = psScript & "} catch { $models = @('" & currentModel & "') }; " & vbCrLf
    psScript = psScript & "if ($models.Count -eq 0) { $models = @('" & currentModel & "') }; " & vbCrLf
    psScript = psScript & "$defaultModel = '" & EscapeForPS(currentModel) & "'; " & vbCrLf
    psScript = psScript & "$defaultPrompt = '" & EscapeForPS(currentPrompt) & "'; " & vbCrLf
    psScript = psScript & "$defaultTimeout = " & currentTimeout & "; " & vbCrLf
    psScript = psScript & "$selectedModel = $models[0]; " & vbCrLf
    psScript = psScript & "Add-Type -AssemblyName System.Windows.Forms; " & vbCrLf
    psScript = psScript & "Add-Type -AssemblyName System.Drawing; " & vbCrLf
    psScript = psScript & "$form = New-Object System.Windows.Forms.Form; " & vbCrLf
    psScript = psScript & "$form.Text = 'Ollama AI - Settings'; " & vbCrLf
    psScript = psScript & "$form.Size = New-Object System.Drawing.Size(520,420); " & vbCrLf
    psScript = psScript & "$form.StartPosition = 'CenterScreen'; " & vbCrLf
    psScript = psScript & "$form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $pid).MainModule.FileName); " & vbCrLf
    psScript = psScript & "$form.FormBorderStyle = 'FixedDialog'; " & vbCrLf
    psScript = psScript & "$form.MaximizeBox = $false; " & vbCrLf
    
    ' Model selection
    psScript = psScript & "$lblModel = New-Object System.Windows.Forms.Label; " & vbCrLf
    psScript = psScript & "$lblModel.Text = 'AI Model:'; " & vbCrLf
    psScript = psScript & "$lblModel.Location = New-Object System.Drawing.Point(15,20); " & vbCrLf
    psScript = psScript & "$lblModel.Size = New-Object System.Drawing.Size(100,25); " & vbCrLf
    psScript = psScript & "$form.Controls.Add($lblModel); " & vbCrLf
    
    psScript = psScript & "$cmbModel = New-Object System.Windows.Forms.ComboBox; " & vbCrLf
    psScript = psScript & "$cmbModel.Location = New-Object System.Drawing.Point(120,18); " & vbCrLf
    psScript = psScript & "$cmbModel.Size = New-Object System.Drawing.Size(370,25); " & vbCrLf
    psScript = psScript & "$cmbModel.DropDownStyle = 'DropDownList'; " & vbCrLf
    psScript = psScript & "foreach ($m in $models) { [void]$cmbModel.Items.Add($m); " & vbCrLf
    psScript = psScript & "  if ($m -eq $defaultModel) { $cmbModel.SelectedItem = $m } }; " & vbCrLf
    psScript = psScript & "if ($cmbModel.SelectedIndex -eq -1 -and $cmbModel.Items.Count -gt 0) { $cmbModel.SelectedIndex = 0 }; " & vbCrLf
    psScript = psScript & "$selectedModel = $cmbModel.SelectedItem; " & vbCrLf
    psScript = psScript & "$form.Controls.Add($cmbModel); " & vbCrLf
    
    ' Prompt label
    psScript = psScript & "$lblPrompt = New-Object System.Windows.Forms.Label; " & vbCrLf
    psScript = psScript & "$lblPrompt.Text = 'System Prompt:'; " & vbCrLf
    psScript = psScript & "$lblPrompt.Location = New-Object System.Drawing.Point(15,55); " & vbCrLf
    psScript = psScript & "$lblPrompt.Size = New-Object System.Drawing.Size(100,25); " & vbCrLf
    psScript = psScript & "$form.Controls.Add($lblPrompt); " & vbCrLf
    
    ' Prompt textbox
    psScript = psScript & "$txtPrompt = New-Object System.Windows.Forms.TextBox; " & vbCrLf
    psScript = psScript & "$txtPrompt.Location = New-Object System.Drawing.Point(15,80); " & vbCrLf
    psScript = psScript & "$txtPrompt.Size = New-Object System.Drawing.Size(475,120); " & vbCrLf
    psScript = psScript & "$txtPrompt.Multiline = $true; " & vbCrLf
    psScript = psScript & "$txtPrompt.ScrollBars = 'Vertical'; " & vbCrLf
    psScript = psScript & "$txtPrompt.Text = $defaultPrompt; " & vbCrLf
    psScript = psScript & "$form.Controls.Add($txtPrompt); " & vbCrLf
    
    ' Timeout label
    psScript = psScript & "$lblTimeout = New-Object System.Windows.Forms.Label; " & vbCrLf
    psScript = psScript & "$lblTimeout.Text = 'Timeout (seconds):'; " & vbCrLf
    psScript = psScript & "$lblTimeout.Location = New-Object System.Drawing.Point(15,215); " & vbCrLf
    psScript = psScript & "$lblTimeout.Size = New-Object System.Drawing.Size(120,25); " & vbCrLf
    psScript = psScript & "$form.Controls.Add($lblTimeout); " & vbCrLf
    
    ' Timeout nud
    psScript = psScript & "$nudTimeout = New-Object System.Windows.Forms.NumericUpDown; " & vbCrLf
    psScript = psScript & "$nudTimeout.Location = New-Object System.Drawing.Point(140,213); " & vbCrLf
    psScript = psScript & "$nudTimeout.Size = New-Object System.Drawing.Size(80,25); " & vbCrLf
    psScript = psScript & "$nudTimeout.Minimum = 10; " & vbCrLf
    psScript = psScript & "$nudTimeout.Maximum = 600; " & vbCrLf
    psScript = psScript & "$nudTimeout.Value = $defaultTimeout; " & vbCrLf
    psScript = psScript & "$form.Controls.Add($nudTimeout); " & vbCrLf
    
    ' Info label
    psScript = psScript & "$lblInfo = New-Object System.Windows.Forms.Label; " & vbCrLf
    psScript = psScript & "$lblInfo.Text = 'Settings are saved to Windows Registry (HKCU\Software\OllamaAI).'; " & vbCrLf
    psScript = psScript & "$lblInfo.Location = New-Object System.Drawing.Point(15,255); " & vbCrLf
    psScript = psScript & "$lblInfo.Size = New-Object System.Drawing.Size(475,25); " & vbCrLf
    psScript = psScript & "$lblInfo.ForeColor = 'Gray'; " & vbCrLf
    psScript = psScript & "$form.Controls.Add($lblInfo); " & vbCrLf
    
    ' OK button
    psScript = psScript & "$btnOK = New-Object System.Windows.Forms.Button; " & vbCrLf
    psScript = psScript & "$btnOK.Text = 'Save'; " & vbCrLf
    psScript = psScript & "$btnOK.Location = New-Object System.Drawing.Point(160,300); " & vbCrLf
    psScript = psScript & "$btnOK.Size = New-Object System.Drawing.Size(90,30); " & vbCrLf
    psScript = psScript & "$btnOK.Add_Click({ " & vbCrLf
    psScript = psScript & "  $selectedModel = $cmbModel.SelectedItem; " & vbCrLf
    psScript = psScript & "  $form.Tag = $selectedModel + '|' + $txtPrompt.Text + '|' + $nudTimeout.Value; " & vbCrLf
    psScript = psScript & "  $form.DialogResult = 'OK'; " & vbCrLf
    psScript = psScript & "  $form.Close() }); " & vbCrLf
    psScript = psScript & "$form.Controls.Add($btnOK); " & vbCrLf
    
    ' Cancel button
    psScript = psScript & "$btnCancel = New-Object System.Windows.Forms.Button; " & vbCrLf
    psScript = psScript & "$btnCancel.Text = 'Cancel'; " & vbCrLf
    psScript = psScript & "$btnCancel.Location = New-Object System.Drawing.Point(270,300); " & vbCrLf
    psScript = psScript & "$btnCancel.Size = New-Object System.Drawing.Size(90,30); " & vbCrLf
    psScript = psScript & "$btnCancel.Add_Click({ $form.DialogResult = 'Cancel'; $form.Close() }); " & vbCrLf
    psScript = psScript & "$form.Controls.Add($btnCancel); " & vbCrLf
    
    ' Show form
    psScript = psScript & "$result = $form.ShowDialog(); " & vbCrLf
    psScript = psScript & "if ($result -eq 'OK') { " & vbCrLf
    psScript = psScript & "  $parts = $form.Tag -split '\|'; " & vbCrLf
    psScript = psScript & "  Write-Output $parts[0]; " & vbCrLf
    psScript = psScript & "  Write-Output $parts[1]; " & vbCrLf
    psScript = psScript & "  Write-Output ([int]$parts[2]); " & vbCrLf
    psScript = psScript & "} else { Write-Output 'CANCELLED' }"
    
    ' Execute PowerShell and parse results
    Dim result As String
    result = RunPowerShell(psScript)
    
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
    MsgBox APP_NAME & " v" & APP_VERSION & vbCrLf & vbCrLf & _
        "Local AI Email Assistant for Outlook" & vbCrLf & _
        "Powered by Ollama" & vbCrLf & vbCrLf & _
        "All processing is fully local." & vbCrLf & _
        "No data leaves your machine.", _
        vbInformation, "About " & APP_NAME
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
    http.SetOption 0, "Ollama-Outlook-AI/1.0" ' UserAgent
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
        
        ' Parse model names from JSON response
        ' Expected: {"models":[{"name":"llama3.2:3b","modified_at":"...","size":...}]}
        Dim models As String
        models = ExtractModelNames(json)
        
        If models <> "" Then
            SaveRegSetting "AvailableModels", models
        Else
            SaveRegSetting "AvailableModels", OLLAMA_DEFAULT_MODEL
        End If
    Else
        ' Fall back to default model if Ollama not available
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
    
    ' Simple JSON parser for known structure
    ' Find "models":[{ ... "name":"..." ... }]
    Dim modelsStart As Long
    modelsStart = InStr(json, """models"":")
    If modelsStart = 0 Then
        ExtractModelNames = OLLAMA_DEFAULT_MODEL
        Exit Function
    End If
    
    ' Extract each "name":"..." value
    Dim searchPos As Long
    Dim nameStart As Long
    Dim nameEnd As Long
    Dim modelName As String
    searchPos = modelsStart
    
    Do
        nameStart = InStr(searchPos, json, """name"":""")
        If nameStart = 0 Or nameStart > modelsStart + 5000 Then Exit Do
        
        nameStart = nameStart + 8 ' Skip past "name":"
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
    
    ' Check if Ollama is running
    If Not Ollama_IsOllamaRunning() Then
        MsgBox "Ollama is not running." & vbCrLf & vbCrLf & _
            "Please start Ollama and try again.", _
            vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Get active inspector and email
    Dim insp As Outlook.Inspector
    Set insp = GetActiveInspector()
    If insp Is Nothing Then
        MsgBox "No email composition window detected." & vbCrLf & _
            "Please open a new email or reply to an existing one.", _
            vbExclamation, APP_NAME
        Exit Sub
    End If
    
    Dim mail As Outlook.MailItem
    Set mail = insp.CurrentItem
    If mail Is Nothing Then
        MsgBox "Could not access the email content.", vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Get email content
    Dim subject As String
    Dim body As String
    subject = mail.Subject
    body = mail.Body
    
    If body = "" Then
        MsgBox "The email body is empty." & vbCrLf & _
            "Please write some content in your email first.", _
            vbExclamation, APP_NAME
        Exit Sub
    End If
    
    ' Get settings from registry
    Dim modelName As String
    Dim systemPrompt As String
    Dim timeoutSecs As Long
    
    modelName = GetRegSetting("Model", OLLAMA_DEFAULT_MODEL)
    systemPrompt = GetRegSetting("Prompt", "You are a helpful email assistant. Analyze the email content and respond professionally. Keep your response concise and well-structured.")
    timeoutSecs = CLng(GetRegSetting("Timeout", CStr(REQUEST_TIMEOUT_SECS)))
    
    ' Build user content with context
    Dim userContent As String
    userContent = "Email Subject: " & subject & vbCrLf & vbCrLf & "Email Body:" & vbCrLf & body
    
    ' Show progress to user
    Dim progressMsg As String
    progressMsg = "Processing email with " & modelName & "..." & vbCrLf & _
        "This may take a moment for local AI inference."
    
    ' Process the request
    Dim response As String
    response = Ollama_ProcessRequest(systemPrompt, userContent, modelName, timeoutSecs)
    
    If response <> "" Then
        ' Insert response into email
        Ollama_InsertResponseIntoEmail response, insp
    Else
        MsgBox "No response received from Ollama." & vbCrLf & _
            "Check that your model is downloaded and try again.", _
            vbExclamation, APP_NAME
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
    
    ' Get the Word editor to insert at cursor position
    Dim wordDoc As Object
    Set wordDoc = insp.WordEditor
    
    If wordDoc Is Nothing Then
        ' Fallback: replace selection
        Dim mail As Outlook.MailItem
        Set mail = insp.CurrentItem
        If Not mail Is Nothing Then
            mail.Body = mail.Body & vbCrLf & vbCrLf & responseText
        End If
        Exit Sub
    End If
    
    ' Insert at current cursor position in the Word editor
    Dim selection As Object
    Set selection = wordDoc.Application.Selection
    If Not selection Is Nothing Then
        selection.Text = responseText
    End If
End Sub

'====================================================================
' API Communication - Send request to Ollama
'====================================================================
Private Function Ollama_ProcessRequest(ByVal systemPrompt As String, _
                                       ByVal userContent As String, _
                                       ByVal modelName As String, _
                                       ByVal timeoutSecs As Long) As String
    On Error GoTo RequestError
    
    ' Build the OpenAI-compatible payload
    Dim payload As String
    payload = "{"
    payload = payload & """model"":""" & EscapeJson(modelName) & ""","
    payload = payload & """messages"":["
    payload = payload & "{""role"":""system"",""content"":""" & EscapeJson(systemPrompt) & """},"
    payload = payload & "{""role"":""user"",""content"":""" & EscapeJson(userContent) & """}"
    payload = payload & "],"
    payload = payload & """stream"":false"
    payload = payload & "}"
    
    ' Send request
    Dim url As String
    url = OLLAMA_BASE_URL & "/v1/chat/completions"
    
    Dim response As String
    response = Ollama_SendHttpRequest(url, "POST", payload, timeoutSecs)
    
    If response = "" Then
        Ollama_ProcessRequest = ""
        Exit Function
    End If
    
    ' Parse response to extract content
    ' Expected: {"id":"...","choices":[{"index":0,"message":{"role":"assistant","content":"..."}},...]}
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
Private Function Ollama_SendHttpRequest(ByVal url As String, _
                                        ByVal method As String, _
                                        ByVal payload As String, _
                                        ByVal timeoutSecs As Long) As String
    On Error GoTo HttpError
    
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Convert timeout to milliseconds
    Dim timeoutMs As Long
    timeoutMs = timeoutSecs * 1000
    
    ' Open connection
    http.Open method, url, False
    
    ' Set timeouts (resolve, connect, send, receive)
    http.SetTimeouts 10000, 15000, timeoutMs, timeoutMs
    
    ' Set user agent
    http.SetOption 0, "Ollama-Outlook-AI/1.0"
    
    ' Set headers
    http.SetRequestHeader "Content-Type", "application/json"
    
    ' Send request
    If method = "POST" Then
        http.Send payload
    Else
        http.Send
    End If
    
    ' Check response
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
    
    ' Parse: find "choices":[{"index":0,"message":{"role":"assistant","content":"..."}}]
    ' Strategy: find "content":" then extract until next "
    
    Dim contentMarker As String
    contentMarker = """content"":"""
    
    ' Find the content field in the choices array (skip system message content)
    ' First find "choices":[{"index":0,"message":{
    Dim msgStart As Long
    msgStart = InStr(json, """message"":{")
    If msgStart = 0 Then
        ' Try direct content parse
        msgStart = InStr(json, contentMarker)
        If msgStart = 0 Then
            ExtractResponseContent = ""
            Exit Function
        End If
    Else
        ' Find content after message
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
    
    ' Find the closing quote (handle escaped quotes)
    Dim depth As Long
    depth = 0
    Do While contentEnd <= Len(json)
        Dim ch As String
        ch = Mid(json, contentEnd, 1)
        
        If ch = "\"" Then
            ' Escaped quote - skip 2 chars
            contentEnd = contentEnd + 2
        ElseIf ch = "\" Then
            ' Escaped character - skip 2 chars
            contentEnd = contentEnd + 2
        ElseIf ch = """" Then
            ' Unescaped closing quote
            Exit Do
        Else
            contentEnd = contentEnd + 1
        End If
    Loop
    
    If contentEnd > contentStart Then
        ExtractResponseContent = Mid(json, contentStart, contentEnd - contentStart)
        ' Unescape JSON special characters
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
    
    ' Escape backslashes first
    result = Replace(result, "\", "\\")
    ' Escape quotes
    result = Replace(result, """", "\""")
    ' Escape control characters
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
    ' In PowerShell single-quoted strings, only ' needs escaping (doubled)
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
    
    ' Create path if it doesn't exist
    ws.RegWrite REG_PATH_ROOT, "", "REG_SZ"
    ws.RegWrite regPath, value, "REG_SZ"
End Sub

'====================================================================
' Run PowerShell script and capture output
'====================================================================
Private Function RunPowerShell(ByVal script As String) As String
    On Error GoTo PSError
    
    ' Create temporary script file
    Dim tempFile As String
    Dim tempDir As String
    tempDir = Environ("TEMP")
    tempFile = tempDir & "\ollama_config.ps1"
    
    ' Write script to file
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object
    Set file = fso.CreateTextFile(tempFile, True)
    file.Write script
    file.Close
    
    ' Execute PowerShell
    Dim ws As Object
    Set ws = CreateObject("WScript.Shell")
    Dim exec As Object
    Dim cmd As String
    
    cmd = "powershell.exe -ExecutionPolicy Bypass -File """ & tempFile & """"
    Set exec = ws.Exec(cmd)
    
    ' Wait for completion with timeout
    Dim startTime As Double
    startTime = Timer
    Do While exec.Status = 0
        DoEvents
        If Timer > startTime + 60 Then
            exec.Terminate
            Exit Do
        End If
    Loop
    
    ' Read output
    Dim output As String
    If exec.Status <> 0 Then
        output = exec.StdOut.ReadAll()
    End If
    
    ' Clean up temp file
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
    Dim olApp As Outlook.Application
    Set olApp = Outlook.Application
    Set GetActiveInspector = olApp.ActiveInspector
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
' Utility: Get email content from inspector
'====================================================================
Private Function GetEmailBody(ByVal inspector As Outlook.Inspector) As String
    On Error Resume Next
    Dim mail As Outlook.MailItem
    Set mail = inspector.CurrentItem
    If Not mail Is Nothing Then
        GetEmailBody = mail.Body
    Else
        GetEmailBody = ""
    End If
End Function

'====================================================================
' Utility: Get email subject from inspector
'====================================================================
Private Function GetEmailSubject(ByVal inspector As Outlook.Inspector) As String
    On Error Resume Next
    Dim mail As Outlook.MailItem
    Set mail = inspector.CurrentItem
    If Not mail Is Nothing Then
        GetEmailSubject = mail.Subject
    Else
        GetEmailSubject = ""
    End If
End Function

'====================================================================
' End of Module
'====================================================================
