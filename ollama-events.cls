VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "OllamaEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_Inspectors As Outlook.Inspectors

Public Sub Init(ByVal inspectors As Outlook.Inspectors)
    Set m_Inspectors = inspectors
End Sub

Private Sub m_Inspectors_NewInspector(ByVal Inspector As Outlook.Inspector)
    On Error GoTo HandlerError
    If Inspector.CurrentItem.Class = olMail Then
        AddToolbarButton Inspector
    End If
    Exit Sub
HandlerError:
End Sub

Private Sub AddToolbarButton(ByVal Inspector As Outlook.Inspector)
    On Error Resume Next
    Dim cmdBar As Office.CommandBar
    Dim btn As Office.CommandBarButton
    Dim existingBar As Office.CommandBar

    Set existingBar = Inspector.CommandBars("Ollama AI")
    If Not existingBar Is Nothing Then Exit Sub

    Set cmdBar = Inspector.CommandBars.Add("Ollama AI", msoBarTop, False, True)
    If cmdBar Is Nothing Then Exit Sub

    cmdBar.Visible = True

    Set btn = cmdBar.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Process with AI"
        .Tag = "OllamaAI_Process"
        .OnAction = "'OllamaAI.Ollama_ProcessWithAI'"
        .TooltipText = "Send email to Ollama AI for processing"
        .Style = msoButtonCaption
    End With

    cmdBar.Controls.Add Type:=msoControlSeparator

    Set btn = cmdBar.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Settings"
        .Tag = "OllamaAI_Settings"
        .OnAction = "'OllamaAI.Ollama_ShowConfigurationForm'"
        .TooltipText = "Configure Ollama AI settings"
        .Style = msoButtonCaption
    End With

    Set btn = cmdBar.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "About"
        .Tag = "OllamaAI_About"
        .OnAction = "'OllamaAI.Ollama_ShowAboutForm'"
        .TooltipText = "About Ollama Outlook AI"
        .Style = msoButtonCaption
    End With
End Sub
