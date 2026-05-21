VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "OllamaRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRibbonExtensibility

Private Function IRibbonExtensibility_GetCustomUI(ByVal RibbonID As String) As String
    IRibbonExtensibility_GetCustomUI = GetRibbonXML()
End Function

Private Function GetRibbonXML() As String
    Dim xml As String
    xml = "<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">"
    xml = xml & "<ribbon>"
    xml = xml & "<tabs>"
    xml = xml & "<tab idMso=""TabNewMailMessage"">"
    xml = xml & "<group id=""OllamaAIGroup"" label=""Ollama AI"" insertAfterMso=""GroupClipboard"">"
    xml = xml & "<button id=""OllamaProcessBtn"" label=""Process&#10;with AI"" size=""large"" onAction=""OllamaRibbon_OnProcess"" imageMso=""CreateMessage"" screentip=""Process email with Ollama AI"" supertip=""Send email content to your local Ollama model for AI assistance"" />"
    xml = xml & "<separator id=""OllamaSep1"" />"
    xml = xml & "<button id=""OllamaSettingsBtn"" label=""Settings"" size=""normal"" onAction=""OllamaRibbon_OnSettings"" imageMso=""ApplicationOptionsDialog"" screentip=""Configure Ollama AI settings"" />"
    xml = xml & "<button id=""OllamaAboutBtn"" label=""About"" size=""normal"" onAction=""OllamaRibbon_OnAbout"" imageMso=""Info"" screentip=""About Ollama Outlook AI"" />"
    xml = xml & "</group>"
    xml = xml & "</tab>"
    xml = xml & "<tab idMso=""TabReadMessage"">"
    xml = xml & "<group id=""OllamaAIGroupRead"" label=""Ollama AI"" insertAfterMso=""GroupClipboard"">"
    xml = xml & "<button id=""OllamaProcessBtnRead"" label=""Process&#10;with AI"" size=""large"" onAction=""OllamaRibbon_OnProcess"" imageMso=""CreateMessage"" screentip=""Process email with Ollama AI"" />"
    xml = xml & "<separator id=""OllamaSep1Read"" />"
    xml = xml & "<button id=""OllamaSettingsBtnRead"" label=""Settings"" size=""normal"" onAction=""OllamaRibbon_OnSettings"" imageMso=""ApplicationOptionsDialog"" />"
    xml = xml & "<button id=""OllamaAboutBtnRead"" label=""About"" size=""normal"" onAction=""OllamaRibbon_OnAbout"" imageMso=""Info"" />"
    xml = xml & "</group>"
    xml = xml & "</tab>"
    xml = xml & "</tabs>"
    xml = xml & "</ribbon>"
    xml = xml & "</customUI>"
    GetRibbonXML = xml
End Function

Public Sub OllamaRibbon_OnProcess(ByVal control As IRibbonControl)
    On Error Resume Next
    Ollama_ProcessWithAI
End Sub

Public Sub OllamaRibbon_OnSettings(ByVal control As IRibbonControl)
    On Error Resume Next
    Ollama_ShowConfigurationForm
End Sub

Public Sub OllamaRibbon_OnAbout(ByVal control As IRibbonControl)
    On Error Resume Next
    Ollama_ShowAboutForm
End Sub
