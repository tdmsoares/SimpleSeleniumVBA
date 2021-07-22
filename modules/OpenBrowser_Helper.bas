Option Compare Database
Option Explicit


'------------------------------------------------------------------------------------------------------------------
' Coded by tdmsoares
'------------------------------------------------------------------------------------------------------------------


Function ExtractSessionIdFromServerResponse(ByVal strServerResponse As String) As String
'
Dim objRegex As New VBScript_RegExp_55.RegExp
Dim objMatchCollection As VBScript_RegExp_55.MatchCollection
With objRegex
    .Global = False
    '.Pattern = "(?=""sessionId"":"".*"")"
    .Pattern = """sessionId"":"".*"""
End With
Set objMatchCollection = objRegex.Execute(strServerResponse)
ExtractSessionIdFromServerResponse = Strings.Replace(Strings.Replace(objMatchCollection(0), """sessionId"":""", """"), """", "")
End Function

Function OpenChrome() As String
    '
    '
    Dim strServerResponse As String
    Dim objMXSML2ServerXMLHTPP As New MSXML2.ServerXMLHTTP
    Call objMXSML2ServerXMLHTPP.Open("POST", "http://localhost:9515/session")
    Call objMXSML2ServerXMLHTPP.setRequestHeader("Content-Type", "application/json; charset=utf-8")
    Call objMXSML2ServerXMLHTPP.send("{""capabilities"":{""firstMatch"":[{""browserName"":""chrome"",""goog:chromeOptions"":{""args"":[""user-data-dir=D:\\Profiles\\tdmsoares\\Desktop\\tmpChromeUserData\\User Data""],""excludeSwitches"":[""enable-automation""],""extensions"":[],""prefs"":{""credentials_enable_service"":false,""profile.default_content_setting_values.notifications"":1,""profile.default_content_settings.popups"":0,""profile.password_manager_enabled"":false}}}]},""desiredCapabilities"":{""browserName"":""chrome"",""goog:chromeOptions"":{""args"":[],""excludeSwitches"":[""enable-automation""],""extensions"":[],""prefs"":{""credentials_enable_service"":false,""profile.default_content_setting_values.notifications"":1,""profile.default_content_settings.popups"":0,""profile.password_manager_enabled"":false}}}}")
    strServerResponse = objMXSML2ServerXMLHTPP.responseText
    Debug.Print "Server Response: " & strServerResponse
    Debug.Print "SessionId: " & ExtractSessionIdFromServerResponse(strServerResponse)
    OpenChrome = ExtractSessionIdFromServerResponse(strServerResponse)
End Function