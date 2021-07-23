Option Compare Database
Option Explicit


'------------------------------------------------------------------------------------------------------------------
' Coded by tdmsoares
'------------------------------------------------------------------------------------------------------------------


Function ExtractSessionIdFromServerResponse(ByVal strServerResponse As String) As String
'
Dim objRegex As Object          'VBScript_RegExp_55.RegExp
Set objRegex = CreateObject("VBScript.Regexp")
Dim objMatchCollection As Object      'VBScript_RegExp_55.MatchCollection
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
    'Open Chrome Browser using an opened ChromeDriver
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

Function OpenFirefox() As String
    '
    '
    Dim strServerResponse As String
    Dim objMXSML2ServerXMLHTPP As New MSXML2.ServerXMLHTTP
    Call objMXSML2ServerXMLHTPP.Open("POST", "http://127.0.0.1:4444/session")
    Call objMXSML2ServerXMLHTPP.setRequestHeader("Content-Type", "application/json; charset=utf-8")
    'Call objMXSML2ServerXMLHTPP.send("{""capabilities"":{""firstMatch"":[{""browserName"":""chrome"",""goog:chromeOptions"":{""args"":[""user-data-dir=D:\\Profiles\\tdmsoares\\Desktop\\tmpChromeUserData\\User Data""],""excludeSwitches"":[""enable-automation""],""extensions"":[],""prefs"":{""credentials_enable_service"":false,""profile.default_content_setting_values.notifications"":1,""profile.default_content_settings.popups"":0,""profile.password_manager_enabled"":false}}}]},""desiredCapabilities"":{""browserName"":""chrome"",""goog:chromeOptions"":{""args"":[],""excludeSwitches"":[""enable-automation""],""extensions"":[],""prefs"":{""credentials_enable_service"":false,""profile.default_content_setting_values.notifications"":1,""profile.default_content_settings.popups"":0,""profile.password_manager_enabled"":false}}}}")
    'Call objMXSML2ServerXMLHTPP.send("{""capabilities"":{""alwaysMatch"":{""browserName"":""firefox""}}}")
    Call objMXSML2ServerXMLHTPP.send("{""capabilities"":{}}")
    strServerResponse = objMXSML2ServerXMLHTPP.responseText
    Debug.Print "Server Response: " & strServerResponse
    Debug.Print "SessionId: " & ExtractSessionIdFromServerResponse(strServerResponse)
    OpenFirefox = ExtractSessionIdFromServerResponse(strServerResponse)
End Function

Sub CloseBrowser(ByVal strWebdriverURL As String, ByVal strBrowserSessionId As String)
    '
    'Fecha o Browser identificado pelo strBrowserSessionId
    Dim strServerResponse As String
    Dim objMXSML2ServerXMLHTPP As New MSXML2.ServerXMLHTTP
    'Call objMXSML2ServerXMLHTPP.setRequestHeader("Content-Type", "application/json; charset=utf-8")
    Call objMXSML2ServerXMLHTPP.Open("DELETE", strWebdriverURL & "/session/" & strBrowserSessionId)
    objMXSML2ServerXMLHTPP.send
    strServerResponse = objMXSML2ServerXMLHTPP.responseText
    '
    'TODO: Get Only URL not entire response
    Debug.Print strServerResponse
End Sub