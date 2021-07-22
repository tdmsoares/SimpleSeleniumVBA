Option Compare Database
Option Explicit


'------------------------------------------------------------------------------------------------------------------
' Coded by tdmsoares
'------------------------------------------------------------------------------------------------------------------


Function FindElementByXpath(ByVal strWebdriverURL As String, ByVal strBrowserSessionId As String, ByVal strElementIdentifier As String) As String
    '
    'Encontra um elemento pelo seu XPath
    Dim strServerResponse As String
    Dim locatorStrategy As String
    locatorStrategy = "xpath"
    Dim objMXSML2ServerXMLHTPP As New MSXML2.ServerXMLHTTP
    Call objMXSML2ServerXMLHTPP.Open("POST", strWebdriverURL & "/session/" & strBrowserSessionId & "/element")
    Call objMXSML2ServerXMLHTPP.setRequestHeader("Content-Type", "application/json; charset=utf-8")
    objMXSML2ServerXMLHTPP.send ("{""using"":""" & locatorStrategy & """, ""value"": """ & strElementIdentifier & """}")
    '
    'Call GetCurrentURL(strWebdriverURL, strBrowserSessionId)
    strServerResponse = objMXSML2ServerXMLHTPP.responseText
    FindElementByXpath = ExtractElementIdFromServerResponse(strServerResponse)
    
End Function

Function ExtractElementIdFromServerResponse(ByVal strServerResponse As String) As String
'
'TODO
Dim objRegex As Object
Set objRegex = CreateObject("VBScript.Regexp")
Dim objMatchCollection As Object         'VBScript_RegExp_55.MatchCollection
With objRegex
    .Global = False
    '.Pattern = "(?=""sessionId"":"".*"")"
    .Pattern = """element.*"":"".*"""
End With
Set objMatchCollection = objRegex.Execute(strServerResponse)
With objRegex:
    .Global = False
    .Pattern = ":.*"""
End With
Set objMatchCollection = objRegex.Execute(objMatchCollection.Item(0))
With objRegex:
    .Global = False
    .Pattern = """\b.*"""
End With
Set objMatchCollection = objRegex.Execute(objMatchCollection.Item(0))
ExtractElementIdFromServerResponse = objMatchCollection.Item(0)
End Function

Sub SendTextKeys(ByVal strWebdriverURL As String, ByVal strBrowserSessionId As String, ByVal strElementId As String, ByVal strTextToSend As String)
    '
    '
    Dim strServerResponse As String
    Dim locatorStrategy As String
    locatorStrategy = "xpath"
    Dim objMXSML2ServerXMLHTPP As New MSXML2.ServerXMLHTTP
    Call objMXSML2ServerXMLHTPP.Open("POST", strWebdriverURL & "/session/" & strBrowserSessionId & "/element/" & strElementId & "/value")
    Call objMXSML2ServerXMLHTPP.setRequestHeader("Content-Type", "application/json; charset=utf-8")
    objMXSML2ServerXMLHTPP.send ("{""text"": """ & strTextToSend & """}")
    '
    '{"value":null} = success
    strServerResponse = objMXSML2ServerXMLHTPP.responseText
    If (strServerResponse = "{""value"":null}") Then
        Debug.Print "success"
    End If
    
End Sub

Sub Click(ByVal strWebdriverURL As String, ByVal strBrowserSessionId As String, ByVal strElementId As String)
    '
    '
    Dim strServerResponse As String
    Dim locatorStrategy As String
    locatorStrategy = "xpath"
    Dim objMXSML2ServerXMLHTPP As New MSXML2.ServerXMLHTTP
    Call objMXSML2ServerXMLHTPP.Open("POST", strWebdriverURL & "/session/" & strBrowserSessionId & "/element/" & strElementId & "/click")
    'Call objMXSML2ServerXMLHTPP.setRequestHeader("Content-Type", "application/json; charset=utf-8")
    objMXSML2ServerXMLHTPP.send ("{}")
    '
    '{"value":null} = success
    strServerResponse = objMXSML2ServerXMLHTPP.responseText
    If (strServerResponse = "{""value"":null}") Then
        Debug.Print "success"
    End If
    
End Sub