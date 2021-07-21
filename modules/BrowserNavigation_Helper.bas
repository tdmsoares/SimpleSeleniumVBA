Option Compare Database
Option Explicit


'------------------------------------------------------------------------------------------------------------------
' Coded by tdmsoares
'------------------------------------------------------------------------------------------------------------------


Sub NavigateTo(ByVal strWebdriverURL As String, ByVal strBrowserSessionId As String, ByVal strURL As String)
    '
    'Navega para o endereço strURL
    Dim objMXSML2ServerXMLHTPP As New MSXML2.ServerXMLHTTP
    Call objMXSML2ServerXMLHTPP.Open("POST", strWebdriverURL & "/session/" & strBrowserSessionId & "/url")
    Call objMXSML2ServerXMLHTPP.setRequestHeader("Content-Type", "application/json; charset=utf-8")
    objMXSML2ServerXMLHTPP.send ("{""url"":""" & strURL & """}")
    '
    Call GetCurrentURL(strWebdriverURL, strBrowserSessionId)
End Sub

Function GetCurrentURL(ByVal strWebdriverURL As String, ByVal strBrowserSessionId As String)
    '
    'Obtem Resposta da URL atual
    Dim strServerResponse As String
    Dim objMXSML2ServerXMLHTPP As New MSXML2.ServerXMLHTTP
    'Call objMXSML2ServerXMLHTPP.setRequestHeader("Content-Type", "application/json; charset=utf-8")
    Call objMXSML2ServerXMLHTPP.Open("GET", strWebdriverURL & "/session/" & strBrowserSessionId & "/url")
    objMXSML2ServerXMLHTPP.send
    strServerResponse = objMXSML2ServerXMLHTPP.responseText
    '
    'TODO: Get Only URL not entire response
    Debug.Print strServerResponse
End Function