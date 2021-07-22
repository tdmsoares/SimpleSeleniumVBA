Option Compare Database
Option Explicit

Function ExecuteSyncScript(ByVal strWebdriverURL As String, ByVal strBrowserSessionId As String, ByVal strScript As String, Optional ByVal strValue As String, Optional ByVal strArgs As String) As String
    '
    'Executa um Script Sync
    Dim strServerResponse As String
    Dim objMXSML2ServerXMLHTPP As New MSXML2.ServerXMLHTTP
    Call objMXSML2ServerXMLHTPP.Open("POST", strWebdriverURL & "/session/" & strBrowserSessionId & "/execute/sync")
    Call objMXSML2ServerXMLHTPP.setRequestHeader("Content-Type", "application/json; charset=utf-8")
    Call objMXSML2ServerXMLHTPP.send("{""script"":""" & strScript & """,""value"": """ & strValue & """, ""args"": [""" & strArgs & """]}")
    '
    'Valor Retornado
    strServerResponse = objMXSML2ServerXMLHTPP.responseText
    Debug.Print "Execute Script - Returned Value: " & strServerResponse
End Function