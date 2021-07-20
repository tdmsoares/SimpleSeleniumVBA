Option Explicit

'------------------------------------------------------------------------------------------------------------------
' Coded by tdmsoares
'------------------------------------------------------------------------------------------------------------------

Sub OpenChromeDriver(ByVal strPathDriver As String)
    '
    'Open ChromeWebdriver based on its path passed by strPathDriver
    Dim nIdTaskWebDriver As Integer
    nIdTaskWebDriver = VBA.Interaction.Shell(strPathDriver)
    If (nIdTaskWebDriver) Then
        Call DebugWarningHandler.OtherWarnings("OpenChromeDriver", "OpenWebdriver_Helper", , _
            vbCrLf & "ChromeWebDriver was opened" & vbCrLf & "Id Task = " & nIdTaskWebDriver, False, vbInformation)
    Else:
        Call DebugWarningHandler.OtherWarnings("OpenChromeDriver", "OpenWebdriver_Helper", , _
            vbCrLf & "ChromeWebDriver could not be opened", True, vbCritical)
    End If
End Sub