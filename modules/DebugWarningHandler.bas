Option Explicit


'------------------------------------------------------------------------------------------------------------------
' Coded by tdmsoares
'------------------------------------------------------------------------------------------------------------------


Sub ErrHandler(ByRef ErrorObject As ErrObject, ByVal ProcedureName As String, ByVal ModuleName As String, Optional ByVal additionalMessage As String, _
                    Optional ByVal displayMSGBox As Boolean = False, Optional ByVal MsgBoxType As VbMsgBoxStyle = vbCritical)
'
'Gera uma mensagem de Erro padronizada no Debug Window
'   Opcional uma MsgBox
Dim strMessage As String
strMessage = ErrorObject.Number & " - "
If (Not EmptyTextChecker.isEmptyText(additionalMessage)) Then
    strMessage = strMessage & additionalMessage & " - "
End If
strMessage = strMessage & ErrorObject.Description & " - Module: " & ModuleName & " - Procedure: " & ProcedureName
Debug.Print strMessage
'
If (displayMSGBox) Then
    MsgBox strMessage, MsgBoxType
End If
'
End Sub

Sub OtherWarnings(ByVal ProcedureName As String, ByVal ModuleName As String, _
                    Optional ByVal Identifier As String = "Warning", Optional ByVal additionalMessage As String = "Pending", _
                    Optional ByVal displayMSGBox As Boolean = False, Optional ByVal MsgBoxType As VbMsgBoxStyle = vbCritical)
'
'Gera uma mensagem padronizada no Debug Window
'   Opcional uma MsgBox
Dim strMessage As String
If (Not EmptyTextChecker.isEmptyText(Identifier)) Then
    strMessage = Identifier & " - "
End If
If (Not EmptyTextChecker.isEmptyText(additionalMessage)) Then
    strMessage = strMessage & additionalMessage & " - "
End If
strMessage = strMessage & " - Module: " & ModuleName & " - Procedure: " & ProcedureName
Debug.Print strMessage
'
If (displayMSGBox) Then
    MsgBox strMessage, MsgBoxType
End If
'
End Sub

Sub DisplayVariableValues(ByVal ProcedureName As String, ByVal ModuleName As String, _
                            ByVal VariableName As String, ByVal VariableValue As String, _
                            Optional ByVal displayMSGBox As Boolean = False, Optional ByVal MsgBoxType As VbMsgBoxStyle = vbInformation)
'
'Gera uma mensagem padronizada no Debug Window para mostrar valores de uma determinada variável
Dim strMessage As String
strMessage = "- Module: " & ModuleName & " - Procedure: " & ProcedureName & vbCrLf & _
            "-- Variable: " & VariableName & vbCrLf & _
            "--- Value: " & VariableValue
Debug.Print strMessage
'
If (displayMSGBox) Then
    MsgBox strMessage, MsgBoxType
End If
End Sub