Option Explicit

'------------------------------------------------------------------------------------------------------------------
' Coded by tdmsoares
'------------------------------------------------------------------------------------------------------------------

Function isEmptyText(ByVal Text) As Boolean
    '
    'Checks whether the Text is null or empty
    Dim isEmpty As Boolean
    isEmpty = False
    '
    If (IsNull(Text)) Then
        isEmpty = True
    ElseIf (Text = "") Then
        isEmpty = True
    ElseIf (Trim(Text)) = "" Then
        isEmpty = True
    End If
    '
    isEmptyText = isEmpty
End Function

Function AvisoPreenchimentoIncompleto(Campo, TextoMensagem, Requerido As Boolean) As Boolean
On Error GoTo Errado
'
'Lida caso haja Campo não preenchido
If (Requerido = False) Then
    Dim corrigir As VbMsgBoxResult
    corrigir = MsgBox(TextoMensagem & vbCrLf & "Deseja corrigir?", vbQuestion + vbYesNo, "Cadastro Incompleto")
    If (corrigir = vbYes) Then
        Campo.SetFocus
        AvisoPreenchimentoIncompleto = True
    End If
Else:
    MsgBox TextoMensagem & vbCrLf & "Corriga e Tente Novamente", vbExclamation, "Campo Requerido não Preenchido"
    Campo.SetFocus
    AvisoPreenchimentoIncompleto = True
End If
'
Errado:
    If (Err.Number <> 0) Then
        If (Err.Number = 2110) Then
            MsgBox "Campo Inativo", vbCritical
            Resume Next
        Else:
            Call DebugWarningHandler.ErrHandler(Err, "AvisoPreenchimentoIncompleto", "EmptyTextChecker")
            AvisoPreenchimentoIncompleto = True
            Exit Function
        End If
    End If
End Function