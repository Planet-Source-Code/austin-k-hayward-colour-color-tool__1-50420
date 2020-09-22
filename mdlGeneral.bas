Attribute VB_Name = "mdlGeneral"
Option Explicit


Public Sub HandleError(ByVal CurrentModule As String, ByVal CurrentProcedure As String, ByVal ErrNum As Long, ByVal ErrDescription As String)

On Error GoTo Err_Init

    MsgBox CurrentModule & " " & CurrentProcedure & ": " & ErrNum & " - " & ErrDescription

Exit Sub

Err_Init:
    MsgBox CurrentModule & " HandleError: " & Err.Number & " - " & Err.Description

End Sub

