Attribute VB_Name = "modProteccion"
Option Explicit
Private Const PWD As String = "TuClaveFuerte123"  ' cámbiala

Public Sub ProtegerTodo(Optional ByVal silencioso As Boolean = False)
    Dim sh As Worksheet
    ' Protege la estructura del libro (oculta mover/insertar/eliminar hojas)
    On Error Resume Next
    ThisWorkbook.Protect password:=PWD, Structure:=True, Windows:=False
    On Error GoTo 0

    For Each sh In ThisWorkbook.Worksheets
        On Error Resume Next
        sh.EnableSelection = xlNoSelection ' o xlUnlockedCells si prefieres
        sh.Protect password:=PWD, _
            UserInterfaceOnly:=True, _
            AllowFiltering:=True, _
            AllowSorting:=True, _
            AllowFormattingCells:=True, _
            AllowFormattingColumns:=True, _
            AllowFormattingRows:=True
        On Error GoTo 0
    Next sh

    If Not silencioso Then MsgBox "Hojas protegidas (UIOnly).", vbInformation
End Sub

Public Sub DesprotegerTodo(Optional ByVal silencioso As Boolean = False)
    Dim sh As Worksheet
    On Error Resume Next
    ThisWorkbook.Unprotect password:=PWD
    For Each sh In ThisWorkbook.Worksheets
        sh.Unprotect password:=PWD
    Next sh
    On Error GoTo 0
    If Not silencioso Then MsgBox "Hojas desprotegidas.", vbInformation
End Sub


