Attribute VB_Name = "BotonesColoresPOS"
Sub ConfigurarBotonesActiveX()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("POS")
    
    ' Botón ABRIR POS
    With ws.OLEObjects("btnAbrirPOS").Object
        .BackColor = RGB(46, 204, 113)  ' Verde
        .ForeColor = RGB(255, 255, 255)  ' Texto blanco
        .Font.Bold = True
        .Font.Size = 12
        .Caption = "ABRIR POS"
    End With
    
    ' Botón CERRAR Y GUARDAR
    With ws.OLEObjects("btnCerrarGuardar").Object
        .BackColor = RGB(231, 76, 60)   ' Rojo
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
        .Caption = "CERRAR Y GUARDAR"
    End With
    
    ' Botón BLOQUEAR
    With ws.OLEObjects("btnBloquear").Object
        .BackColor = RGB(243, 156, 18)  ' Naranja
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
        .Caption = "BLOQUEAR TODO"
    End With
    
    ' Botón DESBLOQUEAR
    With ws.OLEObjects("btnDesbloquear").Object
        .BackColor = RGB(52, 152, 219)  ' Azul
        .ForeColor = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
        .Caption = "DESBLOQUEAR"
    End With
    
    MsgBox "Colores aplicados. Ahora asigna las macros.", vbInformation
End Sub
