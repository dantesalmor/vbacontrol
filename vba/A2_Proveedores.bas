Attribute VB_Name = "A2_Proveedores"
Option Explicit

Sub LimpiarContactos()
    Dim ws As Worksheet
    Dim ultima As Long, fila As Long
    Dim contacto As String, correo As String
    Dim arrobas As Long ' cuántas @ tiene el correo

    Set ws = ThisWorkbook.Worksheets("A2 Proveedores")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Cells(1, 7).Value = "Contacto limpio"
    ws.Cells(1, 8).Value = "Correo limpio"
    ws.Cells(1, 9).Value = "Dominio"

    For fila = 2 To ultima
        ' Contacto: Trim de Excel (también quita dobles del medio) y luego Proper
        contacto = ws.Cells(fila, 2).Value
        contacto = Application.WorksheetFunction.Trim(contacto)
        contacto = Application.WorksheetFunction.Proper(contacto)
        ws.Cells(fila, 7).Value = contacto

        ' Correo: minúsculas y sin espacios alrededor
        correo = LCase(Trim(ws.Cells(fila, 3).Value))
        ws.Cells(fila, 8).Value = correo

        ' Contar @: largo total menos largo sin las @
        arrobas = Len(correo) - Len(Replace(correo, "@", ""))

        If arrobas = 1 Then
            ws.Cells(fila, 9).Value = Mid(correo, InStr(correo, "@") + 1) ' todo lo que va después de la @
            ws.Cells(fila, 9).Interior.ColorIndex = xlNone
        Else
            ws.Cells(fila, 9).Value = "Revisar"
            ws.Cells(fila, 9).Interior.Color = vbYellow
        End If
    Next fila
End Sub

Sub VencimientosProveedores()
    Dim ws As Worksheet
    Dim ultima As Long, fila As Long
    Dim corte As Date, venc As Date
    Dim atraso As Long
    Dim nVencidas As Long, montoVencido As Double

    Set ws = ThisWorkbook.Worksheets("A2 Proveedores")
    corte = ThisWorkbook.Worksheets("A2 Parámetros").Range("B1").Value ' se lee de la hoja, no se escribe en el código
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Cells(1, 10).Value = "Vencimiento"
    ws.Cells(1, 11).Value = "Días atraso"
    ws.Cells(1, 12).Value = "Estado"

    For fila = 2 To ultima
        ' Una fecha es un número: fecha + días = otra fecha
        venc = ws.Cells(fila, 4).Value + ws.Cells(fila, 5).Value

        ' Weekday: 1 = domingo ... 7 = sábado
        If Weekday(venc) = 7 Then
            venc = venc + 2 ' sábado -> lunes
        ElseIf Weekday(venc) = 1 Then
            venc = venc + 1 ' domingo -> lunes
        End If
        ws.Cells(fila, 10).Value = venc
        ws.Cells(fila, 10).NumberFormat = "dd/mm/yyyy" ' si no, se ve como número

        ' Atraso: fecha - fecha = días; negativo = aún no vence
        atraso = corte - venc
        If atraso < 0 Then atraso = 0
        ws.Cells(fila, 11).Value = atraso

        ' Estado por tramos
        If atraso = 0 Then
            ws.Cells(fila, 12).Value = "Al día"
            ws.Cells(fila, 12).Interior.ColorIndex = xlNone
        ElseIf atraso <= 30 Then
            ws.Cells(fila, 12).Value = "Vencida 1-30"
            ws.Cells(fila, 12).Interior.Color = vbYellow
        Else
            ws.Cells(fila, 12).Value = "Vencida +30"
            ws.Cells(fila, 12).Interior.Color = RGB(255, 199, 206) ' rojo claro
        End If

        If atraso > 0 Then
            nVencidas = nVencidas + 1
            montoVencido = montoVencido + ws.Cells(fila, 6).Value
        End If
    Next fila

    MsgBox "Facturas vencidas: " & nVencidas & vbCrLf & _
           "Monto vencido: " & Format(montoVencido, "$#,##0.00") & vbCrLf & _
           "Fecha de corte: " & Format(corte, "dd/mm/yyyy"), _
           vbInformation, "Vencimientos"
End Sub
