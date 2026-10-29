Attribute VB_Name = "A1_Parqueo"
Option Explicit

Sub CalcularParqueo()
    Dim ws As Worksheet ' la hoja de datos
    Dim ultima As Long, fila As Long ' última fila y contador
    Dim minutos As Double, medias As Double
    Dim cobro As Double, desc As Double, pagar As Double
    Dim total As Double, cortesias As Long ' acumuladores (arrancan en 0)

    Set ws = ThisWorkbook.Worksheets("A1 Parqueo")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' última fila con datos

    ' Encabezados
    ws.Cells(1, 6).Value = "Minutos"
    ws.Cells(1, 7).Value = "Cobro"
    ws.Cells(1, 8).Value = "Descuento"
    ws.Cells(1, 9).Value = "Valor a pagar"

    For fila = 2 To ultima ' fila 1 = encabezados
        ' Minutos: una hora es fracción del día; x 1440 = minutos
        minutos = Round((ws.Cells(fila, 4).Value _
                  - ws.Cells(fila, 3).Value) * 1440, 0)

        ' Tramos: VBA entra al primero que se cumpla
        If minutos <= 15 Then
            cobro = 0
            cortesias = cortesias + 1
        ElseIf minutos <= 60 Then
            cobro = 1
        Else
            medias = -Int(-(minutos - 60) / 30) ' bloques extra, hacia arriba
            cobro = 1 + 0.5 * medias
        End If
        If cobro > 6 Then cobro = 6 ' tope del día

        ' Descuento por sello
        If ws.Cells(fila, 5).Value = "Sí" Then
            desc = 1
        Else
            desc = 0
        End If
        pagar = Application.WorksheetFunction.Max(cobro - desc, 0) ' nunca negativo

        ' Escribir la fila
        ws.Cells(fila, 6).Value = minutos
        ws.Cells(fila, 7).Value = cobro
        ws.Cells(fila, 8).Value = desc
        ws.Cells(fila, 9).Value = pagar
        total = total + pagar
    Next fila

    Debug.Print "Tickets procesados: " & (ultima - 1)
    Debug.Print "Cortesías: " & cortesias
    Debug.Print "Total a cobrar: " & Format(total, "$#,##0.00")
End Sub

Sub BuscarTicket()
    Dim ws As Worksheet
    Dim resp As String ' lo que escribe el usuario
    Dim ultima As Long, fila As Long
    Dim filaEnc As Long ' fila encontrada (0 = no existe)

    Set ws = ThisWorkbook.Worksheets("A1 Parqueo")
    If ws.Cells(2, 9).Value = "" Then CalcularParqueo ' si falta calcular, calcular primero

    ' Pedir el dato
    resp = InputBox("Número de ticket:", "Buscar ticket")
    If StrPtr(resp) = 0 Then Exit Sub ' presionó Cancelar: no pasa nada
    resp = UCase(Trim(resp)) ' " pq-0117 " -> "PQ-0117"
    If resp = "" Then Exit Sub ' aceptó la caja vacía

    ' Buscar fila por fila
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For fila = 2 To ultima
        If UCase(Trim(ws.Cells(fila, 1).Value)) = resp Then
            filaEnc = fila
            Exit For ' ya está: dejar de buscar
        End If
    Next fila

    If filaEnc = 0 Then
        MsgBox "El ticket " & resp & " no existe.", vbExclamation, "Buscar ticket"
        Exit Sub
    End If

    ' Un solo MsgBox, un dato por línea
    With ws
        MsgBox "Ticket: " & .Cells(filaEnc, 1).Value & vbCrLf & _
               "Placa: " & .Cells(filaEnc, 2).Value & vbCrLf & _
               "Entrada: " & Format(.Cells(filaEnc, 3).Value, "h:mm AM/PM") & vbCrLf & _
               "Salida: " & Format(.Cells(filaEnc, 4).Value, "h:mm AM/PM") & vbCrLf & _
               "Minutos: " & .Cells(filaEnc, 6).Value & vbCrLf & _
               "Valor a pagar: " & Format(.Cells(filaEnc, 9).Value, "$#,##0.00"), _
               vbInformation, "Ticket encontrado"
    End With
End Sub

Sub ValidarPlacas()
    Dim ws As Worksheet
    Dim cel As Range ' cada celda del For Each
    Dim ultima As Long, malas As Long

    Set ws = ThisWorkbook.Worksheets("A1 Parqueo")
    ultima = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    For Each cel In ws.Range("B2:B" & ultima)
        If cel.Value Like "P###-###" Then ' # = un dígito
            cel.Interior.ColorIndex = xlNone ' correcta: quitar relleno viejo
        Else
            cel.Interior.Color = vbYellow
            Debug.Print ws.Cells(cel.Row, 1).Value ' ticket de esa fila
            malas = malas + 1
        End If
    Next cel

    If malas > 0 Then
        MsgBox "Hay " & malas & " placas incorrectas.", vbExclamation, "Validar placas"
    Else
        MsgBox "Todas las placas están correctas.", vbInformation, "Validar placas"
    End If
End Sub
