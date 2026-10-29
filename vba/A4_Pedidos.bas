Attribute VB_Name = "A4_Pedidos"
Option Explicit

Sub TotalesPedidos()
    Dim ws As Worksheet
    Dim ultima As Long, n As Long, i As Long
    Dim datos As Variant ' aquí cae la tabla entera
    Dim totales() As Double ' tamaño se define con ReDim
    Dim suma As Double

    Set ws = ThisWorkbook.Worksheets("A4 Pedidos")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    datos = ws.Range("A2:F" & ultima).Value ' la hoja se lee UNA sola vez
    n = UBound(datos, 1) ' cuántas filas tiene el arreglo

    ReDim totales(1 To n, 1 To 1) ' n filas, 1 columna

    For i = 1 To n
        totales(i, 1) = Application.WorksheetFunction.Round(datos(i, 5) * datos(i, 6), 2) ' unidades x precio
        suma = suma + totales(i, 1)
    Next i

    ws.Cells(1, 7).Value = "Total"
    ws.Range("G2:G" & ultima).Value = totales ' todo de un golpe

    Debug.Print "Pedidos: " & n
    Debug.Print "Suman: " & Format(suma, "$#,##0.00")
End Sub

Sub ClientesDistintos()
    Dim ws As Worksheet
    Dim ultima As Long, i As Long, j As Long
    Dim datos As Variant
    Dim clientes As Collection
    Dim nombre As String
    Dim existe As Boolean
    Dim c As Variant

    Set ws = ThisWorkbook.Worksheets("A4 Pedidos")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    datos = ws.Range("A2:F" & ultima).Value

    Set clientes = New Collection ' lista vacía (objeto: lleva Set)

    For i = 1 To UBound(datos, 1)
        nombre = datos(i, 3)

        existe = False ' ¿ya está en la lista? recorrerla y comparar
        For j = 1 To clientes.Count
            If clientes(j) = nombre Then
                existe = True
                Exit For
            End If
        Next j

        If existe = False Then clientes.Add nombre ' agregar al final
    Next i

    Debug.Print "Clientes distintos: " & clientes.Count
    For Each c In clientes
        Debug.Print c
    Next c
End Sub

Sub ResumenPorCategoria()
    Dim ws As Worksheet, wsR As Worksheet, hoja As Worksheet
    Dim ultima As Long, i As Long, fila As Long
    Dim datos As Variant
    Dim ventas As Object, pedidos As Object ' dos diccionarios: grupo -> $ y grupo -> cantidad
    Dim cat As Variant
    Dim monto As Double, total As Double
    Dim mejorCat As String, mejorMonto As Double

    Set ws = ThisWorkbook.Worksheets("A4 Pedidos")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    datos = ws.Range("A2:F" & ultima).Value

    ' Acumular en una sola pasada
    Set ventas = CreateObject("Scripting.Dictionary")
    Set pedidos = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(datos, 1)
        cat = datos(i, 4)
        monto = Application.WorksheetFunction.Round(datos(i, 5) * datos(i, 6), 2)
        If Not ventas.Exists(cat) Then ' grupo nuevo: crearlo en 0
            ventas.Add cat, 0
            pedidos.Add cat, 0
        End If
        ventas(cat) = ventas(cat) + monto ' como d[clave] += x
        pedidos(cat) = pedidos(cat) + 1
        total = total + monto
    Next i

    ' Borrar la hoja de la ejecución anterior, sin preguntar
    Application.DisplayAlerts = False
    For Each hoja In ThisWorkbook.Worksheets
        If hoja.Name = "A4 Resumen" Then
            hoja.Delete
            Exit For
        End If
    Next hoja
    Application.DisplayAlerts = True

    ' Crear la hoja al final del libro
    Set wsR = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsR.Name = "A4 Resumen"

    ' Tabla
    wsR.Range("A1:D1").Value = Array("Categoría", "Pedidos", "Ventas ($)", "% del total")
    fila = 2
    For Each cat In ventas.Keys ' en el orden en que aparecieron
        wsR.Cells(fila, 1).Value = cat
        wsR.Cells(fila, 2).Value = pedidos(cat)
        wsR.Cells(fila, 3).Value = ventas(cat)
        wsR.Cells(fila, 4).Value = ventas(cat) / total
        If ventas(cat) > mejorMonto Then ' el mayor, a mano
            mejorMonto = ventas(cat)
            mejorCat = cat
        End If
        fila = fila + 1
    Next cat

    ' Fila Total
    wsR.Cells(fila, 1).Value = "Total"
    wsR.Cells(fila, 2).Value = UBound(datos, 1)
    wsR.Cells(fila, 3).Value = total
    wsR.Cells(fila, 4).Value = 1

    ' Formatos
    wsR.Range("A1:D1").Font.Bold = True
    wsR.Range("A" & fila & ":D" & fila).Font.Bold = True
    wsR.Range("C2:C" & fila).NumberFormat = "#,##0.00"
    wsR.Range("D2:D" & fila).NumberFormat = "0.0%"
    wsR.Columns("A:D").AutoFit

    MsgBox "Categoría que más vendió: " & mejorCat & vbCrLf & _
           "Ventas: " & Format(mejorMonto, "$#,##0.00"), vbInformation, "Resumen"
End Sub
