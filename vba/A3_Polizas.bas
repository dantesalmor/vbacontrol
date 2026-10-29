Attribute VB_Name = "A3_Polizas"
Option Explicit

Public Function Comision(ByVal prima As Double, ByVal ramo As String, _
                         Optional ByVal bono As Double = 0) As Double
    Dim tasa As Double

    ramo = LCase(Trim(ramo)) ' " auto" o "HOGAR" -> "auto", "hogar"

    Select Case ramo ' elegir la tasa según el ramo
        Case "auto": tasa = 0.08
        Case "vida": tasa = 0.12
        Case "hogar": tasa = 0.1
        Case "salud": tasa = 0.09
        Case Else: tasa = 0 ' ramo desconocido
    End Select

    If prima <= 0 Or tasa = 0 Then
        Comision = 0 ' "return 0": se asigna al nombre de la función
        Exit Function
    End If

    Comision = Application.WorksheetFunction.Round(prima * tasa + bono, 2) ' redondeo comercial
End Function

Sub ProbarComision()
    Debug.Print Comision(850, "Auto") ' 68
    Debug.Print Comision(1200, "vida", 25) ' 169
    Debug.Print Comision(500, "Viaje") ' 0
    Debug.Print Comision(-100, "Auto") ' 0
End Sub

Sub LlenarComisiones()
    Dim ws As Worksheet
    Dim ultima As Long

    Set ws = ThisWorkbook.Worksheets("A3 Polizas")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Cells(1, 6).Value = "Comisión"
    ws.Cells(1, 7).Value = "Comisión con bono"

    ws.Range("F2:F" & ultima).Formula = "=Comision(D2,C2)"
    ws.Range("G2:G" & ultima).Formula = _
        "=Comision(D2,C2,IF(E2=""Sí"",25,0))"
End Sub

Private Function EsPrimaAlta(ByVal prima As Double) As Boolean
    EsPrimaAlta = (prima >= 1000) ' la comparación ya es True o False
End Function

Private Sub PintarFila(ByVal ws As Worksheet, ByVal fila As Long, ByVal colorFila As Long)
    ws.Range("A" & fila & ":G" & fila).Interior.Color = colorFila ' "A5:G5" si fila = 5
End Sub

Sub ResaltarPrimasAltas()
    Dim ws As Worksheet
    Dim ultima As Long, fila As Long, altas As Long

    Set ws = ThisWorkbook.Worksheets("A3 Polizas")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For fila = 2 To ultima
        If EsPrimaAlta(ws.Cells(fila, 4).Value) Then
            PintarFila ws, fila, RGB(221, 235, 247)
            altas = altas + 1
        Else
            ws.Range("A" & fila & ":G" & fila).Interior.ColorIndex = xlNone ' quitar relleno
        End If
    Next fila

    Debug.Print "Pólizas de prima alta: " & altas
End Sub
