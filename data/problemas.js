/* Los 4 problemas de la ayudantía del Control 2.

   CÓMO SE ESCRIBEN LAS PLANTILLAS
   {{CLAVE}}    se reemplaza por el valor del adaptador.
   {{CLAVE#}}   en parámetros de columna: el número de la columna (C -> 3).
   '1 texto     comentario corto (se ve en "Cortos" y "Completos").
   '2 texto     comentario largo (solo en "Completos").
   '! texto     trampa: la línea se marca en rojo; el texto se ve al pasar
                el cursor y como comentario en "Completos".
   Los marcadores en línea van después del código, con un espacio antes.
   Si una línea tiene comentario y trampa, la trampa va al final. */

window.PROBLEMAS = [

/* =====================================================================
   A1  PARQUEO
   ===================================================================== */
{
id:"A1", titulo:"Parqueo del centro comercial", modulo:"A1_Parqueo",
resumen:"Cobro por tramos de 30 tickets, búsqueda de un ticket con InputBox y validación de placas con Like.",
pide:[
 ["CalcularParqueo","Minutos, cobro, descuento y valor a pagar en F:I. Imprime en Immediate tickets, cortesías y total."],
 ["BuscarTicket","Pide un ticket (acepta minúsculas y espacios). Si cancela, nada; si no existe, avisa; si existe, un MsgBox con 6 datos."],
 ["ValidarPlacas","Recorre las placas con For Each, pinta de amarillo las incorrectas, imprime su ticket y avisa cuántas son."]
],
params:[
 {g:"Hoja", k:"HOJA", l:"Hoja de datos", v:"A1 Parqueo", t:"txt"},
 {g:"Columnas de datos", k:"C_TICKET", l:"Ticket", v:"A", t:"col"},
 {g:"Columnas de datos", k:"C_PLACA", l:"Placa", v:"B", t:"col"},
 {g:"Columnas de datos", k:"C_ENT", l:"Entrada", v:"C", t:"col"},
 {g:"Columnas de datos", k:"C_SAL", l:"Salida", v:"D", t:"col"},
 {g:"Columnas de datos", k:"C_SELLO", l:"Sello", v:"E", t:"col"},
 {g:"Columnas de resultado", k:"C_MIN", l:"Minutos", v:"F", t:"col"},
 {g:"Columnas de resultado", k:"C_COBRO", l:"Cobro", v:"G", t:"col"},
 {g:"Columnas de resultado", k:"C_DESC", l:"Descuento", v:"H", t:"col"},
 {g:"Columnas de resultado", k:"C_PAGAR", l:"Valor a pagar", v:"I", t:"col"},
 {g:"Tarifa", k:"CORTESIA", l:"Gratis hasta (min)", v:"15", t:"num"},
 {g:"Tarifa", k:"HORA", l:"Primer tramo (min)", v:"60", t:"num"},
 {g:"Tarifa", k:"PRIMERA", l:"Precio primer tramo", v:"1", t:"num"},
 {g:"Tarifa", k:"BLOQUE", l:"Bloque adicional (min)", v:"30", t:"num"},
 {g:"Tarifa", k:"MEDIA", l:"Precio por bloque", v:"0.5", t:"num"},
 {g:"Tarifa", k:"TOPE", l:"Cobro máximo", v:"6", t:"num"},
 {g:"Tarifa", k:"DESC", l:"Descuento con sello", v:"1", t:"num"},
 {g:"Textos", k:"SI", l:"Texto de \"sí\" en Sello", v:"Sí", t:"txt"},
 {g:"Textos", k:"PATRON", l:"Patrón de placa", v:"P###-###", t:"txt"}
],
macros:[
{n:"CalcularParqueo", kw:"tramos tarifa cobro minutos horas tope descuento cortesia recorrer filas", d:"Llena las 4 columnas de resultado e imprime el resumen.", code:
`Sub CalcularParqueo()
    '2 Calcula minutos, cobro, descuento y valor a pagar de cada ticket. Resultado en Immediate (Ctrl+G).
    Dim ws As Worksheet '1 la hoja de datos
    Dim ultima As Long, fila As Long '1 última fila y contador
    Dim minutos As Double, medias As Double
    Dim cobro As Double, desc As Double, pagar As Double
    Dim total As Double, cortesias As Long '1 acumuladores (arrancan en 0)

    Set ws = ThisWorkbook.Worksheets("{{HOJA}}") '! una hoja es un objeto: lleva Set
    ultima = ws.Cells(ws.Rows.Count, "{{C_TICKET}}").End(xlUp).Row '1 última fila con datos

    '1 Encabezados
    ws.Cells(1, {{C_MIN#}}).Value = "Minutos"
    ws.Cells(1, {{C_COBRO#}}).Value = "Cobro"
    ws.Cells(1, {{C_DESC#}}).Value = "Descuento"
    ws.Cells(1, {{C_PAGAR#}}).Value = "Valor a pagar"

    For fila = 2 To ultima '1 fila 1 = encabezados
        '1 Minutos: una hora es fracción del día; x 1440 = minutos
        minutos = Round((ws.Cells(fila, {{C_SAL#}}).Value _
                  - ws.Cells(fila, {{C_ENT#}}).Value) * 1440, 0)

        '1 Tramos: VBA entra al primero que se cumpla
        If minutos <= {{CORTESIA}} Then
            cobro = 0
            cortesias = cortesias + 1 '! cortesía se cuenta por tiempo, no por pagar = 0
        ElseIf minutos <= {{HORA}} Then
            cobro = {{PRIMERA}}
        Else
            '2 "Cada bloque adicional o fracción" = redondear hacia arriba. -Int(-x) es el techo de x.
            medias = -Int(-(minutos - {{HORA}}) / {{BLOQUE}}) '1 bloques extra, hacia arriba '! Round aquí da mal: 61 min daría 0 bloques
            cobro = {{PRIMERA}} + {{MEDIA}} * medias
        End If
        If cobro > {{TOPE}} Then cobro = {{TOPE}} '1 tope del día

        '1 Descuento por sello
        If ws.Cells(fila, {{C_SELLO#}}).Value = "{{SI}}" Then '! texto exacto: "Si" sin tilde no coincide
            desc = {{DESC}}
        Else
            desc = 0
        End If
        pagar = Application.WorksheetFunction.Max(cobro - desc, 0) '1 nunca negativo

        '1 Escribir la fila
        ws.Cells(fila, {{C_MIN#}}).Value = minutos
        ws.Cells(fila, {{C_COBRO#}}).Value = cobro
        ws.Cells(fila, {{C_DESC#}}).Value = desc
        ws.Cells(fila, {{C_PAGAR#}}).Value = pagar
        total = total + pagar
    Next fila

    Debug.Print "Tickets procesados: " & (ultima - 1) '! ultima es la fila, no la cantidad
    Debug.Print "Cortesías: " & cortesias
    Debug.Print "Total a cobrar: " & Format(total, "$#,##0.00")
End Sub`},

{n:"BuscarTicket", kw:"buscar codigo pedir dato inputbox cancelar no existe mostrar mensaje", d:"Pide un ticket, lo busca y muestra sus datos.", code:
`Sub BuscarTicket()
    '2 Pide un ticket con InputBox, lo busca fila por fila y muestra sus datos en un solo MsgBox.
    Dim ws As Worksheet
    Dim resp As String '1 lo que escribe el usuario
    Dim ultima As Long, fila As Long
    Dim filaEnc As Long '1 fila encontrada (0 = no existe)

    Set ws = ThisWorkbook.Worksheets("{{HOJA}}")
    If ws.Cells(2, {{C_PAGAR#}}).Value = "" Then CalcularParqueo '1 si falta calcular, calcular primero

    '1 Pedir el dato
    resp = InputBox("Número de ticket:", "Buscar ticket")
    If StrPtr(resp) = 0 Then Exit Sub '1 presionó Cancelar: no pasa nada
    resp = UCase(Trim(resp)) '1 " pq-0117 " -> "PQ-0117"
    If resp = "" Then Exit Sub '1 aceptó la caja vacía

    '1 Buscar fila por fila
    ultima = ws.Cells(ws.Rows.Count, "{{C_TICKET}}").End(xlUp).Row
    For fila = 2 To ultima
        If UCase(Trim(ws.Cells(fila, {{C_TICKET#}}).Value)) = resp Then '! normalizar los dos lados antes de comparar
            filaEnc = fila
            Exit For '1 ya está: dejar de buscar
        End If
    Next fila

    If filaEnc = 0 Then
        MsgBox "El ticket " & resp & " no existe.", vbExclamation, "Buscar ticket"
        Exit Sub
    End If

    '1 Un solo MsgBox, un dato por línea
    With ws
        MsgBox "Ticket: " & .Cells(filaEnc, {{C_TICKET#}}).Value & vbCrLf & _
               "Placa: " & .Cells(filaEnc, {{C_PLACA#}}).Value & vbCrLf & _
               "Entrada: " & Format(.Cells(filaEnc, {{C_ENT#}}).Value, "h:mm AM/PM") & vbCrLf & _
               "Salida: " & Format(.Cells(filaEnc, {{C_SAL#}}).Value, "h:mm AM/PM") & vbCrLf & _
               "Minutos: " & .Cells(filaEnc, {{C_MIN#}}).Value & vbCrLf & _
               "Valor a pagar: " & Format(.Cells(filaEnc, {{C_PAGAR#}}).Value, "$#,##0.00"), _
               vbInformation, "Ticket encontrado"
    End With
End Sub`},

{n:"ValidarPlacas", kw:"validar formato patron like resaltar amarillo for each", d:"Marca las placas que no cumplen el patrón.", code:
`Sub ValidarPlacas()
    '2 Compara cada placa con el patrón. Incorrecta: amarillo e imprime su ticket. Correcta: sin relleno.
    Dim ws As Worksheet
    Dim cel As Range '1 cada celda del For Each
    Dim ultima As Long, malas As Long

    Set ws = ThisWorkbook.Worksheets("{{HOJA}}")
    ultima = ws.Cells(ws.Rows.Count, "{{C_PLACA}}").End(xlUp).Row

    For Each cel In ws.Range("{{C_PLACA}}2:{{C_PLACA}}" & ultima)
        If cel.Value Like "{{PATRON}}" Then '1 # = un dígito '! Like distingue mayúsculas: "p245-118" no cumple
            cel.Interior.ColorIndex = xlNone '1 correcta: quitar relleno viejo
        Else
            cel.Interior.Color = vbYellow
            Debug.Print ws.Cells(cel.Row, {{C_TICKET#}}).Value '1 ticket de esa fila
            malas = malas + 1
        End If
    Next cel

    If malas > 0 Then
        MsgBox "Hay " & malas & " placas incorrectas.", vbExclamation, "Validar placas"
    Else
        MsgBox "Todas las placas están correctas.", vbInformation, "Validar placas"
    End If
End Sub`}
],
comprobar:[
 "CalcularParqueo: 30 tickets, 3 cortesías y $59.00 por cobrar.",
 "PQ-0120 (15 min) es cortesía; PQ-0121 (16 min) paga $1.00; PQ-0105 (61 min) paga $1.50.",
 "PQ-0114 (335 min) llega al máximo de $6.00; PQ-0110 (420 min, con sello) paga $5.00.",
 "BuscarTicket con \" pq-0117 \": 45 minutos y $0.00. Con PQ-0999 aparece el aviso. Con Cancelar no pasa nada.",
 "ValidarPlacas: 4 incorrectas (PQ-0105, PQ-0112, PQ-0120 y PQ-0127)."
],
variantes:[
 {t:"Descuento en porcentaje", d:"\"Con sello se descuenta el 20% del cobro.\"",
  a:`            desc = {{DESC}}`, b:`            desc = cobro * 0.2 '1 20% del cobro`},
 {t:"Cobro por hora completa o fracción", d:"\"$1.00 por cada hora o fracción\", sin bloques de media hora.",
  a:`        ElseIf minutos <= {{HORA}} Then
            cobro = {{PRIMERA}}
        Else
            medias = -Int(-(minutos - {{HORA}}) / {{BLOQUE}})
            cobro = {{PRIMERA}} + {{MEDIA}} * medias
        End If`,
  b:`        Else
            cobro = 1 * -Int(-minutos / 60) '1 horas hacia arriba x tarifa
        End If`},
 {t:"Resumen en MsgBox en vez de Immediate", d:"\"Al final, un mensaje con tickets, cortesías y total.\"",
  a:`    Debug.Print "Tickets procesados: " & (ultima - 1)
    Debug.Print "Cortesías: " & cortesias
    Debug.Print "Total a cobrar: " & Format(total, "$#,##0.00")`,
  b:`    MsgBox "Tickets procesados: " & (ultima - 1) & vbCrLf & _
           "Cortesías: " & cortesias & vbCrLf & _
           "Total a cobrar: " & Format(total, "$#,##0.00"), _
           vbInformation, "Parqueo"`},
 {t:"Buscar por placa en vez de ticket", d:"\"Pedir la placa y mostrar el ticket.\" Se cambia la columna donde se busca y el texto del InputBox.",
  a:`    resp = InputBox("Número de ticket:", "Buscar ticket")
…
        If UCase(Trim(ws.Cells(fila, {{C_TICKET#}}).Value)) = resp Then`,
  b:`    resp = InputBox("Placa:", "Buscar placa")
…
        If UCase(Trim(ws.Cells(fila, {{C_PLACA#}}).Value)) = resp Then`},
 {t:"Pintar toda la fila en vez de la celda", d:"\"Resaltar en amarillo la fila del ticket con placa incorrecta.\"",
  a:`            cel.Interior.Color = vbYellow`,
  b:`            ws.Range("A" & cel.Row & ":{{C_PAGAR}}" & cel.Row).Interior.Color = vbYellow`}
]
},

/* =====================================================================
   A2  PROVEEDORES
   ===================================================================== */
{
id:"A2", titulo:"Facturas de proveedores", modulo:"A2_Proveedores",
resumen:"Limpieza de nombres y correos con funciones de texto, y vencimientos con fechas y fin de semana.",
pide:[
 ["LimpiarContactos","G: contacto limpio y en tipo nombre propio. H: correo en minúsculas. I: dominio, o \"Revisar\" en amarillo si no tiene exactamente una @."],
 ["VencimientosProveedores","J: vencimiento (pasa al lunes si cae en fin de semana). K: días de atraso. L: estado con color. MsgBox con vencidas, monto y fecha de corte."]
],
params:[
 {g:"Hojas", k:"HOJA", l:"Hoja de facturas", v:"A2 Proveedores", t:"txt"},
 {g:"Hojas", k:"HOJA_PAR", l:"Hoja de parámetros", v:"A2 Parámetros", t:"txt"},
 {g:"Hojas", k:"CELDA_CORTE", l:"Celda de la fecha de corte", v:"B1", t:"txt"},
 {g:"Columnas de datos", k:"C_FACT", l:"Factura", v:"A", t:"col"},
 {g:"Columnas de datos", k:"C_CONT", l:"Contacto", v:"B", t:"col"},
 {g:"Columnas de datos", k:"C_CORREO", l:"Correo", v:"C", t:"col"},
 {g:"Columnas de datos", k:"C_FECHA", l:"Fecha factura", v:"D", t:"col"},
 {g:"Columnas de datos", k:"C_DIAS", l:"Días de crédito", v:"E", t:"col"},
 {g:"Columnas de datos", k:"C_MONTO", l:"Monto", v:"F", t:"col"},
 {g:"Columnas de resultado", k:"C_CONT_L", l:"Contacto limpio", v:"G", t:"col"},
 {g:"Columnas de resultado", k:"C_CORREO_L", l:"Correo limpio", v:"H", t:"col"},
 {g:"Columnas de resultado", k:"C_DOM", l:"Dominio", v:"I", t:"col"},
 {g:"Columnas de resultado", k:"C_VENC", l:"Vencimiento", v:"J", t:"col"},
 {g:"Columnas de resultado", k:"C_ATR", l:"Días de atraso", v:"K", t:"col"},
 {g:"Columnas de resultado", k:"C_EST", l:"Estado", v:"L", t:"col"},
 {g:"Reglas", k:"TRAMO", l:"Límite del primer tramo (días)", v:"30", t:"num"}
],
macros:[
{n:"LimpiarContactos", kw:"limpiar texto espacios mayusculas minusculas nombre propio correo arroba dominio", d:"Nombre limpio, correo en minúsculas y dominio.", code:
`Sub LimpiarContactos()
    '2 Limpia el contacto (G), el correo (H) y saca el dominio (I). Correos sin exactamente una @: "Revisar".
    Dim ws As Worksheet
    Dim ultima As Long, fila As Long
    Dim contacto As String, correo As String
    Dim arrobas As Long '1 cuántas @ tiene el correo

    Set ws = ThisWorkbook.Worksheets("{{HOJA}}")
    ultima = ws.Cells(ws.Rows.Count, "{{C_FACT}}").End(xlUp).Row

    ws.Cells(1, {{C_CONT_L#}}).Value = "Contacto limpio"
    ws.Cells(1, {{C_CORREO_L#}}).Value = "Correo limpio"
    ws.Cells(1, {{C_DOM#}}).Value = "Dominio"

    For fila = 2 To ultima
        '1 Contacto: Trim de Excel (también quita dobles del medio) y luego Proper
        contacto = ws.Cells(fila, {{C_CONT#}}).Value
        contacto = Application.WorksheetFunction.Trim(contacto) '! el Trim de VBA solo no quita los espacios del medio
        contacto = Application.WorksheetFunction.Proper(contacto)
        ws.Cells(fila, {{C_CONT_L#}}).Value = contacto

        '1 Correo: minúsculas y sin espacios alrededor
        correo = LCase(Trim(ws.Cells(fila, {{C_CORREO#}}).Value))
        ws.Cells(fila, {{C_CORREO_L#}}).Value = correo

        '1 Contar @: largo total menos largo sin las @
        arrobas = Len(correo) - Len(Replace(correo, "@", ""))

        If arrobas = 1 Then
            ws.Cells(fila, {{C_DOM#}}).Value = Mid(correo, InStr(correo, "@") + 1) '1 todo lo que va después de la @
            ws.Cells(fila, {{C_DOM#}}).Interior.ColorIndex = xlNone
        Else
            ws.Cells(fila, {{C_DOM#}}).Value = "Revisar"
            ws.Cells(fila, {{C_DOM#}}).Interior.Color = vbYellow
        End If
    Next fila
End Sub`},

{n:"VencimientosProveedores", kw:"vencimiento fecha fin de semana sabado domingo lunes atraso dias vencida corte colores", d:"Vencimiento con fin de semana, atraso y estado.", code:
`Sub VencimientosProveedores()
    '2 Vencimiento = fecha + días de crédito; si cae sábado o domingo, pasa al lunes. Atraso a la fecha de corte.
    Dim ws As Worksheet
    Dim ultima As Long, fila As Long
    Dim corte As Date, venc As Date
    Dim atraso As Long
    Dim nVencidas As Long, montoVencido As Double

    Set ws = ThisWorkbook.Worksheets("{{HOJA}}")
    corte = ThisWorkbook.Worksheets("{{HOJA_PAR}}").Range("{{CELDA_CORTE}}").Value '1 se lee de la hoja, no se escribe en el código
    ultima = ws.Cells(ws.Rows.Count, "{{C_FACT}}").End(xlUp).Row

    ws.Cells(1, {{C_VENC#}}).Value = "Vencimiento"
    ws.Cells(1, {{C_ATR#}}).Value = "Días atraso"
    ws.Cells(1, {{C_EST#}}).Value = "Estado"

    For fila = 2 To ultima
        '1 Una fecha es un número: fecha + días = otra fecha
        venc = ws.Cells(fila, {{C_FECHA#}}).Value + ws.Cells(fila, {{C_DIAS#}}).Value

        '1 Weekday: 1 = domingo ... 7 = sábado
        If Weekday(venc) = 7 Then
            venc = venc + 2 '1 sábado -> lunes
        ElseIf Weekday(venc) = 1 Then
            venc = venc + 1 '1 domingo -> lunes
        End If
        ws.Cells(fila, {{C_VENC#}}).Value = venc
        ws.Cells(fila, {{C_VENC#}}).NumberFormat = "dd/mm/yyyy" '1 si no, se ve como número

        '1 Atraso: fecha - fecha = días; negativo = aún no vence
        atraso = corte - venc
        If atraso < 0 Then atraso = 0
        ws.Cells(fila, {{C_ATR#}}).Value = atraso

        '1 Estado por tramos
        If atraso = 0 Then
            ws.Cells(fila, {{C_EST#}}).Value = "Al día"
            ws.Cells(fila, {{C_EST#}}).Interior.ColorIndex = xlNone
        ElseIf atraso <= {{TRAMO}} Then
            ws.Cells(fila, {{C_EST#}}).Value = "Vencida 1-{{TRAMO}}"
            ws.Cells(fila, {{C_EST#}}).Interior.Color = vbYellow
        Else
            ws.Cells(fila, {{C_EST#}}).Value = "Vencida +{{TRAMO}}"
            ws.Cells(fila, {{C_EST#}}).Interior.Color = RGB(255, 199, 206) '1 rojo claro
        End If

        If atraso > 0 Then
            nVencidas = nVencidas + 1
            montoVencido = montoVencido + ws.Cells(fila, {{C_MONTO#}}).Value
        End If
    Next fila

    MsgBox "Facturas vencidas: " & nVencidas & vbCrLf & _
           "Monto vencido: " & Format(montoVencido, "$#,##0.00") & vbCrLf & _
           "Fecha de corte: " & Format(corte, "dd/mm/yyyy"), _
           vbInformation, "Vencimientos"
End Sub`}
],
comprobar:[
 "Fila 2: \"  maría josé   hernández \" queda María José Hernández, ventas@laceiba.com.sv y laceiba.com.sv.",
 "2 correos para revisar: CCF-0475 (sin @) y CCF-0511 (con dos @).",
 "CCF-0451 vence el domingo 04/10/2026, pasa al lunes 05/10/2026 y lleva 25 días de atraso.",
 "CCF-0484 vence el sábado 31/10/2026, pasa al lunes 02/11/2026 y está al día.",
 "8 al día, 10 vencidas 1-30 y 4 vencidas +30: 14 vencidas por $19,517.79."
],
variantes:[
 {t:"Pagar el viernes anterior", d:"\"Si vence en fin de semana, se paga el viernes anterior.\"",
  a:`        If Weekday(venc) = 7 Then
            venc = venc + 2
        ElseIf Weekday(venc) = 1 Then
            venc = venc + 1
        End If`,
  b:`        If Weekday(venc) = 7 Then
            venc = venc - 1 '1 sábado -> viernes
        ElseIf Weekday(venc) = 1 Then
            venc = venc - 2 '1 domingo -> viernes
        End If`},
 {t:"Tres tramos de atraso", d:"\"Al día, 1-15, 16-45 y +45.\" Se agrega un ElseIf por tramo.",
  a:`        ElseIf atraso <= {{TRAMO}} Then
            ws.Cells(fila, {{C_EST#}}).Value = "Vencida 1-{{TRAMO}}"
            ws.Cells(fila, {{C_EST#}}).Interior.Color = vbYellow
        Else`,
  b:`        ElseIf atraso <= 15 Then
            ws.Cells(fila, {{C_EST#}}).Value = "Vencida 1-15"
            ws.Cells(fila, {{C_EST#}}).Interior.Color = vbYellow
        ElseIf atraso <= 45 Then
            ws.Cells(fila, {{C_EST#}}).Value = "Vencida 16-45"
            ws.Cells(fila, {{C_EST#}}).Interior.Color = RGB(255, 221, 170)
        Else`},
 {t:"Validar el correo completo", d:"\"Válido si tiene una sola @, un punto después de la @ y ningún espacio.\"",
  a:`        If arrobas = 1 Then`,
  b:`        If arrobas = 1 And _
           InStr(InStr(correo, "@"), correo, ".") > 0 And _
           InStr(correo, " ") = 0 Then '1 InStr con inicio: busca el punto después de la @`},
 {t:"Sacar el usuario (antes de la @)", d:"\"En la columna I, el usuario del correo.\"",
  a:`            ws.Cells(fila, {{C_DOM#}}).Value = Mid(correo, InStr(correo, "@") + 1)`,
  b:`            ws.Cells(fila, {{C_DOM#}}).Value = Left(correo, InStr(correo, "@") - 1)`},
 {t:"Colorear la fila entera", d:"\"Resaltar la fila de cada factura vencida.\"",
  a:`            ws.Cells(fila, {{C_EST#}}).Interior.Color = vbYellow`,
  b:`            ws.Range("A" & fila & ":{{C_EST}}" & fila).Interior.Color = vbYellow`}
]
},

/* =====================================================================
   A3  PÓLIZAS
   ===================================================================== */
{
id:"A3", titulo:"Comisiones de una aseguradora", modulo:"A3_Polizas",
resumen:"Una función propia que se usa desde VBA y como fórmula, Optional, y procedimientos Public y Private.",
nota:"Las 6 piezas van en el MISMO módulo: las privadas solo se pueden usar desde su propio módulo. Usá \"Copiar módulo completo\".",
pide:[
 ["Comision","Función: prima × tasa + bono (opcional), a 2 decimales. Ramo desconocido o prima ≤ 0: devuelve 0."],
 ["ProbarComision","Macro de prueba con los 4 casos del enunciado."],
 ["LlenarComisiones","Una instrucción por columna: fórmula en F y fórmula con bono en G."],
 ["EsPrimaAlta","Función privada: True si la prima es de 1,000 o más."],
 ["PintarFila","Sub privado: rellena las columnas A a G de una fila."],
 ["ResaltarPrimasAltas","Macro pública: pinta de celeste las primas altas, quita el relleno a las demás e imprime cuántas son."]
],
params:[
 {g:"Hoja", k:"HOJA", l:"Hoja de pólizas", v:"A3 Polizas", t:"txt"},
 {g:"Columnas", k:"C_RAMO", l:"Ramo", v:"C", t:"col"},
 {g:"Columnas", k:"C_PRIMA", l:"Prima", v:"D", t:"col"},
 {g:"Columnas", k:"C_NUEVA", l:"Nueva", v:"E", t:"col"},
 {g:"Columnas", k:"C_F", l:"Comisión", v:"F", t:"col"},
 {g:"Columnas", k:"C_G", l:"Comisión con bono", v:"G", t:"col"},
 {g:"Columnas", k:"C_ULT", l:"Última columna a pintar", v:"G", t:"col"},
 {g:"Tasas", k:"R1", l:"Ramo 1", v:"auto", t:"low"}, {g:"Tasas", k:"T1", l:"Tasa 1", v:"0.08", t:"num"},
 {g:"Tasas", k:"R2", l:"Ramo 2", v:"vida", t:"low"}, {g:"Tasas", k:"T2", l:"Tasa 2", v:"0.12", t:"num"},
 {g:"Tasas", k:"R3", l:"Ramo 3", v:"hogar", t:"low"}, {g:"Tasas", k:"T3", l:"Tasa 3", v:"0.1", t:"num"},
 {g:"Tasas", k:"R4", l:"Ramo 4", v:"salud", t:"low"}, {g:"Tasas", k:"T4", l:"Tasa 4", v:"0.09", t:"num"},
 {g:"Reglas", k:"BONO", l:"Bono de pólizas nuevas", v:"25", t:"num"},
 {g:"Reglas", k:"SI", l:"Texto de \"sí\" en Nueva", v:"Sí", t:"txt"},
 {g:"Reglas", k:"LIMITE", l:"Prima alta desde", v:"1000", t:"num"},
 {g:"Reglas", k:"COLOR", l:"Color RGB de prima alta", v:"221, 235, 247", t:"txt"}
],
macros:[
{n:"Comision", kw:"funcion propia optional byval tasa ramo select case redondeo", d:"Función pública: se usa en VBA y como fórmula.", code:
`Public Function Comision(ByVal prima As Double, ByVal ramo As String, _
                         Optional ByVal bono As Double = 0) As Double
    '2 Devuelve prima x tasa + bono, a 2 decimales. Public: sirve como fórmula =Comision() en la hoja.
    Dim tasa As Double

    ramo = LCase(Trim(ramo)) '1 " auto" o "HOGAR" -> "auto", "hogar"

    Select Case ramo '1 elegir la tasa según el ramo
        Case "{{R1}}": tasa = {{T1}} '! los Case van en minúsculas porque arriba se usó LCase
        Case "{{R2}}": tasa = {{T2}}
        Case "{{R3}}": tasa = {{T3}}
        Case "{{R4}}": tasa = {{T4}}
        Case Else: tasa = 0 '1 ramo desconocido
    End Select

    If prima <= 0 Or tasa = 0 Then
        Comision = 0 '1 "return 0": se asigna al nombre de la función
        Exit Function
    End If

    Comision = Application.WorksheetFunction.Round(prima * tasa + bono, 2) '1 redondeo comercial '! Round de VBA solo es bancario
End Function`},

{n:"ProbarComision", kw:"probar funcion immediate", d:"Imprime los casos del enunciado en Immediate.", code:
`Sub ProbarComision()
    '2 Llama a la función con los casos del enunciado y los imprime en Immediate (Ctrl+G).
    Debug.Print Comision(850, "Auto") '1 68
    Debug.Print Comision(1200, "vida", 25) '1 169
    Debug.Print Comision(500, "Viaje") '1 0
    Debug.Print Comision(-100, "Auto") '1 0
End Sub`},

{n:"LlenarComisiones", kw:"formula columna una instruccion comillas", d:"Fórmulas en F y G, una instrucción por columna.", code:
`Sub LlenarComisiones()
    '2 Escribe la FÓRMULA de la fila 2 en todo el rango; Excel ajusta D2 -> D3 -> D4 como al arrastrar.
    Dim ws As Worksheet
    Dim ultima As Long

    Set ws = ThisWorkbook.Worksheets("{{HOJA}}")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Cells(1, {{C_F#}}).Value = "Comisión"
    ws.Cells(1, {{C_G#}}).Value = "Comisión con bono"

    ws.Range("{{C_F}}2:{{C_F}}" & ultima).Formula = "=Comision({{C_PRIMA}}2,{{C_RAMO}}2)" '! .Formula va en inglés y con comas aunque tu Excel use ;
    ws.Range("{{C_G}}2:{{C_G}}" & ultima).Formula = _
        "=Comision({{C_PRIMA}}2,{{C_RAMO}}2,IF({{C_NUEVA}}2=""{{SI}}"",{{BONO}},0))" '! comillas dentro del texto: se escriben dobles ""
End Sub`},

{n:"EsPrimaAlta", kw:"funcion privada boolean", d:"Función privada: True o False.", code:
`Private Function EsPrimaAlta(ByVal prima As Double) As Boolean
    '2 Private: solo la usan procedimientos de este módulo; no aparece en Alt+F8.
    EsPrimaAlta = (prima >= {{LIMITE}}) '1 la comparación ya es True o False
End Function`},

{n:"PintarFila", kw:"sub privado parametros pintar fila color", d:"Sub privado con parámetros.", code:
`Private Sub PintarFila(ByVal ws As Worksheet, ByVal fila As Long, ByVal colorFila As Long)
    '2 Recibe la hoja, la fila y el color, y rellena desde la columna A hasta la última indicada.
    ws.Range("A" & fila & ":{{C_ULT}}" & fila).Interior.Color = colorFila '1 "A5:G5" si fila = 5
End Sub`},

{n:"ResaltarPrimasAltas", kw:"resaltar pintar quitar relleno contar privado call", d:"Pinta las primas altas y cuenta cuántas son.", code:
`Sub ResaltarPrimasAltas()
    '2 Usa las dos piezas privadas: EsPrimaAlta decide y PintarFila pinta.
    Dim ws As Worksheet
    Dim ultima As Long, fila As Long, altas As Long

    Set ws = ThisWorkbook.Worksheets("{{HOJA}}")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For fila = 2 To ultima
        If EsPrimaAlta(ws.Cells(fila, {{C_PRIMA#}}).Value) Then
            PintarFila ws, fila, RGB({{COLOR}}) '! Sub con parámetros: sin paréntesis (o con Call adelante)
            altas = altas + 1
        Else
            ws.Range("A" & fila & ":{{C_ULT}}" & fila).Interior.ColorIndex = xlNone '1 quitar relleno
        End If
    Next fila

    Debug.Print "Pólizas de prima alta: " & altas
End Sub`}
],
comprobar:[
 "ProbarComision imprime 68, 169, 0 y 0.",
 "La columna F suma $1,912.33 y la G, $2,112.33 (8 pólizas nuevas).",
 "POL-26-008 (\"  auto\", prima 975.40, nueva): 78.03 en F y 103.03 en G.",
 "ResaltarPrimasAltas imprime 9.",
 "En Alt+F8 aparece ResaltarPrimasAltas, pero no EsPrimaAlta ni PintarFila."
],
variantes:[
 {t:"Bono como porcentaje", d:"\"Las pólizas nuevas reciben un 2% adicional de la prima.\"",
  a:`        "=Comision({{C_PRIMA}}2,{{C_RAMO}}2,IF({{C_NUEVA}}2=""{{SI}}"",{{BONO}},0))"`,
  b:`        "=Comision({{C_PRIMA}}2,{{C_RAMO}}2,IF({{C_NUEVA}}2=""{{SI}}"",{{C_PRIMA}}2*0.02,0))"`},
 {t:"Agregar un ramo", d:"\"Transporte paga 7%.\" Se agrega una línea antes de Case Else, en minúsculas.",
  a:`        Case Else: tasa = 0`,
  b:`        Case "transporte": tasa = 0.07
        Case Else: tasa = 0`},
 {t:"Valores en vez de fórmulas", d:"\"Calcular la comisión desde VBA y escribir el valor.\" Se llama a la función en un bucle.",
  a:`    ws.Range("{{C_F}}2:{{C_F}}" & ultima).Formula = "=Comision({{C_PRIMA}}2,{{C_RAMO}}2)"`,
  b:`    Dim fila As Long
    For fila = 2 To ultima
        ws.Cells(fila, {{C_F#}}).Value = Comision(ws.Cells(fila, {{C_PRIMA#}}).Value, ws.Cells(fila, {{C_RAMO#}}).Value)
    Next fila`},
 {t:"Fórmula con tu separador", d:"Si preferís escribir la fórmula con \";\", como en tu Excel, usá FormulaLocal.",
  a:`    ws.Range("{{C_F}}2:{{C_F}}" & ultima).Formula = "=Comision({{C_PRIMA}}2,{{C_RAMO}}2)"`,
  b:`    ws.Range("{{C_F}}2:{{C_F}}" & ultima).FormulaLocal = "=Comision({{C_PRIMA}}2;{{C_RAMO}}2)"`},
 {t:"Total al final de la columna", d:"\"Agregar una fila Total con la suma de la comisión.\"",
  a:`End Sub`,
  b:`    ws.Cells(ultima + 1, 1).Value = "Total"
    ws.Cells(ultima + 1, {{C_F#}}).Formula = "=SUM({{C_F}}2:{{C_F}}" & ultima & ")"
    ws.Rows(ultima + 1).Font.Bold = True
End Sub`}
]
},

/* =====================================================================
   A4  PEDIDOS
   ===================================================================== */
{
id:"A4", titulo:"Pedidos de la tienda en línea", modulo:"A4_Pedidos",
resumen:"Leer la hoja una sola vez en un arreglo, clientes sin repetir con Collection y resumen por categoría con Dictionary.",
nota:"El arreglo se lee desde la columna A, por eso la columna 3 del arreglo es la C de la hoja. Si el rango empezara en otra columna, los números cambian.",
pide:[
 ["TotalesPedidos","Total de cada pedido en un arreglo y escrito de una sola vez en G. Imprime cuántos pedidos y cuánto suman."],
 ["ClientesDistintos","Collection con los clientes sin repetir, en orden de aparición. Imprime cuántos y sus nombres."],
 ["ResumenPorCategoria","Dictionary con ventas y pedidos por categoría. Hoja nueva al final (si existe, se borra sin preguntar), fila Total, formatos y MsgBox con la que más vendió."]
],
params:[
 {g:"Hojas", k:"HOJA", l:"Hoja de pedidos", v:"A4 Pedidos", t:"txt"},
 {g:"Hojas", k:"HOJA_RES", l:"Hoja de resumen", v:"A4 Resumen", t:"txt"},
 {g:"Columnas", k:"C_ULT", l:"Última columna de datos", v:"F", t:"col"},
 {g:"Columnas", k:"C_CLI", l:"Cliente", v:"C", t:"col"},
 {g:"Columnas", k:"C_CAT", l:"Agrupar por (categoría)", v:"D", t:"col"},
 {g:"Columnas", k:"C_UNI", l:"Unidades", v:"E", t:"col"},
 {g:"Columnas", k:"C_PRE", l:"Precio", v:"F", t:"col"},
 {g:"Columnas", k:"C_TOT", l:"Total (resultado)", v:"G", t:"col"},
 {g:"Textos", k:"ETIQ", l:"Encabezado del grupo", v:"Categoría", t:"txt"}
],
macros:[
{n:"TotalesPedidos", kw:"arreglo leer hoja una vez memoria escribir de una sola vez redim ubound", d:"Arreglo de entrada y de salida.", code:
`Sub TotalesPedidos()
    '2 Lee la tabla a memoria, calcula en un arreglo y pega todos los totales de una sola vez.
    Dim ws As Worksheet
    Dim ultima As Long, n As Long, i As Long
    Dim datos As Variant '1 aquí cae la tabla entera '! un arreglo leído de la hoja va As Variant
    Dim totales() As Double '1 tamaño se define con ReDim
    Dim suma As Double

    Set ws = ThisWorkbook.Worksheets("{{HOJA}}")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    datos = ws.Range("A2:{{C_ULT}}" & ultima).Value '1 la hoja se lee UNA sola vez
    n = UBound(datos, 1) '1 cuántas filas tiene el arreglo

    ReDim totales(1 To n, 1 To 1) '1 n filas, 1 columna '! debe ser 2D para pegarlo en una columna

    For i = 1 To n
        totales(i, 1) = Application.WorksheetFunction.Round(datos(i, {{C_UNI#}}) * datos(i, {{C_PRE#}}), 2) '1 unidades x precio
        suma = suma + totales(i, 1)
    Next i

    ws.Cells(1, {{C_TOT#}}).Value = "Total"
    ws.Range("{{C_TOT}}2:{{C_TOT}}" & ultima).Value = totales '1 todo de un golpe '! el rango debe tener el mismo tamaño que el arreglo

    Debug.Print "Pedidos: " & n
    Debug.Print "Suman: " & Format(suma, "$#,##0.00")
End Sub`},

{n:"ClientesDistintos", kw:"sin repetir unicos distintos collection lista", d:"Collection sin repetidos.", code:
`Sub ClientesDistintos()
    '2 Recorre los pedidos y agrega cada cliente a la lista solo si todavía no está.
    Dim ws As Worksheet
    Dim ultima As Long, i As Long, j As Long
    Dim datos As Variant
    Dim clientes As Collection
    Dim nombre As String
    Dim existe As Boolean
    Dim c As Variant '! la variable del For Each sobre una Collection va As Variant

    Set ws = ThisWorkbook.Worksheets("{{HOJA}}")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    datos = ws.Range("A2:{{C_ULT}}" & ultima).Value

    Set clientes = New Collection '1 lista vacía (objeto: lleva Set)

    For i = 1 To UBound(datos, 1)
        nombre = datos(i, {{C_CLI#}})

        existe = False '1 ¿ya está en la lista? recorrerla y comparar
        For j = 1 To clientes.Count
            If clientes(j) = nombre Then
                existe = True
                Exit For
            End If
        Next j

        If existe = False Then clientes.Add nombre '1 agregar al final
    Next i

    Debug.Print "Clientes distintos: " & clientes.Count
    For Each c In clientes
        Debug.Print c
    Next c
End Sub`},

{n:"ResumenPorCategoria", kw:"total por categoria agrupar dictionary diccionario hoja nueva eliminar reemplazar resumen porcentaje mayor", d:"Dictionary, hoja nueva y formatos.", code:
`Sub ResumenPorCategoria()
    '2 Una pasada por los datos acumula ventas y pedidos por grupo; luego arma la hoja de resumen.
    Dim ws As Worksheet, wsR As Worksheet, hoja As Worksheet
    Dim ultima As Long, i As Long, fila As Long
    Dim datos As Variant
    Dim ventas As Object, pedidos As Object '1 dos diccionarios: grupo -> $ y grupo -> cantidad
    Dim cat As Variant '! la clave del For Each va As Variant
    Dim monto As Double, total As Double
    Dim mejorCat As String, mejorMonto As Double

    Set ws = ThisWorkbook.Worksheets("{{HOJA}}")
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    datos = ws.Range("A2:{{C_ULT}}" & ultima).Value

    '1 Acumular en una sola pasada
    Set ventas = CreateObject("Scripting.Dictionary")
    Set pedidos = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(datos, 1)
        cat = datos(i, {{C_CAT#}})
        monto = Application.WorksheetFunction.Round(datos(i, {{C_UNI#}}) * datos(i, {{C_PRE#}}), 2)
        If Not ventas.Exists(cat) Then '1 grupo nuevo: crearlo en 0
            ventas.Add cat, 0
            pedidos.Add cat, 0
        End If
        ventas(cat) = ventas(cat) + monto '1 como d[clave] += x
        pedidos(cat) = pedidos(cat) + 1
        total = total + monto
    Next i

    '1 Borrar la hoja de la ejecución anterior, sin preguntar
    Application.DisplayAlerts = False
    For Each hoja In ThisWorkbook.Worksheets
        If hoja.Name = "{{HOJA_RES}}" Then
            hoja.Delete
            Exit For
        End If
    Next hoja
    Application.DisplayAlerts = True '! volver a encender los avisos

    '1 Crear la hoja al final del libro
    Set wsR = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsR.Name = "{{HOJA_RES}}"

    '1 Tabla
    wsR.Range("A1:D1").Value = Array("{{ETIQ}}", "Pedidos", "Ventas ($)", "% del total")
    fila = 2
    For Each cat In ventas.Keys '1 en el orden en que aparecieron
        wsR.Cells(fila, 1).Value = cat
        wsR.Cells(fila, 2).Value = pedidos(cat)
        wsR.Cells(fila, 3).Value = ventas(cat)
        wsR.Cells(fila, 4).Value = ventas(cat) / total
        If ventas(cat) > mejorMonto Then '1 el mayor, a mano
            mejorMonto = ventas(cat)
            mejorCat = cat
        End If
        fila = fila + 1
    Next cat

    '1 Fila Total
    wsR.Cells(fila, 1).Value = "Total"
    wsR.Cells(fila, 2).Value = UBound(datos, 1)
    wsR.Cells(fila, 3).Value = total
    wsR.Cells(fila, 4).Value = 1

    '1 Formatos
    wsR.Range("A1:D1").Font.Bold = True
    wsR.Range("A" & fila & ":D" & fila).Font.Bold = True
    wsR.Range("C2:C" & fila).NumberFormat = "#,##0.00"
    wsR.Range("D2:D" & fila).NumberFormat = "0.0%"
    wsR.Columns("A:D").AutoFit

    MsgBox "{{ETIQ}} que más vendió: " & mejorCat & vbCrLf & _
           "Ventas: " & Format(mejorMonto, "$#,##0.00"), vbInformation, "Resumen"
End Sub`}
],
comprobar:[
 "TotalesPedidos: 120 pedidos por $16,314.00. TL-5002 queda con 324.00 en G.",
 "ClientesDistintos: 14 clientes, empezando por Andrea López y Mario Castillo.",
 "ResumenPorCategoria corrido dos veces: sigue habiendo una sola hoja A4 Resumen.",
 "Electrónica 26 / 6,146.00 / 37.7% · Hogar 35 / 4,001.75 / 24.5% · Juguetes 22 / 1,656.50 / 10.2% · Libros 21 / 1,320.00 / 8.1% · Deportes 16 / 3,189.75 / 19.6% · Total 120 / 16,314.00.",
 "El MsgBox muestra Electrónica con $6,146.00."
],
variantes:[
 {t:"Agrupar por cliente", d:"\"Resumen por cliente.\" En el adaptador: \"Agrupar por\" = C y encabezado = Cliente. Sin tocar el código.",
  a:`        cat = datos(i, {{C_CAT#}})`, b:`        cat = datos(i, {{C_CLI#}}) '1 o cambiar la columna en el adaptador`},
 {t:"Columna de ticket promedio", d:"\"Agregar el promedio por pedido.\" Una columna más en la tabla.",
  a:`        wsR.Cells(fila, 4).Value = ventas(cat) / total`,
  b:`        wsR.Cells(fila, 4).Value = ventas(cat) / total
        wsR.Cells(fila, 5).Value = ventas(cat) / pedidos(cat) '1 promedio = suma / cantidad`},
 {t:"Ordenar de mayor a menor", d:"\"Ordenar la tabla por ventas.\" Después de escribirla y antes de la fila Total.",
  a:`    '1 Fila Total`,
  b:`    wsR.Range("A2:D" & fila - 1).Sort Key1:=wsR.Range("C2"), Order1:=xlDescending, Header:=xlNo
    '1 Fila Total`},
 {t:"Limpiar la clave", d:"Si las categorías vienen con espacios o mayúsculas distintas, \"Hogar\" y \" hogar\" serían dos grupos.",
  a:`        cat = datos(i, {{C_CAT#}})`, b:`        cat = Trim(datos(i, {{C_CAT#}}))`},
 {t:"Total con WorksheetFunction", d:"\"Calcular el total de la columna G.\" Sin bucle.",
  a:`    Debug.Print "Suman: " & Format(suma, "$#,##0.00")`,
  b:`    Debug.Print "Suman: " & Format(Application.WorksheetFunction.Sum(ws.Range("{{C_TOT}}2:{{C_TOT}}" & ultima)), "$#,##0.00")`}
]
}
];
