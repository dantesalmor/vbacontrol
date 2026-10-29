/* Fase 2: De enunciado a código.
   Cada entrada:
     id     identificador
     cat    categoría (orden en que se arma una macro)
     t      título de la pieza
     fr     frases del enunciado que la piden (se buscan sin tildes ni mayúsculas)
     q      qué significa en código, en una o dos oraciones
     code   la pieza (mismos marcadores que las plantillas: '1 corto, '2 largo, '! trampa)
     tr     trampa típica (opcional)
     ver    dónde aparece completa: "A1:CalcularParqueo" (opcional) */

window.CATEGORIAS = [
  ["base", "Estructura"],
  ["pedir", "Pedir datos al usuario"],
  ["calc", "Calcular y decidir"],
  ["texto", "Texto"],
  ["fecha", "Fechas"],
  ["agrupar", "Listas y agrupar"],
  ["hojas", "Hojas, libros y fórmulas"],
  ["formato", "Formato y colores"],
  ["proc", "Funciones y procedimientos"],
  ["mostrar", "Mostrar resultados"]
];

window.FRASES = [

/* ---------- Estructura ---------- */
{id:"ultima", cat:"base", t:"Encontrar la última fila",
 fr:["encuentra sola la ultima fila","ultima fila","aunque haya mas","aunque manana haya","funcionar aunque","cualquier cantidad de filas","sin importar cuantas filas"],
 q:"No escribas 31 ni 121 fijo: la macro sube desde el fondo de la hoja hasta el último dato.",
 code:`ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row '1 columna llena en todas las filas
For fila = 2 To ultima '1 fila 1 = encabezados
    ' ...
Next fila`,
 tr:"Usá la columna del código o ticket: si la columna tiene vacíos, la última fila sale corta.", ver:"A1:CalcularParqueo"},

{id:"encab", cat:"base", t:"Escribir encabezados",
 fr:["con sus encabezados","con su encabezado","encabezado","encabezados"],
 q:"Escribí los títulos en la fila 1 antes del bucle. Con Array escribís varias celdas en una línea.",
 code:`ws.Range("F1:I1").Value = Array("Minutos", "Cobro", "Descuento", "Valor a pagar") '1 tantos textos como celdas
ws.Range("F1:I1").Font.Bold = True '1 opcional: negrita`,
 tr:"La cantidad de textos del Array debe ser igual a la cantidad de celdas del rango.", ver:"A1:CalcularParqueo"},

{id:"columnas", cat:"base", t:"Escribir el resultado en una columna",
 fr:["en las columnas","en la columna","escribir en","columnas f a i","columna g"],
 q:"Dentro del bucle, cada resultado va a Cells(fila, número de columna). A=1, B=2, … G=7.",
 code:`ws.Cells(fila, 7).Value = resultado '1 columna G de esta fila`,
 ver:"A1:CalcularParqueo"},

{id:"arrIn", cat:"base", t:"Leer la hoja una sola vez (arreglo)",
 fr:["lee la hoja una sola vez","una sola vez con","trabaja con el arreglo","en un arreglo","arreglo","en memoria"],
 q:"Toda la tabla pasa a memoria en una línea. datos(i, c): fila i del arreglo (fila i+1 de la hoja), columna c.",
 code:`Dim datos As Variant '! un arreglo leído de la hoja va As Variant
datos = ws.Range("A2:F" & ultima).Value
For i = 1 To UBound(datos, 1) '1 UBound(datos, 1) = cuántas filas
    total = total + datos(i, 5) * datos(i, 6) '1 columnas E y F
Next i`,
 tr:"La fila 1 del arreglo es la fila 2 de la hoja (el rango empezó en A2).", ver:"A4:TotalesPedidos"},

{id:"arrOut", cat:"base", t:"Escribir todos los resultados de una sola vez",
 fr:["de una sola vez","todos los totales","escribir todos","de un solo golpe"],
 q:"Llenás un arreglo de n filas × 1 columna y lo pegás entero en la columna.",
 code:`Dim res() As Double
ReDim res(1 To n, 1 To 1) '1 n filas, 1 columna '! debe ser 2D para pegarlo en una columna
For i = 1 To n
    res(i, 1) = datos(i, 5) * datos(i, 6)
Next i
ws.Range("G2:G" & ultima).Value = res '! mismo tamaño que el arreglo`,
 ver:"A4:TotalesPedidos"},

{id:"param", cat:"base", t:"Leer un valor de la hoja de parámetros",
 fr:["se lee de la hoja","hoja de parametros","sin tocar el codigo","desde la hoja","si cambian ahi"],
 q:"El valor no se escribe en el código: se lee de su celda al inicio de la macro, así si cambia la hoja, cambia el resultado.",
 code:`corte = ThisWorkbook.Worksheets("A2 Parámetros").Range("B1").Value '1 nombre exacto de la hoja`,
 tr:"Error 9: el nombre de la hoja no coincide exactamente (tilde, espacio).", ver:"A2:VencimientosProveedores"},

{id:"noDup", cat:"base", t:"No duplicar si se corre dos veces",
 fr:["no debe duplicar","si se ejecuta dos veces","ejecuta otra vez","al ejecutar otra vez","volver a ejecutar","se reemplaza","resultado anterior"],
 q:"Antes de escribir, borrá lo que dejó la ejecución anterior (o la hoja, si la macro crea una).",
 code:`ws.Range("G2:I" & ws.Rows.Count).ClearContents '1 borra valores, deja el formato
ws.Range("G2:I" & ws.Rows.Count).Interior.ColorIndex = xlNone '1 y los colores viejos`,
 tr:"Si la macro agrega una fila Total, calculá la última fila ANTES de esa fila, o borrala primero."},

{id:"optexp", cat:"base", t:"Option Explicit",
 fr:["option explicit","declarar las variables"],
 q:"Primera línea de cada módulo: obliga a declarar todo con Dim.",
 code:`Option Explicit`},

/* ---------- Pedir datos ---------- */
{id:"input", cat:"pedir", t:"Pedir un dato con InputBox",
 fr:["inputbox","pedir el","preguntar el","preguntar entre","preguntar cuantas","preguntar cuantos","que escriba","pida"],
 q:"InputBox devuelve siempre texto. Guardalo en un String.",
 code:`Dim resp As String
resp = InputBox("Número de ticket:", "Buscar ticket") '1 mensaje, título`,
 ver:"A1:BuscarTicket"},

{id:"cancel", cat:"pedir", t:"Si cancela, no pasa nada",
 fr:["si el usuario cancela","si cancela","cancelar","no pasa nada","terminar sin mensaje"],
 q:"StrPtr da 0 solo cuando se presionó Cancelar. Exit Sub termina la macro ahí.",
 code:`If StrPtr(resp) = 0 Then Exit Sub '1 presionó Cancelar
If Trim(resp) = "" Then Exit Sub '1 aceptó vacío`,
 ver:"A1:BuscarTicket"},

{id:"normal", cat:"pedir", t:"Aceptar minúsculas y espacios",
 fr:["aunque se escriba en minusculas","con espacios","minusculas o con espacios","venir con espacios"],
 q:"Normalizá lo que escribe el usuario y también el dato de la hoja antes de comparar.",
 code:`resp = UCase(Trim(resp)) '1 " pq-0117 " -> "PQ-0117"
If UCase(Trim(ws.Cells(fila, 1).Value)) = resp Then '! normalizar los dos lados`,
 ver:"A1:BuscarTicket"},

{id:"buscar", cat:"pedir", t:"Buscar y avisar si no existe",
 fr:["si no existe","no existe","se avisa","buscar","busque","encontrar"],
 q:"Recorré la columna; al encontrar, guardá la fila y salí con Exit For. Si la fila sigue en 0, no existe.",
 code:`filaEnc = 0
For fila = 2 To ultima
    If UCase(Trim(ws.Cells(fila, 1).Value)) = resp Then
        filaEnc = fila
        Exit For '1 = break
    End If
Next fila
If filaEnc = 0 Then
    MsgBox "El código " & resp & " no existe.", vbExclamation
    Exit Sub
End If`,
 ver:"A1:BuscarTicket"},

{id:"entero", cat:"pedir", t:"Validar un número entero (y su rango)",
 fr:["entero mayor que 0","numero entero","entero entre","entre 1 y","no sea un numero"],
 q:"Primero que sea número, después que sea entero, después el rango. Si falla, avisar y salir.",
 code:`If Not IsNumeric(resp) Then GoTo Invalido
If Val(resp) <> Int(Val(resp)) Or Val(resp) < 1 Then GoTo Invalido '1 cambiar el rango aquí
n = CLng(resp)
' ... el resto de la macro usa n ...
Exit Sub '! sin este Exit Sub, siempre cae en el aviso
Invalido:
MsgBox "Debe escribir un entero mayor que 0.", vbExclamation`},

{id:"sino", cat:"pedir", t:"Preguntar Sí o No",
 fr:["pregunte si","si la respuesta es si","confirmar","desea registrar","se registra"],
 q:"MsgBox con vbYesNo devuelve el botón presionado.",
 code:`If MsgBox("¿Registrar la cotización?", vbYesNo + vbQuestion) = vbNo Then Exit Sub`},

/* ---------- Calcular y decidir ---------- */
{id:"tramos", cat:"calc", t:"Tramos (desde… hasta…)",
 fr:["tramos","por tramos","hasta antes de","menos del","mas de","a partir de","hasta","desde"],
 q:"Cadena If / ElseIf de menor a mayor: VBA entra al PRIMERO que se cumpla.",
 code:`If x <= 15 Then
    tasa = 0
ElseIf x <= 60 Then '1 ya se sabe que x > 15
    tasa = 1
Else
    tasa = 2
End If`,
 tr:"\"Hasta antes de 100%\" es < 1; \"desde 100%\" es >= 1. Leé con cuidado si el límite entra o no.", ver:"A1:CalcularParqueo"},

{id:"techo", cat:"calc", t:"Redondear hacia arriba (o fracción)",
 fr:["o fraccion","fraccion","hacia arriba","cajas enteras","redondea hacia arriba"],
 q:"Si pasa aunque sea un poco, cuenta el bloque completo. -Int(-x) es el techo de x.",
 code:`bloques = -Int(-(minutos - 60) / 30) '1 61 min -> 1 bloque '! Round aquí da mal`,
 ver:"A1:CalcularParqueo"},

{id:"round2", cat:"calc", t:"Dos decimales, redondeo comercial",
 fr:["dos decimales","2 decimales","redondeo comercial","redondeado a","a centavos","se redondean"],
 q:"El Round de VBA es bancario (2.125 → 2.12). El de Excel sube el .5.",
 code:`x = Application.WorksheetFunction.Round(prima * tasa, 2)`,
 ver:"A3:Comision"},

{id:"tope", cat:"calc", t:"Tope o máximo",
 fr:["cobro maximo","maximo del dia","no puede superar","como maximo","tope","no supera"],
 q:"Si se pasa del límite, se deja en el límite.",
 code:`If cobro > 6 Then cobro = 6 '1 o: cobro = Application.WorksheetFunction.Min(cobro, 6)`,
 ver:"A1:CalcularParqueo"},

{id:"noneg", cat:"calc", t:"Que no quede negativo",
 fr:["sin que quede negativo","no quede negativo","nunca negativo","0 si todavia no","0 si aun no"],
 q:"Max(x, 0): si x es negativo, queda 0.",
 code:`pagar = Application.WorksheetFunction.Max(cobro - desc, 0)`,
 ver:"A1:CalcularParqueo"},

{id:"select", cat:"calc", t:"Un valor según la categoría (tasa por tipo)",
 fr:["segun el ramo","segun el tipo","segun la categoria","las tasas son","cualquier otro","no se reconoce"],
 q:"Select Case compara una variable contra varios valores. Case Else = ninguno.",
 code:`Select Case LCase(Trim(ramo)) '1 limpiar antes de comparar
    Case "auto": tasa = 0.08
    Case "vida": tasa = 0.12
    Case Else: tasa = 0 '1 cualquier otro
End Select`,
 tr:"Si limpiás con LCase, los textos de Case van en minúsculas.", ver:"A3:Comision"},

{id:"contar", cat:"calc", t:"Contar cuántos cumplen algo",
 fr:["cuantos","cuantas","numero de","cantidad de"],
 q:"Un contador Long que suma 1 dentro del If.",
 code:`If condicion Then cuenta = cuenta + 1 '1 arranca en 0 solo`},

{id:"sumar", cat:"calc", t:"Total o suma",
 fr:["el total","total a cobrar","cuanto suman","suma de","monto total","cuanto suma"],
 q:"Un acumulador Double que suma dentro del bucle.",
 code:`total = total + pagar '1 dentro del For`},

{id:"mayor", cat:"calc", t:"El mayor (el que más…)",
 fr:["el que mas","la que mas","que mas vendio","el mayor","la mayor","mas alto","el mejor"],
 q:"Guardá el mejor valor visto hasta ahora y su nombre; si aparece uno mayor, lo reemplaza.",
 code:`If ventas(cat) > mejorMonto Then
    mejorMonto = ventas(cat)
    mejorCat = cat
End If`,
 ver:"A4:ResumenPorCategoria"},

{id:"pct", cat:"calc", t:"Porcentaje del total",
 fr:["% del total","porcentaje del","porcentaje","cumplimiento"],
 q:"Parte / total; el formato % se pone en la celda.",
 code:`ws.Cells(fila, 4).Value = parte / total
ws.Range("D2:D" & fila).NumberFormat = "0.0%"`,
 ver:"A4:ResumenPorCategoria"},

/* ---------- Texto ---------- */
{id:"trim", cat:"texto", t:"Quitar espacios (también los dobles del medio)",
 fr:["sin espacios al inicio","espacios dobles","dobles o triples","espacios de mas","sin espacios"],
 q:"El Trim de Excel quita los extremos y deja un solo espacio entre palabras.",
 code:`x = Application.WorksheetFunction.Trim(x) '! el Trim de VBA solo quita los extremos`,
 ver:"A2:LimpiarContactos"},

{id:"proper", cat:"texto", t:"Primera letra de cada palabra en mayúscula",
 fr:["primera letra de cada palabra","nombre propio","mayuscula la primera"],
 q:"Proper de Excel: \"CARLOS rivas\" → \"Carlos Rivas\".",
 code:`x = Application.WorksheetFunction.Proper(x)`,
 ver:"A2:LimpiarContactos"},

{id:"caso", cat:"texto", t:"Mayúsculas o minúsculas",
 fr:["en minusculas","minusculas","en mayusculas","mayusculas"],
 q:"LCase pasa a minúsculas; UCase, a mayúsculas.",
 code:`correo = LCase(Trim(correo)) '1 minúsculas y sin espacios alrededor`,
 ver:"A2:LimpiarContactos"},

{id:"despues", cat:"texto", t:"Lo que va después (o antes) de un carácter",
 fr:["lo que va despues","despues de la @","el dominio","dominio","antes de la @","usuario del correo"],
 q:"InStr da la posición del carácter; Mid toma desde ahí; Left toma lo anterior.",
 code:`dominio = Mid(correo, InStr(correo, "@") + 1) '1 después de la @
usuario = Left(correo, InStr(correo, "@") - 1) '1 antes de la @`,
 ver:"A2:LimpiarContactos"},

{id:"contarCar", cat:"texto", t:"Cuántas veces aparece un carácter",
 fr:["mas de una","una sola @","no tiene @","exactamente una","cuantas veces"],
 q:"Largo del texto menos largo del texto sin ese carácter.",
 code:`arrobas = Len(correo) - Len(Replace(correo, "@", ""))`,
 ver:"A2:LimpiarContactos"},

{id:"like", cat:"texto", t:"Validar un formato (patrón)",
 fr:["deben tener la forma","tiene la forma","forma p","formato","patron","validar","digitos"],
 q:"Like compara con un patrón: # dígito, ? cualquier carácter, * cualquier cosa, [A-Z] una letra.",
 code:`If cel.Value Like "P###-###" Then
    cel.Interior.ColorIndex = xlNone '1 correcta
Else
    cel.Interior.Color = vbYellow '1 incorrecta
End If`,
 tr:"Like distingue mayúsculas. Para aceptar minúsculas: UCase(x) Like …", ver:"A1:ValidarPlacas"},

{id:"tel", cat:"texto", t:"Quitar caracteres (guiones, puntos, paréntesis)",
 fr:["guiones","parentesis","puntos o","telefono","codigo de pais"],
 q:"Un Replace por cada carácter, uno dentro de otro.",
 code:`tel = Replace(Replace(Replace(tel, " ", ""), "-", ""), ".", "")
If Len(tel) = 11 And Left(tel, 3) = "503" Then tel = Mid(tel, 4) '1 quitar 503`},

/* ---------- Fechas ---------- */
{id:"sumaDias", cat:"fecha", t:"Fecha más días",
 fr:["mas los dias","fecha de vencimiento","dias de credito","vencimiento","dias despues"],
 q:"Una fecha es un número: fecha + días = otra fecha.",
 code:`venc = ws.Cells(fila, 4).Value + ws.Cells(fila, 5).Value
ws.Cells(fila, 10).NumberFormat = "dd/mm/yyyy" '1 si no, se ve como número`,
 ver:"A2:VencimientosProveedores"},

{id:"finde", cat:"fecha", t:"Si cae sábado o domingo",
 fr:["sabado o domingo","fin de semana","lunes siguiente","pasa al lunes","viernes anterior","dia habil"],
 q:"Weekday: 1 = domingo … 7 = sábado. Sumá o restá según a qué día pasa.",
 code:`If Weekday(venc) = 7 Then
    venc = venc + 2 '1 sábado -> lunes
ElseIf Weekday(venc) = 1 Then
    venc = venc + 1 '1 domingo -> lunes
End If`,
 tr:"Para el viernes anterior: sábado - 1 y domingo - 2.", ver:"A2:VencimientosProveedores"},

{id:"atraso", cat:"fecha", t:"Días de atraso a una fecha de corte",
 fr:["dias de atraso","dias vencidos","a la fecha de corte","fecha de corte","atraso"],
 q:"Fecha - fecha = días. Si da negativo, todavía no vence: 0.",
 code:`atraso = corte - venc
If atraso < 0 Then atraso = 0`,
 ver:"A2:VencimientosProveedores"},

{id:"horas", cat:"fecha", t:"Minutos entre dos horas",
 fr:["minutos","hora de entrada","hora de salida","horas trabajadas"],
 q:"Una hora es una fracción del día: (salida - entrada) × 1440 = minutos (× 24 = horas).",
 code:`minutos = Round((salida - entrada) * 1440, 0)`,
 ver:"A1:CalcularParqueo"},

{id:"mes", cat:"fecha", t:"Mes, año o fecha como texto",
 fr:["aaaa-mm","aaaammdd","formato aaaa","del mes","por mes","fecha del dia","fecha de hoy"],
 q:"Date es hoy. Month, Year y Day sacan partes. Format convierte en texto para nombres.",
 code:`hoy = Date
nombre = "OC-" & Format(Date, "yyyymmdd") '1 OC-20261030
If Month(f) = 3 Then '1 solo marzo`},

/* ---------- Listas y agrupar ---------- */
{id:"distintos", cat:"agrupar", t:"Sin repetir (valores distintos)",
 fr:["sin repetir","distintos","unicos","no repetidos","en el orden en que aparecen"],
 q:"Una Collection: antes de agregar, recorrerla para ver si ya está.",
 code:`Set lista = New Collection
For i = 1 To UBound(datos, 1)
    existe = False
    For j = 1 To lista.Count
        If lista(j) = datos(i, 3) Then existe = True: Exit For
    Next j
    If Not existe Then lista.Add datos(i, 3)
Next i`,
 ver:"A4:ClientesDistintos"},

{id:"dicc", cat:"agrupar", t:"Total por categoría, cliente o región",
 fr:["por categoria","por cliente","por region","por sucursal","por vendedor","por agente","por empleado","acumular","dictionary","resumen por"],
 q:"Un Dictionary: la clave es el grupo, el valor es el acumulado. Si la clave no existe, se crea en 0.",
 code:`Set d = CreateObject("Scripting.Dictionary")
For i = 1 To UBound(datos, 1)
    k = datos(i, 4) '1 columna del grupo
    If Not d.Exists(k) Then d.Add k, 0
    d(k) = d(k) + datos(i, 5) * datos(i, 6)
Next i
For Each k In d.Keys '! k As Variant
    Debug.Print k, d(k)
Next k`,
 ver:"A4:ResumenPorCategoria"},

{id:"dinamico", cat:"agrupar", t:"Grupos que no se escriben en el código",
 fr:["no se escriben en el codigo","no deben escribirse","region nueva","si se agrega un","aparece una"],
 q:"Los grupos salen de los datos (claves del Dictionary), no de una lista fija en el código. Si mañana aparece uno nuevo, entra solo.",
 code:`For Each k In d.Keys '1 cada grupo que apareció en los datos
    ' crear su fila o su hoja
Next k`,
 ver:"A4:ResumenPorCategoria"},

/* ---------- Hojas, libros y fórmulas ---------- */
{id:"hojaNueva", cat:"hojas", t:"Crear una hoja (si existe, eliminarla sin preguntar)",
 fr:["crear la hoja","crear una hoja","hoja nueva","si ya existe","se elimina sin preguntar","se reemplaza sin preguntar"],
 q:"Apagás los avisos, borrás la hoja vieja si la encontrás, y creás la nueva al final.",
 code:`Application.DisplayAlerts = False
For Each h In ThisWorkbook.Worksheets
    If h.Name = "A4 Resumen" Then h.Delete: Exit For
Next h
Application.DisplayAlerts = True '! volver a encender
Set wsR = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
wsR.Name = "A4 Resumen"`,
 tr:"El nombre de hoja tiene máximo 31 caracteres y no admite : \\ / ? * [ ].", ver:"A4:ResumenPorCategoria"},

{id:"junto", cat:"hojas", t:"Crear la hoja junto a otra",
 fr:["junto a la hoja","despues de la hoja","al lado de"],
 q:"After indica después de qué hoja se crea.",
 code:`Set wsN = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets("E09 Inventario"))`},

{id:"porGrupo", cat:"hojas", t:"Una hoja por cada grupo",
 fr:["una hoja por","hoja por region","hoja por cada"],
 q:"Por cada clave del Dictionary, una hoja nueva con su nombre.",
 code:`For Each k In d.Keys
    Set wsN = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsN.Name = "R-" & k
Next k`},

{id:"formula", cat:"hojas", t:"Una fórmula en toda la columna (una instrucción)",
 fr:["una sola instruccion","como formula","la formula","poner en","formula"],
 q:"Se escribe la fórmula de la fila 2 en todo el rango; Excel ajusta las referencias.",
 code:`ws.Range("F2:F" & ultima).Formula = "=Comision(D2,C2)"
ws.Range("G2:G" & ultima).Formula = "=IF(E2=""Sí"",25,0)" '! comillas dobles dentro`,
 tr:".Formula va en inglés y con comas, aunque tu Excel use \";\".", ver:"A3:LlenarComisiones"},

{id:"abrir", cat:"hojas", t:"Abrir otro archivo (si no existe, seguir)",
 fr:["abrir su archivo","si un archivo no existe","no encontrado","no detener el proceso","cerrarlo sin guardar"],
 q:"Dir devuelve \"\" si el archivo no existe. Abrís, copiás y cerrás sin guardar.",
 code:`ruta = ThisWorkbook.Path & "\\" & archivo
If Dir(ruta) = "" Then
    ws.Cells(fila, 3).Value = "No encontrado"
Else
    Set wb = Workbooks.Open(ruta)
    ' ... copiar ...
    wb.Close SaveChanges:=False
End If`},

{id:"guardar", cat:"hojas", t:"Guardar un libro nuevo con nombre",
 fr:["guardar cada libro","guardar como","libro nuevo","reemplazandolo si ya existe"],
 q:"Workbooks.Add crea un libro; SaveAs lo guarda (con avisos apagados reemplaza sin preguntar).",
 code:`Set wbN = Workbooks.Add
Application.DisplayAlerts = False
wbN.SaveAs ThisWorkbook.Path & "\\EstadoCuenta_" & codigo & ".xlsx"
Application.DisplayAlerts = True
wbN.Close`},

/* ---------- Formato y colores ---------- */
{id:"color", cat:"formato", t:"Resaltar en un color y quitar el relleno",
 fr:["amarillo","resaltado","resaltar","resalte","rellenar","colorear","rojo claro","celeste","quita el relleno","quitar el relleno","resaltar"],
 q:"Interior.Color pinta; ColorIndex = xlNone quita. Quitá en el caso correcto para limpiar corridas anteriores.",
 code:`cel.Interior.Color = vbYellow
cel.Interior.Color = RGB(255, 199, 206) '1 rojo claro
cel.Interior.ColorIndex = xlNone '1 sin relleno`,
 ver:"A1:ValidarPlacas"},

{id:"fila", cat:"formato", t:"Pintar la fila entera",
 fr:["toda la fila","pinta la fila","rellena las columnas","columnas a a"],
 q:"Armás el rango como texto: \"A5:G5\".",
 code:`ws.Range("A" & fila & ":G" & fila).Interior.Color = RGB(221, 235, 247)`,
 ver:"A3:PintarFila"},

{id:"negrita", cat:"formato", t:"Fila Total en negrita",
 fr:["en negrita","fila total","una fila total","fila de totales","agregar debajo"],
 q:"La fila siguiente a la última es ultima + 1.",
 code:`ws.Cells(ultima + 1, 1).Value = "Total"
ws.Cells(ultima + 1, 7).Value = total
ws.Rows(ultima + 1).Font.Bold = True`},

{id:"numfmt", cat:"formato", t:"Formato de moneda, porcentaje o fecha en la celda",
 fr:["formato de moneda","formato de numero","formatos de numero","formato de porcentaje","dar formato"],
 q:"NumberFormat cambia cómo se ve, no el valor.",
 code:`ws.Range("G2:G" & ultima).NumberFormat = "$#,##0.00"
ws.Range("D2:D" & fila).NumberFormat = "0.0%"`},

{id:"autofit", cat:"formato", t:"Ajustar el ancho de las columnas",
 fr:["ajustar el ancho","ancho de las columnas","autofit"],
 q:"AutoFit sobre las columnas usadas.",
 code:`wsR.Columns("A:D").AutoFit`},

/* ---------- Funciones y procedimientos ---------- */
{id:"funcion", cat:"proc", t:"Una función propia (también como fórmula)",
 fr:["funcion","devuelve","se usa desde vba y como formula","como formula en la hoja","se escribe una sola vez"],
 q:"Function calcula y devuelve: el valor se asigna al nombre de la función. Pública y en un módulo normal sirve como =Nombre() en la hoja.",
 code:`Public Function ConIVA(ByVal monto As Double) As Double
    ConIVA = Application.WorksheetFunction.Round(monto * 1.13, 2) '1 "return"
End Function`,
 tr:"#NAME? en la hoja: la función está en el módulo de una hoja o el nombre está mal.", ver:"A3:Comision"},

{id:"opcional", cat:"proc", t:"Parámetro opcional",
 fr:["opcional","vale 0 si no se envia","si no se envia","optional"],
 q:"Optional con valor por defecto, siempre al final de la lista.",
 code:`Function Comision(ByVal prima As Double, ByVal ramo As String, _
                  Optional ByVal bono As Double = 0) As Double`,
 ver:"A3:Comision"},

{id:"privado", cat:"proc", t:"Función o Sub privado (no aparece en Alt+F8)",
 fr:["privada","privado","no aparece","alt+f8","alt f8"],
 q:"Private: solo lo usan los procedimientos del mismo módulo.",
 code:`Private Function EsPrimaAlta(ByVal prima As Double) As Boolean
    EsPrimaAlta = (prima >= 1000)
End Function`,
 ver:"A3:EsPrimaAlta"},

{id:"llamar", cat:"proc", t:"Llamar un Sub que recibe parámetros",
 fr:["sub privado","que rellena","recibe","pintarfila(","llamar"],
 q:"Sin paréntesis, o con Call y paréntesis.",
 code:`PintarFila ws, fila, RGB(221, 235, 247)
Call PintarFila(ws, fila, RGB(221, 235, 247)) '1 equivalente`,
 tr:"PintarFila(ws, fila, color) sin Call da error de sintaxis.", ver:"A3:ResaltarPrimasAltas"},

{id:"probar", cat:"proc", t:"Una macro para probar la función",
 fr:["pruebela con una macro","probar","macro probar","pruebe"],
 q:"Un Sub que llama a la función con los casos del enunciado y los imprime.",
 code:`Sub ProbarComision()
    Debug.Print Comision(850, "Auto") '1 debe dar 68
End Sub`,
 ver:"A3:ProbarComision"},

/* ---------- Mostrar resultados ---------- */
{id:"immediate", cat:"mostrar", t:"Imprimir en la ventana Immediate",
 fr:["ventana immediate","immediate","imprime","imprimir","debug.print"],
 q:"Debug.Print escribe en Immediate (Ctrl+G para verla).",
 code:`Debug.Print "Tickets procesados: " & (ultima - 1) '! ultima es la fila, no la cantidad
Debug.Print "Total: " & Format(total, "$#,##0.00")`,
 ver:"A1:CalcularParqueo"},

{id:"msg", cat:"mostrar", t:"Un MsgBox con un dato por línea",
 fr:["un solo mensaje","cada dato en su linea","una linea por dato","msgbox","un mensaje con","mostrar un mensaje"],
 q:"Los datos se unen con & vbCrLf &; Format les da formato de moneda u hora.",
 code:`MsgBox "Ticket: " & t & vbCrLf & _
       "Pagar: " & Format(pagar, "$#,##0.00"), vbInformation, "Resultado"`,
 ver:"A1:BuscarTicket"},

{id:"icono", cat:"mostrar", t:"Mensaje de advertencia o de información",
 fr:["de advertencia o de informacion","advertencia","de informacion","avisar"],
 q:"vbExclamation si hay problemas, vbInformation si todo está bien.",
 code:`If malas > 0 Then
    MsgBox "Hay " & malas & " incorrectas.", vbExclamation
Else
    MsgBox "Todas correctas.", vbInformation
End If`,
 ver:"A1:ValidarPlacas"},

{id:"status", cat:"mostrar", t:"Mostrar el avance en la barra de estado",
 fr:["barra de estado","mostrar el avance","avance"],
 q:"StatusBar escribe abajo; al terminar se devuelve con False.",
 code:`Application.StatusBar = "Procesando " & i & " de " & n
' ... al terminar:
Application.StatusBar = False`}
];

/* Enunciados de ejemplo: los 4 problemas de la ayudantía */
window.EJEMPLOS = [
 {id:"A1", t:"A1 Parqueo", txt:"1. CalcularParqueo. Escribir en las columnas F a I, con sus encabezados, los minutos, el cobro, el descuento y el valor a pagar de cada ticket. La macro encuentra sola la última fila. Al final imprime en la ventana Immediate cuántos tickets procesó, cuántos fueron cortesía y el total a cobrar. Hasta 15 minutos: gratis. Más de 15 minutos: $1.00 por la primera hora y $0.50 por cada media hora adicional o fracción. El cobro máximo del día es $6.00. Con sello de compra se descuenta $1.00, sin que el pago quede negativo.\n2. BuscarTicket. Pedir el número de ticket con un InputBox; debe aceptarse aunque se escriba en minúsculas o con espacios. Si el usuario cancela, no pasa nada. Si el ticket no existe, se avisa. Si existe, un solo mensaje muestra ticket, placa, entrada, salida, minutos y valor a pagar, cada dato en su línea.\n3. ValidarPlacas. Las placas deben tener la forma P###-###. Recorrer la columna B con For Each, rellenar de amarillo las incorrectas e imprimir su ticket; a las correctas se les quita el relleno. Un MsgBox de advertencia o de información dice cuántas hay."},
 {id:"A2", t:"A2 Proveedores", txt:"1. LimpiarContactos. En G, el contacto sin espacios al inicio ni al final, sin espacios dobles o triples en medio y con la primera letra de cada palabra en mayúscula. En H, el correo en minúsculas y sin espacios alrededor. En I, el dominio: lo que va después de la @. Si el correo no tiene @ o tiene más de una, se escribe \"Revisar\" resaltado en amarillo.\n2. VencimientosProveedores. En J, la fecha de vencimiento (fecha de la factura más los días de crédito, pasada al lunes si cae en fin de semana). En K, los días de atraso a la fecha de corte (0 si todavía no vence). En L, el estado: \"Al día\", \"Vencida 1-30\" en amarillo o \"Vencida +30\" en rojo claro. La fecha de corte se lee de la hoja de parámetros. Al final, un MsgBox con el número de facturas vencidas, el monto vencido y la fecha de corte."},
 {id:"A3", t:"A3 Pólizas", txt:"1. Función Comision. Devuelve prima × tasa + bono, redondeado a 2 decimales. El bono es opcional y vale 0 si no se envía. Si la prima no es mayor que 0 o el ramo no se reconoce, devuelve 0. Pruébela con una macro ProbarComision. Las tasas son Auto 8%, Vida 12%, Hogar 10% y Salud 9%; cualquier otro ramo no paga comisión. El ramo puede venir con espacios o en mayúsculas.\n2. LlenarComisiones. Con una sola instrucción por columna, poner en F la fórmula =Comision(D2,C2) y en G la comisión con un bono de $25 para las pólizas nuevas. Encabezados \"Comisión\" y \"Comisión con bono\".\n3. EsPrimaAlta(prima): función privada, True si la prima es de 1,000 o más. PintarFila(ws, fila, colorFila): Sub privado que rellena las columnas A a G de una fila. ResaltarPrimasAltas: macro pública que pinta de celeste las pólizas de prima alta, quita el relleno a las demás e imprime cuántas son."},
 {id:"A4", t:"A4 Pedidos", txt:"Cada macro lee la hoja una sola vez con datos = ws.Range(\"A2:F\" & ultima).Value y trabaja con el arreglo.\n1. TotalesPedidos. Calcular el total de cada pedido (unidades × precio, a 2 decimales) en un arreglo y escribir todos los totales de una sola vez en la columna G, con el encabezado \"Total\". Imprimir cuántos pedidos hay y cuánto suman.\n2. ClientesDistintos. Armar una Collection con los clientes sin repetir, en el orden en que aparecen, e imprimir cuántos son y sus nombres.\n3. ResumenPorCategoria. Con un Dictionary, acumular las ventas y el número de pedidos de cada categoría en una sola pasada. Crear al final del libro la hoja \"A4 Resumen\"; si ya existe de una ejecución anterior, se elimina sin preguntar. La tabla lleva Categoría, Pedidos, Ventas y % del total, una fila Total en negrita y formatos de número y porcentaje. Al final, un MsgBox con la categoría que más vendió. Las categorías no se escriben en el código."}
];

/* =====================================================================
   FASE 3: frases de la clase (Ejercicio_Clase_Problemas) y de la guía.
   src = dónde aparece. Estas piezas no tienen plantilla completa.
   ===================================================================== */
window.FRASES.push(

{id:"divmod", cat:"calc", t:"Cuántos completos y cuántos sobran (\\ y Mod)",
 fr:["cajas completas lleva","unidades sueltas","van sueltas","botellas sueltas","operadores \\ y mod","mod","cuantas cajas completas","fardos","paquetes de"],
 q:"\\ da cuántas veces cabe completo; Mod da lo que sobra. Para varios tamaños, se encadena con lo que sobra.",
 code:`cajas = cantidad \\ 12 '1 cajas completas
sueltas = cantidad Mod 12 '1 lo que sobra
' Varios tamaños (fardos de 24, paquetes de 6):
fardos = botellas \\ 24
paquetes = (botellas Mod 24) \\ 6
sueltas = (botellas Mod 24) Mod 6`,
 tr:"\"Se redondea hacia arriba a cajas enteras\" (comprar lo que falta) es otra cosa: -Int(-x / 12).", src:"Clase E1b, Guía E2"},

{id:"find", cat:"pedir", t:"Buscar con Find (y revisar Nothing)",
 fr:["use find","con find","lookat","xlwhole","es nothing","sea nothing","revise si el resultado"],
 q:"Find devuelve la celda encontrada o Nothing. Con Offset llegás a los otros datos de esa fila.",
 code:`Dim c As Range
Set c = ws.Range("A:A").Find(What:=resp, LookAt:=xlWhole) '1 celda completa igual
If c Is Nothing Then '! revisar ANTES de usar c
    MsgBox "El pedido " & resp & " no existe.", vbExclamation
    Exit Sub
End If
MsgBox "Cliente: " & c.Offset(0, 2).Value & vbCrLf & _
       "Total: " & Format(ws.Cells(c.Row, 11).Value, "#,##0.00")`,
 src:"Clase E1d, Guía E6"},

{id:"meta", cat:"calc", t:"Acumular hasta llegar a una meta",
 fr:["se detenga apenas","apenas llegue","llegue a la meta","se alcanzo","alcanza la meta","no se alcanza","meta de ventas","exit for"],
 q:"Sumás fila por fila; al llegar, guardás la fila y salís con Exit For. Si la fila guardada sigue en 0, no se alcanzó.",
 code:`filaMeta = 0
For fila = 2 To ultima
    acum = acum + ws.Cells(fila, 11).Value
    If acum >= meta Then
        filaMeta = fila
        Exit For
    End If
Next fila
If filaMeta = 0 Then
    Debug.Print "La meta no se alcanzó. Acumulado: " & acum
Else
    Debug.Print "Meta con " & ws.Cells(filaMeta, 1).Value & " el " & _
        Format(ws.Cells(filaMeta, 2).Value, "dd/mm/yyyy") & _
        ", acumulado " & Format(acum, "#,##0.00") & ", revisados " & (filaMeta - 1)
End If`,
 src:"Clase E1e"},

{id:"split", cat:"texto", t:"Separar un código en partes (Split)",
 fr:["con split","se separa con split","a partir del codigo","separar","el departamento","el pais"],
 q:"Split parte el texto por un separador; las partes empiezan en 0.",
 code:`partes = Split(ws.Cells(fila, 1).Value, "-") '1 "VEN-GT-031"
ws.Cells(fila, 7).Value = partes(0) '1 VEN
ws.Cells(fila, 8).Value = partes(1) '1 GT`,
 tr:"Split empieza en 0. partes(2) sería el correlativo.", src:"Clase E2b"},

{id:"dobles", cat:"texto", t:"Quitar espacios dobles con Replace (sin WorksheetFunction)",
 fr:["trim, replace","replace y strconv","strconv","vbpropercase"],
 q:"Si piden usar Replace y StrConv: reemplazar \"  \" por \" \" hasta que no quede ninguno, y luego StrConv.",
 code:`x = Trim(ws.Cells(fila, 2).Value)
Do While InStr(x, "  ") > 0 '1 mientras haya dos espacios seguidos
    x = Replace(x, "  ", " ")
Loop
x = StrConv(x, vbProperCase)
Debug.Print "[" & x & "]" '1 corchetes para ver si quedaron espacios`,
 src:"Clase E2a"},

{id:"anios", cat:"fecha", t:"Años de servicio o antigüedad",
 fr:["anos de servicio","antiguedad","ano del corte menos","mayor antiguedad"],
 q:"Año de una fecha menos año de la otra (como lo pide la clase: sin revisar si ya cumplió).",
 code:`anios = Year(corte) - Year(ws.Cells(fila, 4).Value)`,
 tr:"Si piden años CUMPLIDOS: DateDiff(\"yyyy\", ingreso, corte) y restar 1 si todavía no llega el día y mes del aniversario.", src:"Clase E2c"},

{id:"aniversario", cat:"fecha", t:"La misma fecha en otro año (DateSerial)",
 fr:["mismo dia y mes","dateserial","aniversario","se celebra"],
 q:"DateSerial arma una fecha con el año que querés y el mes y día de la original.",
 code:`ing = ws.Cells(fila, 4).Value
aniv = DateSerial(2026, Month(ing), Day(ing))
If Weekday(aniv) = 7 Then aniv = aniv + 2 '1 sábado -> lunes
If Weekday(aniv) = 1 Then aniv = aniv + 1 '1 domingo -> lunes
ws.Cells(fila, 11).Value = aniv
ws.Cells(fila, 11).NumberFormat = "dd/mm/yyyy"`,
 tr:"Un 29 de febrero en un año que no es bisiesto, DateSerial lo pasa al 1 de marzo.", src:"Clase E2c"},

{id:"esnum", cat:"calc", t:"Validar que un valor sea número (y no esté vacío)",
 fr:["si el salario no es un numero","no es un numero","isnumeric","celda vacia tambien","valide con isnumeric"],
 q:"IsNumeric de una celda vacía da True: hay que revisar también que no esté vacía.",
 code:`v = ws.Cells(fila, 5).Value
If IsEmpty(v) Or Not IsNumeric(v) Then
    ws.Cells(fila, 12).Value = "Revisar"
    ws.Cells(fila, 5).Interior.Color = vbYellow '1 marcar el dato malo
Else
    ws.Cells(fila, 12).Value = Application.WorksheetFunction.Round(v * 0.02 * anios, 2)
End If`,
 src:"Clase E2d"},

{id:"wsfResumen", cat:"mostrar", t:"Resumen con Sum, CountIf y Max",
 fr:["worksheetfunction (sum","countif","con worksheetfunction","sum, countif y max","cantidad de empleados de"],
 q:"Las funciones de Excel sobre columnas completas, sin bucle.",
 code:`Dim wf As WorksheetFunction
Set wf = Application.WorksheetFunction '1 atajo para no repetir
totalBonos = wf.Sum(ws.Range("L2:L" & ultima)) '1 Sum ignora los "Revisar"
enSV = wf.CountIf(ws.Range("H2:H" & ultima), "SV")
mayor = wf.Max(ws.Range("J2:J" & ultima))`,
 src:"Clase E2d"},

{id:"cuota", cat:"calc", t:"Cuota fija de un préstamo",
 fr:["cuota fija","cuota mensual","sistema frances","tasa mensual","tasa anual / 12","(1 + r)","prestamo","amortizacion"],
 q:"r = tasa anual / 12. cuota = monto × r / (1 − (1 + r)^−n). ^ es potencia.",
 code:`r = tasaAnual / 12
cuota = Application.WorksheetFunction.Round(monto * r / (1 - (1 + r) ^ (-meses)), 2) '! paréntesis en el exponente negativo`,
 tr:"No uses \"cuota\" como nombre de variable si tu función se llama Cuota.", src:"Clase E3a, Guía E7"},

{id:"pmt", cat:"hojas", t:"Fórmula de verificación con PMT",
 fr:["pmt","verificacion","-pmt"],
 q:"PMT es la función de Excel para la cuota (sale negativa). En .Formula va en inglés y con comas.",
 code:`ws.Range("H2:H" & ultima).Formula = "=ROUND(-PMT(D2/12,E2,C2),2)"`,
 src:"Clase E3b"},

{id:"ultimaFn", cat:"proc", t:"Función que devuelve la última fila",
 fr:["ultimafila(ws)","funcion publica ultimafila","ultimafila"],
 q:"Una función que recibe una hoja y devuelve un Long. Se reutiliza en todas las macros.",
 code:`Public Function UltimaFila(ByVal ws As Worksheet) As Long
    UltimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
End Function
' uso:
ultima = UltimaFila(ws)`,
 src:"Clase E3b"},

{id:"nombresArg", cat:"proc", t:"Nombres de argumentos que no se deben usar",
 fr:["no use \"color\"","como nombre de argumento","no use color"],
 q:"No uses el nombre de tu función ni palabras de VBA (Color, Value, Name) como parámetro.",
 code:`Private Sub PintarFila(ByVal ws As Worksheet, ByVal fila As Long, ByVal colorFila As Long) '1 colorFila, no color`,
 src:"Clase E3c"},

{id:"preserve", cat:"agrupar", t:"Lista de tamaño desconocido (ReDim Preserve)",
 fr:["arreglo dinamico","redim preserve","lista numerada","si no hubiera ninguno"],
 q:"Cada vez que encontrás uno, agrandás el arreglo en 1 sin perder lo anterior.",
 code:`Dim lista() As String, n As Long
For i = 1 To UBound(datos, 1)
    If datos(i, 7) = "Pendiente" Then
        n = n + 1
        ReDim Preserve lista(1 To n)
        lista(n) = datos(i, 1)
    End If
Next i
If n = 0 Then
    Debug.Print "No hay pendientes." '! si n = 0, no usar lista
Else
    Debug.Print "Pendientes: " & n
    For i = 1 To n
        Debug.Print i & ". " & lista(i)
    Next i
End If`,
 src:"Clase E4b"},

{id:"colClave", cat:"agrupar", t:"Sin repetir con Collection y clave (On Error)",
 fr:["como clave","clave repetida","on error resume next","usando el nombre del cliente como clave"],
 q:"Se agrega cada valor con él mismo como clave; si la clave ya existe da error y ese repetido no entra.",
 code:`Set col = New Collection
On Error Resume Next
For i = 1 To UBound(datos, 1)
    col.Add datos(i, 4), CStr(datos(i, 4)) '1 valor, clave (la clave debe ser texto)
Next i
On Error GoTo 0 '! encender los errores apenas termina`,
 src:"Clase E4c"},

{id:"existe", cat:"agrupar", t:"Revisar si existe una clave",
 fr:["revise con exists","con exists","si existe una sucursal","exists"],
 q:"Exists devuelve True o False sin dar error.",
 code:`If d.Exists("Sonsonate") Then
    Debug.Print "Sonsonate: " & d("Sonsonate")
Else
    Debug.Print "No existe la sucursal Sonsonate."
End If`,
 src:"Clase E4d"},

{id:"resize", cat:"base", t:"Pegar un arreglo con Resize",
 fr:["resize","range(\"h2\").resize"],
 q:"Resize agranda la celda inicial al tamaño del arreglo: no hace falta calcular la última fila.",
 code:`ws.Range("H2").Resize(n, 1).Value = totales '1 n filas, 1 columna`,
 src:"Clase E4a"},

{id:"cadaMes", cat:"fecha", t:"Mismo día de cada mes (cuotas)",
 fr:["mismo dia de cada mes","cada mes siguiente","dateadd","fecha de pago"],
 q:"DateAdd con \"m\" suma meses respetando el día.",
 code:`fechaPago = DateAdd("m", k, desembolso) '1 k = número de cuota`,
 src:"Guía E7"},

{id:"alFinal", cat:"hojas", t:"Agregar un registro al final de otra hoja",
 fr:["agregarla al final","al final del registro","agregar al final","siguiente fila vacia","registro de"],
 q:"La siguiente fila libre es la última fila + 1.",
 code:`sig = wsReg.Cells(wsReg.Rows.Count, "A").End(xlUp).Row + 1
wsReg.Cells(sig, 1).Value = Date
wsReg.Cells(sig, 2).Value = codigo`,
 src:"Guía E6"},

{id:"motivos", cat:"texto", t:"Juntar varios motivos separados por punto y coma",
 fr:["separados por punto y coma","todos los motivos","todas las observaciones","observaciones que apliquen"],
 q:"Vas agregando cada motivo a un texto; al final quitás el \"; \" sobrante.",
 code:`obs = ""
If monto > limite Then obs = obs & "Excede límite; "
If comprob = "No" Then obs = obs & "Sin comprobante; "
If Len(obs) > 0 Then obs = Left(obs, Len(obs) - 2) '1 quitar el último "; "`,
 src:"Guía E12, E18"},

{id:"duplicado", cat:"agrupar", t:"Detectar duplicados",
 fr:["duplicado","registrado dos veces","registrados dos veces","ya hubo antes","ya esta en","repetidos"],
 q:"Un Dictionary con lo ya visto. La clave puede juntar varios campos con un separador.",
 code:`Set vistos = CreateObject("Scripting.Dictionary")
For fila = 2 To ultima
    k = ws.Cells(fila, 3).Value & "|" & ws.Cells(fila, 2).Value & "|" & ws.Cells(fila, 5).Value '1 empleado|fecha|monto
    If vistos.Exists(k) Then
        ' es duplicado
    Else
        vistos.Add k, fila
    End If
Next fila`,
 src:"Guía E12, E18"},

{id:"claveDoble", cat:"agrupar", t:"Agrupar por dos cosas a la vez (vendedor y mes)",
 fr:["por vendedor y mes","una fila por vendedor y mes","por empleado, fecha","mismo dia y categoria"],
 q:"La clave del Dictionary junta los dos campos: \"Ana|2026-03\".",
 code:`k = ws.Cells(fila, 3).Value & "|" & Format(ws.Cells(fila, 1).Value, "yyyy-mm")
If Not d.Exists(k) Then d.Add k, 0
d(k) = d(k) + ws.Cells(fila, 5).Value
' al leer: partes = Split(k, "|")`,
 src:"Guía E10, E18"},

{id:"diaPorDia", cat:"fecha", t:"Recorrer día por día (noches de estadía)",
 fr:["noche por noche","dia por dia","recorrer dia por dia","cada noche"],
 q:"Un For con fechas: cada vuelta es un día. Las noches van de la llegada a la salida − 1.",
 code:`For d = llegada To salida - 1 '1 la noche de la salida no se cobra
    If Weekday(d, vbMonday) >= 5 And Weekday(d, vbMonday) <= 6 Then
        ' viernes o sábado: recargo
    End If
Next d`,
 src:"Guía E13, E19"},

{id:"enRango", cat:"fecha", t:"¿La fecha cae dentro de algún periodo?",
 fr:["cae dentro de","dentro de alguna temporada","temporada alta","feriado","entre las fechas"],
 q:"Recorrés la tabla de periodos y marcás si la fecha está entre inicio y fin de alguno.",
 code:`alta = False
For t = 2 To ultT
    If d >= wsT.Cells(t, 1).Value And d <= wsT.Cells(t, 2).Value Then
        alta = True
        Exit For
    End If
Next t`,
 src:"Guía E13, E19"},

{id:"tablaTramos", cat:"calc", t:"Buscar el tramo en una tabla (renta)",
 fr:["tabla por tramos","buscando el tramo","sobre el exceso","cuota fija","tramo donde cae"],
 q:"Recorrés la tabla de tramos y aplicás el que contiene la base.",
 code:`For t = 2 To 5
    If base >= wsR.Cells(t, 1).Value And base <= wsR.Cells(t, 2).Value Then
        renta = (base - wsR.Cells(t, 4).Value) * wsR.Cells(t, 3).Value + wsR.Cells(t, 5).Value
        Exit For
    End If
Next t`,
 src:"Guía E16"},

{id:"fechaTexto", cat:"fecha", t:"Fecha que viene como texto (validar que exista)",
 fr:["dos formatos","dd/mm/aaaa","aaaa-mm-dd","existir en el calendario","fechas imposibles","fecha (texto)"],
 q:"Partís el texto, armás la fecha con DateSerial y comprobás que el día y el mes sigan iguales (31/02 se convertiría en marzo).",
 code:`If InStr(txt, "/") > 0 Then
    p = Split(txt, "/"): dd = Val(p(0)): mm = Val(p(1)): aa = Val(p(2))
Else
    p = Split(txt, "-"): aa = Val(p(0)): mm = Val(p(1)): dd = Val(p(2))
End If
f = DateSerial(aa, mm, dd)
valida = (Day(f) = dd And Month(f) = mm And Year(f) = aa)`,
 tr:"No confiés en IsDate ni CDate: dependen de la configuración regional.", src:"Guía E12"},

{id:"montoTexto", cat:"texto", t:"Monto con $ y comas como texto",
 fr:["signo de dolar","separador de miles","trae $","con $"],
 q:"Quitás $ y comas con Replace y después validás.",
 code:`m = Replace(Replace(Trim(txt), "$", ""), ",", "")
If IsNumeric(m) Then monto = Val(m) Else ' error`,
 tr:"IsNumeric(\"$47.36\") puede dar True según la configuración regional, pero Val(\"$47.36\") da 0: limpiá primero.", src:"Guía E12"},

{id:"vacia", cat:"base", t:"Contar o saltar filas vacías",
 fr:["filas vacias","completamente vacias","fila vacia"],
 q:"CountA de la fila da 0 si está vacía.",
 code:`If Application.WorksheetFunction.CountA(ws.Rows(fila)) = 0 Then
    vacias = vacias + 1
Else
    ' validar la fila
End If`,
 src:"Guía E12"},

{id:"top", cat:"calc", t:"Los tres mayores (top 3)",
 fr:["los tres","top 3","mas vendidos","tres productos"],
 q:"Escribís el resumen en la hoja y lo ordenás de mayor a menor; los tres primeros son el top. O usás Large(rango, 1, 2, 3).",
 code:`wsR.Range("A2:B" & fila - 1).Sort Key1:=wsR.Range("B2"), Order1:=xlDescending, Header:=xlNo
' o el valor: Application.WorksheetFunction.Large(wsR.Range("B2:B13"), 2)`,
 src:"Guía E20"},

{id:"orden", cat:"hojas", t:"Crear hojas en orden alfabético",
 fr:["en este orden","orden alfabetico","ordenadas"],
 q:"Las claves del Dictionary salen en orden de aparición. Para orden alfabético, pasalas a una hoja auxiliar y ordenalas, o ordená un arreglo.",
 code:`claves = d.Keys '1 arreglo desde 0
For i = 0 To UBound(claves) - 1 '1 burbuja: suficiente para pocas claves
    For j = i + 1 To UBound(claves)
        If claves(j) < claves(i) Then tmp = claves(i): claves(i) = claves(j): claves(j) = tmp
    Next j
Next i`,
 src:"Guía E11"},

{id:"prefijo", cat:"hojas", t:"Borrar todas las hojas que empiezan con un prefijo",
 fr:["primero se eliminan","hojas de region de la ejecucion anterior","hojas de la ejecucion anterior"],
 q:"Recorré las hojas de atrás hacia adelante y borrá las que empiezan con el prefijo.",
 code:`Application.DisplayAlerts = False
For i = ThisWorkbook.Worksheets.Count To 1 Step -1 '! al borrar, recorrer hacia atrás
    If Left(ThisWorkbook.Worksheets(i).Name, 6) = "E11 R-" Then ThisWorkbook.Worksheets(i).Delete
Next i
Application.DisplayAlerts = True`,
 src:"Guía E11"},

{id:"rojoTexto", cat:"formato", t:"Texto en rojo (advertencia)",
 fr:["advertencia en rojo","en rojo debajo","letra roja"],
 q:"Font.Color cambia el color de la letra (Interior es el relleno).",
 code:`ws.Cells(fila + 2, 1).Value = "Atención: el saldo supera el límite de crédito."
ws.Cells(fila + 2, 1).Font.Color = vbRed
ws.Cells(fila + 2, 1).Font.Bold = True`,
 src:"Guía E15"},

{id:"rapido", cat:"base", t:"Que la macro corra rápido con muchos datos",
 fr:["mas de 9,000","miles de","rendimiento","screenupdating"],
 q:"Apagás el redibujo de pantalla al inicio y lo encendés al final.",
 code:`Application.ScreenUpdating = False
' ... todo el trabajo ...
Application.ScreenUpdating = True`,
 src:"Guía E20, segmento 14"},

{id:"evento", cat:"proc", t:"Reaccionar cuando el usuario cambia una celda",
 fr:["cada vez que","al cambiar","worksheet_change","cuando el usuario"],
 q:"Un evento en el módulo de la hoja: se ejecuta solo. Intersect limita a ciertas columnas.",
 code:`Private Sub Worksheet_Change(ByVal Target As Range)
    If Intersect(Target, Me.Range("G2:G5000")) Is Nothing Then Exit Sub
    Application.EnableEvents = False '! si no, tu propia escritura vuelve a disparar el evento
    Target.Offset(0, 10).Value = Now
    Application.EnableEvents = True
End Sub`,
 src:"Segmento 14"}
);
