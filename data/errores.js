/* Fase 4: Índice de errores y checklist.
   tipo: "comp" = de compilación (VBA no deja correr), "run" = en ejecución (con número),
         "mudo" = sin mensaje, pero el resultado sale mal.
   msg = el texto tal como lo muestra Excel en inglés; es = qué quiere decir.
   causas / arreglo = listas; fn = fichas relacionadas. */

window.ERRORES = [

/* ---------- De compilación ---------- */
{id:"notdef", tipo:"comp", msg:"Compile error: Variable not defined", es:"Usaste una variable que no declaraste con Dim (y tenés Option Explicit).",
 causas:["Te faltó el Dim de esa variable.","La escribiste distinto: totl en vez de total.","Copiaste una pieza que usa una variable con otro nombre que la tuya."],
 arreglo:["La palabra marcada en azul es la culpable: agregá su Dim arriba del Sub o corregí el nombre.","Si viene de una pieza copiada, renombrala igual que en tu macro."],
 code:"Dim cortesias As Long", fn:["dim","option"]},

{id:"subnotdef", tipo:"comp", msg:"Compile error: Sub or Function not defined", es:"Llamaste a algo que VBA no encuentra.",
 causas:["Nombre mal escrito (Comsion en vez de Comision).","La función es Private y está en otro módulo.","Usaste una función de Excel sin WorksheetFunction: Max(a, b) en vez de Application.WorksheetFunction.Max(a, b).","Escribiste en español: SI, SUMA, REDONDEAR no existen en VBA."],
 arreglo:["Revisá el nombre marcado.","Poné las Private en el mismo módulo que las usa.","Las de Excel van con Application.WorksheetFunction."], fn:["private",".worksheetfunction"]},

{id:"ambiguous", tipo:"comp", msg:"Compile error: Ambiguous name detected: Nombre", es:"Hay dos macros o funciones con el mismo nombre.",
 causas:["Importaste o pegaste la misma plantilla dos veces.","Dos módulos tienen un Sub con el mismo nombre (ClientesDistintos en A4 y en un ejercicio de práctica)."],
 arreglo:["Borrá uno de los dos o cambiale el nombre.","Si importaste un .bas dos veces, eliminá el módulo repetido: clic derecho → Remove."]},

{id:"invalidout", tipo:"comp", msg:"Compile error: Invalid outside procedure", es:"Hay una instrucción fuera de un Sub o Function.",
 causas:["Pegaste código por encima del Sub o después del End Sub.","Un End Sub quedó antes de tiempo."],
 arreglo:["Fuera de los Sub solo pueden ir Option Explicit, Const y Dim del módulo. Mové lo demás adentro."], fn:["sub"]},

{id:"blockif", tipo:"comp", msg:"Compile error: Block If without End If  /  End If without block If", es:"Un If de varias líneas no tiene su End If, o sobra uno.",
 causas:["Olvidaste el End If.","Escribiste If … Then con código en la misma línea y además pusiste End If."],
 arreglo:["Cada If de varias líneas cierra con End If. El If en una sola línea (If x Then y = 1) no lleva End If.","Indentá: así se ve rápido cuál falta."], fn:["if"]},

{id:"fornext", tipo:"comp", msg:"Compile error: For without Next  /  Next without For", es:"Un For no tiene su Next, o un Next no tiene su For.",
 causas:["Falta el Next.","Los Next están cruzados en bucles anidados (Next i antes de Next j).","A veces lo dispara un End If faltante dentro del bucle."],
 arreglo:["Cerrá cada bucle con su Next en orden inverso: el último que abrís es el primero que cerrás.","Revisá también los If de adentro."], fn:["for"]},

{id:"expected", tipo:"comp", msg:"Compile error: Expected: =", es:"Llamaste a un Sub con parámetros usando paréntesis.",
 causas:["Escribiste PintarFila(ws, fila, color) sin Call adelante."],
 arreglo:["Quitá los paréntesis: PintarFila ws, fila, color","O agregá Call: Call PintarFila(ws, fila, color)"], code:"PintarFila ws, fila, RGB(221, 235, 247)", fn:["call"]},

{id:"eos", tipo:"comp", msg:"Compile error: Expected: end of statement  /  Syntax error", es:"La línea está mal escrita (sale en rojo).",
 causas:["Comillas sin cerrar.","Falta un & entre texto y variable.","Usaste comillas tipográficas “ ” copiadas de un PDF o Word.","Usaste ; en vez de , entre argumentos de VBA."],
 arreglo:["Revisá comillas y & en esa línea.","Reemplazá las comillas curvas por comillas rectas \".","En VBA los argumentos siempre van separados por coma."], fn:["&"]},

{id:"method", tipo:"comp", msg:"Compile error: Method or data member not found", es:"Esa propiedad o método no existe para ese objeto.",
 causas:["Error de digitación: .Valeu, .Interor, .Colr.","El objeto es de otro tipo del que creés."],
 arreglo:["Escribí el punto y esperá la lista de IntelliSense: elegí de ahí.","Revisá cómo declaraste la variable."]},

{id:"argnot", tipo:"comp", msg:"Compile error: Argument not optional  /  Wrong number of arguments", es:"Le mandaste menos (o más) parámetros de los que pide.",
 causas:["Llamaste Comision(850) sin el ramo.","Tu Function tiene parámetros que no son Optional."],
 arreglo:["Mandá todos los parámetros obligatorios, en el orden de la declaración.","Si uno puede faltar, declaralo Optional con valor por defecto, al final."], fn:["optional"]},

{id:"byref", tipo:"comp", msg:"Compile error: ByRef argument type mismatch", es:"Pasaste una variable de un tipo distinto al que espera el parámetro.",
 causas:["El parámetro es As Long y le pasás una variable Variant o Double."],
 arreglo:["Declarás el parámetro ByVal (acepta conversión), o declarás tu variable del mismo tipo."], fn:["byval"]},

{id:"dupdecl", tipo:"comp", msg:"Compile error: Duplicate declaration in current scope", es:"Declaraste dos veces la misma variable.",
 causas:["Hay dos Dim de la misma variable (pasa al pegar piezas que traen sus propios Dim)."],
 arreglo:["Borrá uno de los Dim. Juntá todos los Dim al inicio del Sub."], fn:["dim"]},

{id:"label", tipo:"comp", msg:"Compile error: Label not defined", es:"Un GoTo apunta a una etiqueta que no existe.",
 causas:["Escribiste GoTo Invalido pero la etiqueta se llama distinto o le falta los dos puntos."],
 arreglo:["La etiqueta va sola en su línea y termina en dos puntos: Invalido:"], fn:["goto"]},

/* ---------- En ejecución ---------- */
{id:"e5", tipo:"run", num:5, msg:"Run-time error '5': Invalid procedure call or argument", es:"Una función recibió un valor imposible.",
 causas:["Mid, Left o Right con un número negativo o 0: pasa cuando InStr dio 0 porque no encontró la @.","Weekday o DateSerial con valores raros."],
 arreglo:["Antes de cortar, revisá que InStr sea > 0.","Pasá el mouse sobre las variables (o F8) para ver qué valor llegó."], code:"If InStr(correo, \"@\") > 0 Then dominio = Mid(correo, InStr(correo, \"@\") + 1)", fn:["instr","mid","left"]},

{id:"e6", tipo:"run", num:6, msg:"Run-time error '6': Overflow", es:"El número no cabe en el tipo de la variable.",
 causas:["Declaraste As Integer y el valor pasó de 32,767.","Una división entre 0 dentro de una expresión con enteros."],
 arreglo:["Usá Long en vez de Integer, y Double para decimales."], fn:["tipos"]},

{id:"e9", tipo:"run", num:9, msg:"Run-time error '9': Subscript out of range", es:"Lo que pediste no existe en la lista.",
 causas:["El nombre de la hoja no coincide exacto: Worksheets(\"A1 Parqueo \") con un espacio de más, o sin tilde.","Pediste una posición fuera del arreglo: datos(0, 1) o datos(n + 1, 1).","Split no tiene esa parte: Split(x, \"-\")(2) cuando solo hay dos."],
 arreglo:["Copiá el nombre de la hoja desde la pestaña (doble clic sobre la pestaña → Ctrl+C).","Los arreglos leídos de la hoja empiezan en 1; Split empieza en 0.","Revisá el adaptador: la hoja debe estar escrita exacta."], fn:["worksheets","split","ubound"]},

{id:"e11", tipo:"run", num:11, msg:"Run-time error '11': Division by zero", es:"Dividiste entre 0.",
 causas:["El total o la meta es 0.","Una celda vacía usada como divisor."],
 arreglo:["Revisá antes: If total <> 0 Then pct = parte / total"]},

{id:"e13", tipo:"run", num:13, msg:"Run-time error '13': Type mismatch", es:"Mezclaste tipos que no se pueden combinar.",
 causas:["Hiciste cuentas con un texto: \"Revisar\" * 2, o una celda con \"N/A\".","Una celda tiene un error de Excel (#N/A, #VALUE!) y la usás en una cuenta.","Asignaste texto a una variable Long o Date.","Uniste texto con + en vez de &."],
 arreglo:["Validá antes con IsNumeric (y IsEmpty) o IsDate.","Usá & para unir textos.","Para saber qué valor llegó: en Immediate escribí ? ws.Cells(fila, 5).Value"], code:"If IsNumeric(v) And Not IsEmpty(v) Then x = v * 2", fn:["isnumeric","isempty","&"]},

{id:"e91", tipo:"run", num:91, msg:"Run-time error '91': Object variable or With block variable not set", es:"Estás usando un objeto que está vacío.",
 causas:["Olvidaste el Set: ws = Worksheets(\"…\") en vez de Set ws = …","Find no encontró nada (devolvió Nothing) y usaste el resultado.","Usaste una Collection o Dictionary sin crearla (Set col = New Collection)."],
 arreglo:["Hojas, rangos, Collection y Dictionary se asignan con Set.","Después de Find, revisá If c Is Nothing antes de usar c."], code:"Set ws = ThisWorkbook.Worksheets(\"A1 Parqueo\")", fn:["set",".find","nothing"]},

{id:"e424", tipo:"run", num:424, msg:"Run-time error '424': Object required", es:"VBA esperaba un objeto y recibió otra cosa.",
 causas:["Pusiste Set con un valor: Set n = 5 o Set x = ws.Range(\"A1\").Value.","Escribiste mal el nombre del objeto (wss.Cells)."],
 arreglo:["Set solo con objetos. Para valores, sin Set.","Revisá el nombre de la variable del objeto."], fn:["set"]},

{id:"e438", tipo:"run", num:438, msg:"Run-time error '438': Object doesn't support this property or method", es:"Ese objeto no tiene esa propiedad o método.",
 causas:["Usaste .Value sobre una hoja (ws.Value).","Usaste una propiedad de rango sobre un Dictionary, o al revés."],
 arreglo:["Revisá qué es la variable: hoja, rango, libro o diccionario."]},

{id:"e457", tipo:"run", num:457, msg:"Run-time error '457': This key is already associated with an element of this collection", es:"Agregaste una clave que ya existe.",
 causas:["En una Collection con clave: es el repetido (normal si usás On Error Resume Next).","En un Dictionary: hiciste .Add con una clave existente."],
 arreglo:["Dictionary: revisá con .Exists antes de .Add, o asigná directo d(k) = d(k) + x.","Collection: encerrá el bucle entre On Error Resume Next y On Error GoTo 0."], fn:[".exists","clave"]},

{id:"e1004", tipo:"run", num:1004, msg:"Run-time error '1004': Application-defined or object-defined error", es:"Excel no pudo hacer lo que le pediste con la hoja. Es el error más común.",
 causas:["La fórmula de .Formula tiene ; o nombres en español (SI, SUMA). Va en inglés y con comas.","Una celda con fila o columna 0: Cells(0, 1) o Range(\"A0\").","Le pusiste a una hoja un nombre que ya existe, con más de 31 caracteres o con : \\ / ? * [ ].","WorksheetFunction.VLookup o .Match no encontró el valor.","La hoja está protegida."],
 arreglo:["En .Formula: =IF(E2=\"\"Sí\"\",25,0) con comas. O usá .FormulaLocal con tu separador.","Borrá la hoja vieja antes de crear la nueva con el mismo nombre.","Para búsquedas que pueden fallar: Find + Is Nothing."], code:"ws.Range(\"G2:G\" & ultima).Formula = \"=IF(E2=\"\"Sí\"\",25,0)\"", fn:[".formula",".formulalocal",".name",".find"]},

{id:"e1004wsf", tipo:"run", num:1004, msg:"Run-time error '1004': Unable to get the VLookup (Match, Max…) property of the WorksheetFunction class", es:"La función de Excel dio error (no encontró el valor o recibió texto).",
 causas:["VLookup o Match no encontró lo que buscaba.","El rango tiene errores o texto donde se esperan números."],
 arreglo:["Usá Application.VLookup (sin WorksheetFunction) y revisá IsError(resultado).","O buscá con Find."], fn:[".vlookup",".match","iserror"]},

{id:"e53", tipo:"run", num:53, msg:"Run-time error '53' / '1004': File not found / … could not be found", es:"El archivo que querés abrir no está en esa ruta.",
 causas:["El nombre del archivo está mal o no está en la misma carpeta.","Falta la \\ entre ThisWorkbook.Path y el nombre."],
 arreglo:["Revisá antes con If Dir(ruta) = \"\" Then (no existe).","ThisWorkbook.Path & \"\\\" & archivo"], fn:["dir",".path"]},

{id:"e429", tipo:"run", num:429, msg:"Run-time error '429': ActiveX component can't create object", es:"No se pudo crear el objeto (normalmente el Dictionary).",
 causas:["Estás en Excel para Mac: Scripting.Dictionary no existe ahí."],
 arreglo:["En el examen (Windows) no pasa. En Mac, usá una Collection o dos arreglos paralelos."], fn:["createobject"]},

{id:"e10", tipo:"run", num:10, msg:"Run-time error '10': This array is fixed or temporarily locked", es:"Intentaste hacer ReDim de un arreglo que no se puede redimensionar.",
 causas:["Lo declaraste con tamaño fijo: Dim lista(1 To 10). Para usar ReDim, se declara con paréntesis vacíos: Dim lista()."],
 arreglo:["Declaralo como Dim lista() As String y después ReDim."], fn:["redim","preserve"]},

/* ---------- Sin mensaje: el resultado sale mal ---------- */
{id:"m-cero", tipo:"mudo", msg:"El resultado sale 0 o vacío", es:"La macro corre pero no calcula nada.",
 causas:["No tenés Option Explicit y una variable está mal escrita (crea una nueva vacía).","Leés la columna equivocada.","El If nunca se cumple: comparás \"Si\" contra \"Sí\", o un número contra un texto."],
 arreglo:["Poné Option Explicit y corré: te marca el nombre mal escrito.","Con F8 avanzá línea por línea y pasá el mouse sobre las variables."], fn:["option"]},

{id:"m-uno", tipo:"mudo", msg:"El conteo da uno de más (31 en vez de 30)", es:"Contaste filas en vez de registros.",
 causas:["Imprimiste ultima, que es el número de fila, no la cantidad."],
 arreglo:["Cantidad = ultima - 1 (si hay una fila de encabezado). O usá un contador."]},

{id:"m-corta", tipo:"mudo", msg:"La macro no llega hasta la última fila", es:"La última fila salió corta.",
 causas:["La columna que usaste para .End(xlUp) tiene vacíos al final.","Hay una fila Total de una corrida anterior que se suma como dato."],
 arreglo:["Usá la columna del código o ticket, que está llena en todas las filas.","Borrá la fila Total antes de recalcular."], fn:[".end"]},

{id:"m-dup", tipo:"mudo", msg:"Al correr dos veces se duplica algo", es:"Quedan filas, hojas o totales repetidos.",
 causas:["La macro agrega una fila Total o un registro sin borrar el anterior.","Crea una hoja sin borrar la de antes (y si no da error, es porque tiene otro nombre)."],
 arreglo:["Borrá resultados viejos al inicio (ClearContents).","Para hojas: borrar con DisplayAlerts = False y crear de nuevo."], fn:[".clearcontents",".delete"]},

{id:"m-color", tipo:"mudo", msg:"Quedan colores de la corrida anterior", es:"Celdas que ya están bien siguen pintadas.",
 causas:["Solo pintás las malas, nunca quitás el relleno a las buenas."],
 arreglo:["En el Else, Interior.ColorIndex = xlNone."], fn:[".colorindex"]},

{id:"m-redondeo", tipo:"mudo", msg:"Un centavo de diferencia", es:"El redondeo no coincide con el esperado.",
 causas:["Round de VBA usa redondeo bancario: 2.125 → 2.12.","Redondeaste en otro momento (al final en vez de cada línea, o al revés)."],
 arreglo:["Usá Application.WorksheetFunction.Round.","Redondeá exactamente donde lo dice el enunciado."], fn:["round",".round"]},

{id:"m-fecha", tipo:"mudo", msg:"La fecha se ve como un número (46300)", es:"El valor está bien, falta el formato.",
 causas:["Escribiste una fecha en una celda sin formato de fecha."],
 arreglo:["ws.Cells(fila, 10).NumberFormat = \"dd/mm/yyyy\""], fn:[".numberformat"]},

{id:"m-name", tipo:"mudo", msg:"La celda muestra #NAME? con tu función", es:"Excel no encuentra tu función.",
 causas:["La función está en el módulo de una hoja o en ThisWorkbook, no en un módulo normal.","Es Private.","El libro no es .xlsm o las macros están deshabilitadas."],
 arreglo:["Movela a un módulo de Insert → Module y declarala Public.","Después de cambiar la función: Ctrl+Alt+F9 para recalcular."], fn:["function","public"]},

{id:"m-value", tipo:"mudo", msg:"La celda muestra #VALUE! con tu función", es:"Tu función recibió algo que no puede procesar o dio error por dentro.",
 causas:["Le llega texto donde espera un número (una celda vacía o \"N/A\").","Dentro de la función ocurre un error."],
 arreglo:["Probá la función desde una macro con Debug.Print para ver el error con su número."]},

{id:"m-altf8", tipo:"mudo", msg:"La macro no aparece en Alt+F8", es:"Excel no la muestra como macro ejecutable.",
 causas:["Es Private (eso está bien si el enunciado lo pide).","Tiene parámetros.","Es una Function, no un Sub."],
 arreglo:["Las macros que el usuario corre: Sub público y sin parámetros."], fn:["private","sub"]},

{id:"m-immediate", tipo:"mudo", msg:"No veo lo que imprime Debug.Print", es:"La ventana Immediate está cerrada.",
 causas:["No está abierta."],
 arreglo:["En el editor: Ctrl+G (o View → Immediate Window)."], fn:["debug"]},

{id:"m-tildes", tipo:"mudo", msg:"Las tildes salen raras (SÃ­) al importar un .bas", es:"El archivo se guardó con otra codificación.",
 causas:["El .bas se guardó en UTF-8; el editor de VBA lee Windows-1252."],
 arreglo:["Usá el botón Descargar .bas del kit, que ya lo guarda bien.","O copiá y pegá el código en vez de importar."]},

{id:"m-cuelga", tipo:"mudo", msg:"Excel se queda pegado (bucle infinito)", es:"Un bucle nunca termina.",
 causas:["Un Do While cuya condición nunca cambia.","Un For que modifica su propio contador."],
 arreglo:["Ctrl+Pausa (Ctrl+Break) o Esc para detener, y luego Reset (cuadro azul).","Revisá que dentro del Do algo cambie la condición."], fn:["do"]},

{id:"m-eventos", tipo:"mudo", msg:"Un evento dejó de funcionar o los avisos no aparecen", es:"Quedó algo apagado de una corrida que se detuvo.",
 causas:["La macro se detuvo con EnableEvents o DisplayAlerts en False."],
 arreglo:["En Immediate: Application.EnableEvents = True y Application.DisplayAlerts = True"], fn:[".enableevents",".displayalerts"]},

{id:"m-pausa", tipo:"mudo", msg:"La macro queda en pausa con una línea amarilla", es:"VBA está en modo depuración después de un error.",
 causas:["Hubo un error y presionaste Debug, o hay un Stop o un punto de interrupción."],
 arreglo:["Arreglá la línea y presioná F5 para seguir, o Reset (cuadro azul) para terminar.","Mientras esté en pausa, no podés correr otra macro."], fn:["stop"]}
];

window.DEPURAR = [
 ["F5", "Correr la macro donde está el cursor (o seguir después de una pausa)."],
 ["F8", "Ejecutar línea por línea. Pasá el mouse sobre una variable para ver su valor."],
 ["F9", "Poner o quitar un punto de interrupción en la línea (la macro se pausa ahí)."],
 ["Ctrl+G", "Abrir la ventana Immediate. Ahí: ? variable muestra su valor; una instrucción se ejecuta al dar Enter."],
 ["Ctrl+Pausa / Esc", "Detener una macro que no termina."],
 ["Reset (cuadro azul)", "Terminar la pausa de depuración para poder correr de nuevo."],
 ["Debug → Compile VBAProject", "Revisa TODO el código sin correrlo y te lleva al primer error de compilación. Hacelo antes de entregar."],
 ["View → Locals Window", "Muestra todas las variables y sus valores mientras avanzás con F8."],
 ["Alt+F8 (en Excel)", "Lista de macros que se pueden correr. Así revisa el profesor los nombres."]
];

window.CHECKLIST = [
 {t:"Antes del examen", items:[
  "Corrí las 4 plantillas en mi Excel y marqué \"Para comprobar\" en cada problema.",
  "Tengo Kit_VBA.html en la computadora, en un USB y en Drive o GitHub Pages.",
  "Activé Tools → Options → Require Variable Declaration en el editor.",
  "Practiqué pegar un enunciado en De enunciado a código y llenar el adaptador.",
  "Sé dónde están F5, F8, Ctrl+G y Debug → Compile VBAProject."]},
 {t:"Al empezar", items:[
  "Guardé el libro de datos como Excel Macro-Enabled Workbook (.xlsm).",
  "Abrí el editor (Alt+F11) e inserté un módulo (Insert → Module) por problema.",
  "Cada módulo empieza con Option Explicit.",
  "Leí el enunciado completo antes de escribir código.",
  "Lo pegué en el analizador: anoté macros, textos exactos y columnas.",
  "Abrí la plantilla más parecida y llené el adaptador con los datos del examen."]},
 {t:"Por cada macro", items:[
  "El nombre del Sub o Function es idéntico al del enunciado.",
  "El nombre de la hoja es idéntico al de la pestaña (espacios y tildes).",
  "Las columnas de entrada y de resultado son las que pide.",
  "Los encabezados y textos (\"Revisar\", \"Al día\", \"Total\") están escritos igual que en el enunciado.",
  "La última fila se calcula sola; nada está fijo en 31 o 121.",
  "Los valores salen iguales a los de \"Para comprobar\" del enunciado.",
  "Los parámetros (tasas, fechas de corte) se leen de la hoja si el enunciado lo pide."]},
 {t:"Antes de entregar", items:[
  "Debug → Compile VBAProject no marca ningún error.",
  "Corrí cada macro dos veces: no se duplican hojas, filas ni totales, y no quedan colores viejos.",
  "En Alt+F8 aparecen las macros públicas y no aparecen las privadas.",
  "No dejé Stop, puntos de interrupción ni macros de prueba que no pidan.",
  "Si usé DisplayAlerts o EnableEvents, quedaron otra vez en True.",
  "Guardé como .xlsm (Ctrl+S) con el nombre que pide el profesor."]},
 {t:"Manejo del tiempo", items:[
  "Primero resolví lo que más se parece a mis plantillas; lo nuevo, al final.",
  "Si una parte me trabó más de 10 minutos, la dejé y seguí con otra.",
  "Guardé cada vez que una macro funcionó."]}
];
