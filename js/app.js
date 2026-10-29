/* Interfaz del kit. Datos: window.PROBLEMAS y window.FN. Motor: window.Motor. */
(function () {
"use strict";

var P = window.PROBLEMAS || [], FN = window.FN || {}, M = window.Motor;
var FR = window.FRASES || [], CATS = window.CATEGORIAS || [], EJ = window.EJEMPLOS || [];
var ERR = window.ERRORES || [], DEP = window.DEPURAR || [], CHK = window.CHECKLIST || [];
var $ = function (s, r) { return (r || document).querySelector(s); };
var $$ = function (s, r) { return Array.prototype.slice.call((r || document).querySelectorAll(s)); };
var esc = M.esc;

/* ---------- Guardado en el navegador ---------- */
var store = {
  get: function (k, d) { try { var v = localStorage.getItem("kitvba-" + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
  set: function (k, v) { try { localStorage.setItem("kitvba-" + k, JSON.stringify(v)); } catch (e) {} }
};

var estado = {
  pid: null,
  vista: "problemas",
  texto: store.get("texto", ""),
  nivel: store.get("nivel", 1),
  valores: store.get("valores", {}),
  checks: store.get("checks", {}),
  notas: store.get("notas", {}),
  lista: store.get("checklist", {})
};

function problema() { return P.filter(function (p) { return p.id === estado.pid; })[0] || P[0]; }
function valores(p) { return estado.valores[p.id] || (estado.valores[p.id] = {}); }

/* Resolver alias del diccionario */
function ficha(clave) {
  var f = FN[clave], vueltas = 0;
  while (f && f.al && vueltas++ < 5) { clave = f.al; f = FN[clave]; }
  return f ? { clave: clave, f: f } : null;
}

/* ---------- Tema ---------- */
function aplicarTema(t) {
  if (t) document.documentElement.setAttribute("data-theme", t);
  else document.documentElement.removeAttribute("data-theme");
}
aplicarTema(store.get("tema", null));
$("#tema").addEventListener("click", function () {
  var actual = document.documentElement.getAttribute("data-theme");
  var oscuro = actual ? actual === "dark" : window.matchMedia("(prefers-color-scheme: dark)").matches;
  var nuevo = oscuro ? "light" : "dark";
  aplicarTema(nuevo); store.set("tema", nuevo);
});

/* Altura real de la cabecera para lo que queda fijo debajo */
function medir() {
  var h = $(".top").getBoundingClientRect().height;
  document.documentElement.style.setProperty("--alto", Math.round(h) + "px");
}
window.addEventListener("resize", medir);

/* ---------- Toast ---------- */
var tToast;
function toast(msg) {
  var t = $("#toast"); t.textContent = msg; t.hidden = false;
  clearTimeout(tToast); tToast = setTimeout(function () { t.hidden = true; }, 1800);
}

/* ---------- Copiar y descargar ---------- */
function copiar(texto, msg) {
  function plan_b() {
    var ta = document.createElement("textarea");
    ta.value = texto; ta.style.position = "fixed"; ta.style.opacity = "0";
    document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); toast(msg || "Copiado"); } catch (e) { toast("No se pudo copiar: seleccioná el código a mano"); }
    document.body.removeChild(ta);
  }
  if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(texto).then(function () { toast(msg || "Copiado"); }, plan_b);
  else plan_b();
}
function descargarBas(p) {
  var bytes = M.bytesAnsi(M.bas(p, valores(p), estado.nivel));
  var url = URL.createObjectURL(new Blob([bytes], { type: "application/octet-stream" }));
  var a = document.createElement("a"); a.href = url; a.download = p.modulo + ".bas";
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  toast("Descargado " + p.modulo + ".bas: importalo con File → Import File");
}

/* ---------- Navegación ---------- */
function pintarNav() {
  if (estado.vista === "enunciado") { pintarNavEnunciado(); return; }
  if (estado.vista === "funciones") { pintarNavFunciones(); return; }
  if (estado.vista === "errores") { pintarNavErrores(); return; }
  var h = "";
  P.forEach(function (p) {
    h += '<button type="button" class="nav-p' + (p.id === estado.pid ? " activo" : "") + '" data-ir="' + p.id + '"><span class="id">' + p.id + "</span>" + esc(p.titulo) + "</button>";
    if (p.id === estado.pid) {
      p.macros.forEach(function (m) { h += '<button type="button" class="nav-m" data-macro="' + m.n + '">' + m.n + "</button>"; });
      h += '<button type="button" class="nav-m" data-ancla="comprobar">Para comprobar</button>';
      h += '<button type="button" class="nav-m" data-ancla="variantes">Variantes</button>';
      h += '<div class="nav-sep"></div>';
    }
  });
  h += '<div class="nav-ayuda"><p>Clic en una palabra subrayada del código: qué hace y qué va en cada parámetro.</p><p>Clic en una línea amarilla o roja: de qué depende o cuál es la trampa.</p></div>';
  $("#nav").innerHTML = h;
}

/* ---------- Vista de un problema ---------- */
function pintarProblema() {
  var p = problema(), v = valores(p);
  $("#cajaNombre").textContent = p.id;
  var h = "";

  h += '<header class="cabeza"><h1><span class="id">' + p.id + "</span>" + esc(p.titulo) + "</h1><p>" + esc(p.resumen) + "</p>";
  if (p.nota) h += '<p class="nota">' + esc(p.nota) + "</p>";
  h += "</header>";

  /* Qué pide */
  h += '<section class="bloque"><h2>Lo que pide el enunciado <small>Nombres exactos: se revisan en Alt+F8</small></h2><ul class="pide">';
  p.pide.forEach(function (x) {
    h += '<li><button type="button" class="mac" data-macro="' + x[0] + '" title="Ir al código">' + x[0] + "</button><span>" + esc(x[1]) + "</span></li>";
  });
  h += "</ul></section>";

  /* Adaptador */
  h += '<section class="bloque"><h2>Adaptador <small>Lo que cambie el examen se escribe aquí; el código se reescribe solo</small></h2>';
  h += '<div class="adapt"><div class="adapt-cab"><p>Columnas en letra (C, AA); el número se calcula solo. Decimales con punto o coma.</p>';
  h += '<button type="button" class="btn" id="restaurar">Restaurar valores de la ayudantía</button></div><div class="grilla">';
  var grupos = [];
  p.params.forEach(function (x) { if (grupos.indexOf(x.g) < 0) grupos.push(x.g); });
  grupos.forEach(function (g) {
    h += '<div class="grupo"><h3>' + esc(g) + "</h3>";
    p.params.filter(function (x) { return x.g === g; }).forEach(function (x) {
      var val = v[x.k] != null ? v[x.k] : x.v;
      var ancha = x.t === "txt" && x.v.length > 6;
      h += '<div class="celda' + (ancha ? " ancha" : "") + (v[x.k] != null && v[x.k] !== x.v ? " cambio" : "") + '"><label for="p-' + x.k + '">' + esc(x.l) + '</label><div class="val">';
      h += '<input id="p-' + x.k + '" data-k="' + x.k + '" value="' + esc(val) + '" spellcheck="false" autocomplete="off">';
      if (x.t === "col") h += '<span class="eq" data-eq="' + x.k + '">= ' + M.colNum(val) + "</span>";
      h += "</div></div>";
    });
    h += "</div>";
  });
  h += "</div></div></section>";

  /* Herramientas + macros */
  h += '<div class="herr"><div class="seg" role="group" aria-label="Nivel de comentarios">';
  ["Sin comentarios", "Cortos", "Completos"].forEach(function (t, i) {
    h += '<button type="button" data-nivel="' + i + '" aria-pressed="' + (estado.nivel === i) + '">' + t + "</button>";
  });
  h += '</div><button type="button" class="btn prim" id="copiarModulo">Copiar módulo completo</button>';
  h += '<button type="button" class="btn" id="bajarBas" title="Archivo para File → Import File en el editor">Descargar .bas</button>';
  h += '<div class="leyenda"><span><i class="lc"></i>Depende del adaptador</span><span><i class="lt"></i>Trampa</span><span><i class="lf"></i>Clic: qué hace</span></div></div>';
  h += '<section class="bloque" id="macros"></section>';

  /* Comprobar */
  var ch = estado.checks[p.id] || [];
  h += '<section class="bloque" id="comprobar"><h2>Para comprobar <small>Marcá lo que ya verificaste corriendo en tu Excel</small></h2><ul class="check">';
  p.comprobar.forEach(function (c, i) {
    h += '<li><label><input type="checkbox" data-check="' + i + '"' + (ch[i] ? " checked" : "") + "><span>" + esc(c) + "</span></label></li>";
  });
  h += "</ul></section>";

  /* Variantes */
  h += '<section class="bloque" id="variantes"><h2>Variantes probables del examen <small>Qué líneas cambiar si el enunciado pide esto</small></h2><div id="listaVar"></div></section>';

  /* Notas */
  h += '<section class="bloque notas"><h2>Mis notas de ' + p.id + ' <small>Se guardan en este navegador</small></h2>';
  h += '<textarea id="notas" placeholder="Lo que te cueste recordar de este problema.">' + esc(estado.notas[p.id] || "") + "</textarea></section>";

  $("#vista").innerHTML = h;
  pintarCodigo();
  pintarVariantes();
}

/* Código de cada macro (se repinta al cambiar parámetros o nivel) */
function htmlCodigo(lineas, p) {
  var mapa = {}; p.params.forEach(function (x) { mapa[x.k] = x; });
  return lineas.map(function (l, i) {
    var cls = "ln", tip = "";
    var dep = (l.cambiar || []).map(function (k) { return mapa[k] ? mapa[k].l : k; });
    if (l.trampa) { cls += " trampa"; tip = "Trampa: " + l.trampa; }
    else if (dep.length) cls += " cambiar";
    if (dep.length) tip += (tip ? " | " : "") + "Depende de: " + dep.join(", ");
    var data = (l.cambiar && l.cambiar.length) || l.trampa ? ' data-linea="' + i + '"' : "";
    return '<div class="' + cls + '"' + (tip ? ' data-tip="' + esc(tip) + '"' : "") + data + '><span class="t">' + (M.resaltar(l.t, FN) || " ") + "</span></div>";
  }).join("");
}

var cacheLineas = {};
function pintarCodigo() {
  var p = problema(), v = valores(p), h = "";
  cacheLineas = {};
  p.macros.forEach(function (m) {
    var r = M.procesar(m.code, p.params, v, estado.nivel);
    cacheLineas[m.n] = r.lineas;
    h += '<article class="macro" id="m-' + m.n + '"><div class="macro-cab"><h3>' + m.n + "</h3><p>" + esc(m.d) + "</p>";
    h += '<button type="button" class="btn" data-copiar="' + m.n + '">Copiar</button></div>';
    h += '<pre class="cod" data-macro-cod="' + m.n + '">' + htmlCodigo(r.lineas, p) + "</pre></article>";
  });
  $("#macros").innerHTML = h;
}

function pintarVariantes() {
  var p = problema(), v = valores(p), abiertas = $$("#listaVar details[open]").map(function (d) { return d.dataset.i; });
  var h = "";
  p.variantes.forEach(function (x, i) {
    var a = M.procesar(x.a, p.params, v, estado.nivel), b = M.procesar(x.b, p.params, v, estado.nivel);
    h += '<details class="var" id="v-' + i + '" data-i="' + i + '"' + (abiertas.indexOf(String(i)) >= 0 ? " open" : "") + "><summary><b>" + esc(x.t) + "</b><span>" + esc(x.d) + "</span></summary>";
    h += '<div class="var-cuerpo"><div class="diff">';
    h += '<figure class="a"><figcaption>Antes</figcaption><pre class="cod">' + htmlCodigo(a.lineas.map(function (l) { return { t: l.t }; }), p) + "</pre></figure>";
    h += '<figure class="b"><figcaption>Después <button type="button" class="mini" data-copiar-var="' + i + '">Copiar</button></figcaption><pre class="cod">' + htmlCodigo(b.lineas.map(function (l) { return { t: l.t }; }), p) + "</pre></figure>";
    h += "</div></div></details>";
  });
  $("#listaVar").innerHTML = h;
}

/* ---------- Panel lateral ---------- */
function panelAbrir() { $("#panel").classList.add("abierto"); }
function panelInicio() {
  $("#panelIn").innerHTML =
    '<div class="pv"><h2>Explicación</h2>' +
    "<p>Hacé clic en cualquier palabra subrayada del código (Trim, InStr, .End, Set…) y aquí aparece qué hace, qué va en cada parámetro, un ejemplo y su equivalente en Python y en Excel.</p>" +
    "<p>Las líneas amarillas dependen del adaptador; las rojas tienen una trampa. Hacé clic en ellas para ver el detalle.</p>" +
    '<div class="atajos"><div><span>Buscar</span><kbd>/</kbd></div><div><span>Cerrar o limpiar</span><kbd>Esc</kbd></div><div><span>Nivel de comentarios</span><span><kbd>1</kbd> <kbd>2</kbd> <kbd>3</kbd></span></div><div><span>Copiar módulo completo</span><kbd>C</kbd></div></div></div>';
}
function cerrarBtn() { return '<button type="button" class="cerrar" data-cerrar aria-label="Cerrar">×</button>'; }


/* El Control 2 llega hasta el ejercicio 12 de la guía */
function fueraDelControl(src) {
  if (!src) return false;
  var refs = [], pref = "";
  String(src).split(",").forEach(function (x) {
    var m = /(Clase|Gu[ií]a|Segmento)?\s*E?(\d+)/i.exec(x.trim());
    if (!m) return;
    if (m[1]) pref = m[1].toLowerCase();
    refs.push({ p: pref, n: Number(m[2]) });
  });
  return refs.length > 0 && refs.every(function (r) { return /^gu/.test(r.p) && r.n >= 13; });
}
function etiquetaFuera(src) { return fueraDelControl(src) ? ' <span class="fuera" title="La guía entra al control solo hasta el ejercicio 12">Fuera del Control 2</span>' : ""; }

/* Dónde aparece cada ficha: en tus plantillas y en las piezas del enunciado */
var usos = null;
function armarUsos() {
  usos = {};
  function anotar(code, ref) {
    (code.match(/\.?[A-Za-z_][A-Za-z0-9_]*/g) || []).forEach(function (w) {
      var low = w.toLowerCase(), r = ficha(low) || (low[0] === "." ? null : null);
      if (!r && low[0] === ".") r = ficha(low.slice(1));
      if (!r) return;
      var lista = usos[r.clave] || (usos[r.clave] = []);
      if (!lista.some(function (x) { return x.t === ref.t; })) lista.push(ref);
    });
  }
  P.forEach(function (p) { p.macros.forEach(function (m) { anotar(m.code, { t: p.id + " " + m.n, ir: { p: p.id, macro: m.n } }); }); });
  FR.forEach(function (x) { anotar(x.code, { t: x.t, ir: { frase: x.id }, pieza: true }); });
}
function bloqueCodigo(txt) {
  return '<pre class="firma">' + String(txt).split("\n").map(function (l) { return M.resaltar(l, FN); }).join("\n") + "</pre>";
}
function htmlFicha(clave, conCerrar) {
  var r = ficha(clave); if (!r) return "";
  var f = r.f, h = '<div class="pf"><div class="pf-cab"><h2>' + esc(f.n || clave) + "</h2>" + (conCerrar ? cerrarBtn() : "") + "</div>";
  if (f.f) h += bloqueCodigo(f.f);
  if (f.q) h += '<p class="que">' + esc(f.q) + "</p>";
  if (f.p && f.p.length) {
    h += '<h3>Qué va en cada parámetro</h3><table class="params">';
    f.p.forEach(function (x) { h += "<tr><td>" + esc(x[0]) + "</td><td>" + esc(x[1]) + "</td></tr>"; });
    h += "</table>";
  }
  if (f.r) h += "<h3>Devuelve</h3><p>" + esc(f.r) + "</p>";
  if (f.ej) h += "<h3>Ejemplo</h3>" + bloqueCodigo(f.ej);
  if (f.py || f.xl) {
    h += '<h3>Equivalentes</h3><dl class="eqv">';
    if (f.py) h += "<dt>Python</dt><dd>" + esc(f.py) + "</dd>";
    if (f.xl) h += "<dt>Excel</dt><dd>" + esc(f.xl) + "</dd>";
    h += "</dl>";
  }
  if (f.tr) h += '<div class="alerta"><b>Trampa</b>' + esc(f.tr) + "</div>";
  if (!usos) armarUsos();
  var u = usos[r.clave] || [];
  var enPl = u.filter(function (x) { return !x.pieza; }), enPz = u.filter(function (x) { return x.pieza; });
  if (enPl.length || enPz.length || f.src) {
    h += '<h3>Dónde aparece</h3><div class="usos">';
    if (f.src) h += '<p class="src">En tus materiales: ' + esc(f.src) + etiquetaFuera(f.src) + "</p>";
    enPl.forEach(function (x, i) { h += '<button type="button" class="chip cod-chip" data-uso="' + r.clave + "|p|" + i + '">' + esc(x.t) + "</button>"; });
    enPz.slice(0, 8).forEach(function (x, i) { h += '<button type="button" class="chip" data-uso="' + r.clave + "|z|" + i + '">' + esc(x.t) + "</button>"; });
    h += "</div>";
  }
  return h + "</div>";
}
function irUso(v) {
  var p = v.split("|"), u = (usos[p[0]] || []).filter(function (x) { return p[1] === "z" ? x.pieza : !x.pieza; })[Number(p[2])];
  if (u) irA(u.ir);
}

function mostrarFn(clave) {
  var r = ficha(clave);
  if (!r) return;
  $("#panelIn").innerHTML = htmlFicha(clave, true);
  $$(".tok.activo").forEach(function (t) { t.classList.remove("activo"); });
  $$('.tok[data-fn="' + clave + '"]').forEach(function (t) { t.classList.add("activo"); });
  panelAbrir();
}

/* =====================================================================
   FASE 3: Diccionario de funciones
   ===================================================================== */
var TEMAS = window.FN_TEMAS || [];
function abrirFunciones() {
  estado.vista = "funciones";
  marcarPestana();
  if (location.hash !== "#funciones") history.replaceState(null, "", "#funciones");
  $("#cajaNombre").textContent = "( )";
  pintarNav(); pintarFunciones(); panelInicio(); window.scrollTo(0, 0);
}
function pintarNavFunciones() {
  var h = '<div class="nav-tit">Temas</div>';
  TEMAS.forEach(function (t, i) {
    h += '<button type="button" class="nav-m nav-c" data-ancla="tema-' + i + '">' + esc(t[0]) + ' <span class="cuenta">' + t[1].length + "</span></button>";
  });
  h += '<div class="nav-ayuda"><p>Cada ficha dice qué hace, qué va en cada parámetro, su equivalente en Python y Excel, la trampa y dónde aparece en tus plantillas.</p><p>Las palabras de color dentro de los ejemplos también se pueden abrir.</p></div>';
  $("#nav").innerHTML = h;
}
function pintarFunciones() {
  var total = TEMAS.reduce(function (a, t) { return a + t[1].length; }, 0);
  var h = '<header class="cabeza"><h1><span class="id">( )</span>Diccionario de funciones</h1>';
  h += "<p>" + total + " fichas de VBA y de las funciones de Excel que se usan desde VBA, por tema. Filtrá por nombre, por lo que hace o por su equivalente en Python.</p></header>";
  h += '<div class="filtro"><input id="filtroFn" type="search" placeholder="Filtrar: fecha, redondeo, split, len(), quitar espacios…" autocomplete="off" spellcheck="false"><button type="button" class="btn" id="abrirTodas">Abrir todas</button></div>';
  h += '<div id="listaFn"></div>';
  $("#vista").innerHTML = h;
  pintarListaFn("");
}
function pintarListaFn(q) {
  var nq = norm(q).trim(), h = "", n = 0;
  TEMAS.forEach(function (t, i) {
    var items = t[1].filter(function (k) {
      var r = ficha(k); if (!r) return false;
      if (!nq) return true;
      var f = r.f, s = norm(k + " " + (f.n || "") + " " + (f.q || "") + " " + (f.py || "") + " " + (f.xl || "") + " " + (f.f || ""));
      return nq.split(/\s+/).every(function (w) { return s.indexOf(w.length > 5 ? w.slice(0, 5) : w) >= 0; });
    });
    if (!items.length) return;
    h += '<h2 class="cat-tit" id="tema-' + i + '">' + esc(t[0]) + "</h2>";
    items.forEach(function (k) {
      var f = ficha(k).f; n++;
      h += '<details class="ficha" id="fn-' + esc(k) + '"' + (nq ? " open" : "") + "><summary><b>" + esc(f.n || k) + "</b><span>" + esc(f.q || "") + "</span></summary>";
      h += '<div class="ficha-cuerpo">' + htmlFicha(k, false) + "</div></details>";
    });
  });
  $("#listaFn").innerHTML = h || '<p class="vacio">Ninguna ficha con “' + esc(q) + '”. Probá con lo que hace (redondear, fecha, buscar) o con su nombre en Python.</p>';
}

function mostrarLinea(macro, i, el) {
  var p = problema(), l = (cacheLineas[macro] || [])[i];
  if (!l) return;
  var mapa = {}; p.params.forEach(function (x) { mapa[x.k] = x; });
  var v = valores(p);
  var h = '<div class="pf"><div class="pf-cab"><h2>Línea ' + (i + 1) + "</h2>" + cerrarBtn() + '</div><pre class="firma">' + esc(l.t.trim()) + "</pre>";
  if (l.cambiar && l.cambiar.length) {
    h += '<div class="aviso"><b>Depende del adaptador</b><ul>';
    l.cambiar.forEach(function (k) {
      var x = mapa[k]; if (!x) return;
      var val = v[k] != null && v[k] !== "" ? v[k] : x.v;
      h += "<li>" + esc(x.l) + ": <code>" + esc(val) + "</code>" + (x.t === "col" ? " (columna " + M.colNum(val) + ")" : "") + "</li>";
    });
    h += '</ul></div><p class="donde">Para cambiarlo, editá esa casilla del adaptador.</p>';
  }
  if (l.trampa) h += '<div class="alerta"><b>Trampa</b>' + esc(l.trampa) + "</div>";
  h += "</div>";
  $("#panelIn").innerHTML = h;
  $$(".ln.sel").forEach(function (x) { x.classList.remove("sel"); });
  el.classList.add("sel");
  panelAbrir();
}

/* ---------- Búsqueda ---------- */
function norm(s) { return String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, ""); }
var indice = [];
function armarIndice() {
  indice = [];
  Object.keys(FN).forEach(function (k) {
    var f = FN[k]; if (f.al) return;
    indice.push({ g: "Funciones y palabras", t: f.n || k, d: f.q || "", s: norm((f.n || "") + " " + k + " " + (f.q || "") + " " + (f.py || "") + " " + (f.xl || "")), ir: { fn: k } });
  });
  ERR.forEach(function (e) {
    indice.push({ g: "Errores", t: (e.num != null ? e.num + " " : "") + e.msg.replace(/^(Run-time error '\d+': |Compile error: )/, ""), d: e.es, s: norm((e.num || "") + " " + e.msg + " " + e.es), ir: { error: e.id } });
  });
  FR.forEach(function (x) {
    indice.push({ g: "Del enunciado", t: x.t, d: x.q, s: norm(x.t + " " + x.fr.join(" ") + " " + x.q), ir: { frase: x.id } });
  });
  P.forEach(function (p) {
    p.macros.forEach(function (m) {
      indice.push({ g: "Macros", t: m.n, d: p.id + " " + p.titulo + ": " + m.d, s: norm(m.n + " " + m.d + " " + (m.kw || "") + " " + p.titulo + " " + p.id), ir: { p: p.id, macro: m.n } });
    });
    p.variantes.forEach(function (x, i) {
      indice.push({ g: "Variantes", t: x.t, d: p.id + ": " + x.d, s: norm(x.t + " " + x.d + " " + p.titulo), ir: { p: p.id, variante: i } });
    });
  });
}
var resActuales = [], resSel = 0;
function buscar(q) {
  var caja = $("#resultados"), nq = norm(q).trim();
  if (!nq) { caja.hidden = true; return; }
  var vacias = ["de","la","el","en","y","a","que","por","con","los","las","del","un","una","al","se","o"];
  var palabras = nq.split(/\s+/).filter(function (w) { return w && vacias.indexOf(w) < 0; });
  if (!palabras.length) palabras = [nq];
  resActuales = indice.map(function (x) {
    var hits = 0, pts = 0, tn = norm(x.t);
    palabras.forEach(function (w) {
      if (x.s.indexOf(w) >= 0) { hits++; pts++; }
      if (tn === w) pts += 10; else if (tn.indexOf(w) === 0) pts += 3;
    });
    return { x: x, hits: hits, pts: pts };
  }).filter(function (r) { return r.hits === palabras.length; })
    .sort(function (a, b) { return b.pts - a.pts; }).slice(0, 16).map(function (r) { return r.x; });
  /* agrupar respetando el orden del mejor resultado de cada grupo */
  var grupos = [];
  resActuales.forEach(function (x) { if (grupos.indexOf(x.g) < 0) grupos.push(x.g); });
  resActuales.sort(function (a, b) { return grupos.indexOf(a.g) - grupos.indexOf(b.g); });
  resSel = 0;
  if (!resActuales.length) {
    caja.innerHTML = '<div class="res-vacio">Nada con “' + esc(q) + '”. Probá con el nombre de la función (Mid, Weekday) o una palabra del enunciado (fin de semana, sin repetir).</div>';
  } else {
    var h = "", g = "";
    resActuales.forEach(function (x, i) {
      if (x.g !== g) { g = x.g; h += '<div class="res-grupo">' + g + "</div>"; }
      h += '<button type="button" class="res' + (i === 0 ? " sel" : "") + '" data-res="' + i + '"><b>' + esc(x.t) + "</b><span>" + esc(x.d) + "</span></button>";
    });
    caja.innerHTML = h;
  }
  caja.hidden = false;
}
function irA(ir) {
  $("#resultados").hidden = true;
  if (ir.fn) { mostrarFn(ir.fn); return; }
  if (ir.error) {
    if (estado.vista !== "errores") abrirErrores();
    var de = $("#e-" + ir.error); if (de) { de.open = true; scrollA("#e-" + ir.error); }
    return;
  }
  if (ir.frase) {
    if (estado.vista !== "enunciado") abrirEnunciado();
    var df = $("#f-" + ir.frase); if (df) { df.open = true; scrollA("#f-" + ir.frase); }
    return;
  }
  if (ir.p && (ir.p !== estado.pid || estado.vista !== "problemas")) abrir(ir.p);
  if (ir.macro) scrollA("#m-" + ir.macro);
  if (ir.variante != null) { var d = $("#v-" + ir.variante); if (d) { d.open = true; scrollA("#v-" + ir.variante); } }
}
function scrollA(sel) { var el = $(sel); if (el) el.scrollIntoView({ behavior: "smooth", block: "start" }); }

/* ---------- Abrir un problema ---------- */
function abrir(id) {
  if (id === "enunciado") { abrirEnunciado(); return; }
  if (id === "funciones") { abrirFunciones(); return; }
  if (id === "errores") { abrirErrores(); return; }
  estado.vista = "problemas";
  marcarPestana();
  estado.pid = id;
  if (location.hash !== "#" + id) history.replaceState(null, "", "#" + id);
  pintarNav(); pintarProblema(); panelInicio(); window.scrollTo(0, 0);
}


/* =====================================================================
   FASE 2: De enunciado a código
   ===================================================================== */
function marcarPestana() {
  $$("[data-vista]").forEach(function (b) { b.classList.toggle("activa", b.dataset.vista === estado.vista); });
}
function abrirEnunciado() {
  estado.vista = "enunciado";
  marcarPestana();
  if (location.hash !== "#enunciado") history.replaceState(null, "", "#enunciado");
  $("#cajaNombre").textContent = "fx";
  pintarNav(); pintarEnunciado(); panelInicio(); window.scrollTo(0, 0);
}
function pintarNavEnunciado() {
  var h = '<button type="button" class="nav-p activo" data-ancla="analizador"><span class="id">fx</span>Analizar un enunciado</button>';
  h += '<div class="nav-sep"></div><div class="nav-tit">Catálogo</div>';
  CATS.forEach(function (c) {
    var n = FR.filter(function (x) { return x.cat === c[0]; }).length;
    h += '<button type="button" class="nav-m nav-c" data-ancla="cat-' + c[0] + '">' + esc(c[1]) + ' <span class="cuenta">' + n + "</span></button>";
  });
  h += '<div class="nav-ayuda"><p>Pegá el enunciado del examen: se marcan las frases que reconoce y aparece la pieza de código de cada una.</p><p>Todo funciona con palabras clave dentro del archivo: sin internet y sin IA.</p></div>';
  $("#nav").innerHTML = h;
}

var SIN_CAMEL = ["InputBox", "MsgBox", "WorksheetFunction", "ColorIndex", "NumberFormat", "AutoFit", "DisplayAlerts", "ThisWorkbook",
  "StatusBar", "ScreenUpdating", "ActiveSheet", "ReDim", "UBound", "StrPtr", "InStr", "LCase", "UCase", "VLookup", "CountIf", "SumIf",
  "FormulaLocal", "ElseIf", "EndIf", "ForEach", "SaveAs", "LookAt", "LookIn", "GoTo", "OnError", "SaveChanges", "FileFormat", "EnableEvents", "CountIf", "CountIfs", "SumIf", "SumIfs", "CountA", "RoundUp", "RoundDown", "ProperCase", "TextCompare", "PowerPoint", "WhatsApp", "PayPal", "YouTube"];

function htmlPieza(code) { return htmlCodigo(M.procesar(code, [], {}, estado.nivel).lineas, { params: [] }); }

function pintarEnunciado() {
  var h = '<header class="cabeza"><h1><span class="id">fx</span>De enunciado a código</h1>';
  h += "<p>Cada frase típica de un enunciado tiene una pieza de código que la resuelve. Pegá el enunciado del examen y te digo qué piezas necesitás, en el orden en que se arma una macro.</p></header>";

  h += '<section class="bloque" id="analizador"><h2>Analizar un enunciado <small>Pegalo completo o por partes</small></h2>';
  h += '<div class="analiza"><textarea id="texto" spellcheck="false" placeholder="Pegá aquí el enunciado del examen…">' + esc(estado.texto) + "</textarea>";
  h += '<div class="analiza-pie"><span>Probar con:</span>';
  EJ.forEach(function (e, i) { h += '<button type="button" class="chip" data-ejemplo="' + i + '">' + esc(e.t) + "</button>"; });
  h += '<button type="button" class="chip quitar" id="limpiarTexto">Borrar</button></div></div>';
  h += '<div id="analisis"></div></section>';

  h += '<div class="herr"><div class="seg" role="group" aria-label="Nivel de comentarios">';
  ["Sin comentarios", "Cortos", "Completos"].forEach(function (t, i) {
    h += '<button type="button" data-nivel="' + i + '" aria-pressed="' + (estado.nivel === i) + '">' + t + "</button>";
  });
  h += '</div><div class="leyenda"><span><i class="lt"></i>Trampa</span><span><i class="lf"></i>Clic: qué hace</span></div></div>';

  h += '<section class="bloque" id="catalogo"><h2>Catálogo de frases <small>' + FR.length + " piezas; abrí la que necesités</small></h2><div id=\"listaCat\"></div></section>";
  $("#vista").innerHTML = h;
  pintarAnalisis();
  pintarCatalogo();
}

/* Texto normalizado con el mapa de posiciones al original (para marcar frases) */
function normConMapa(txt) {
  var n = "", mapa = [];
  for (var i = 0; i < txt.length; i++) {
    var c = txt[i].normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
    for (var j = 0; j < c.length; j++) { n += c[j]; mapa.push(i); }
  }
  mapa.push(txt.length);
  return { n: n, mapa: mapa };
}
function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }

function analizar(txt) {
  var nm = normConMapa(txt), hits = [], rangos = [];
  FR.forEach(function (x, orden) {
    var primera = -1, frase = "";
    x.fr.forEach(function (f) {
      var re = new RegExp("(^|[^a-z0-9])(" + escRe(norm(f)) + ")(?=$|[^a-z0-9])", "g"), m;
      while ((m = re.exec(nm.n))) {
        var ini = m.index + m[1].length, fin = ini + m[2].length;
        rangos.push({ a: nm.mapa[ini], b: nm.mapa[fin], id: x.id });
        if (primera < 0 || ini < primera) { primera = ini; frase = txt.slice(nm.mapa[ini], nm.mapa[fin]); }
        if (m[2].length === 0) re.lastIndex++;
      }
    });
    if (primera >= 0) hits.push({ x: x, pos: primera, frase: frase, orden: orden });
  });
  var ordenCat = CATS.map(function (c) { return c[0]; });
  hits.sort(function (a, b) {
    return (fueraDelControl(a.x.src) - fueraDelControl(b.x.src)) || (ordenCat.indexOf(a.x.cat) - ordenCat.indexOf(b.x.cat)) || (a.orden - b.orden);
  });

  var macros = [];
  (txt.match(/\b[A-Z][a-z0-9]+(?:[A-Z][a-z0-9]*)+\b/g) || []).forEach(function (w) {
    if (SIN_CAMEL.indexOf(w) >= 0 || FN[w.toLowerCase()] || macros.indexOf(w) >= 0) return;
    macros.push(w);
  });
  var reNom = /\b(?:Funci[oó]n|funci[oó]n|macro|Sub)\s+([A-Z][A-Za-z0-9]+)/g, mn;
  while ((mn = reNom.exec(txt))) {
    var w2 = mn[1];
    if (SIN_CAMEL.indexOf(w2) < 0 && !FN[w2.toLowerCase()] && macros.indexOf(w2) < 0) macros.push(w2);
  }
  var textos = [];
  (txt.match(/["“]([^"”\n]{1,40})["”]/g) || []).forEach(function (q) {
    var t = q.slice(1, -1).trim();
    if (t && !/^=/.test(t) && !/^[A-Z]{1,3}\d*:[A-Z]{1,3}\d*$/.test(t) && textos.indexOf(t) < 0) textos.push(t);
  });
  var cols = [];
  var reCol = /\b(?:columnas?|En)\s+([A-Z]{1,2})\b(?:\s+(?:a|y)\s+([A-Z]{1,2})\b)?/g, mc;
  while ((mc = reCol.exec(txt))) { var c = mc[1] + (mc[2] ? " a " + mc[2] : ""); if (cols.indexOf(c) < 0) cols.push(c); }

  var votos = {};
  hits.forEach(function (hh) { if (hh.x.ver) { var pid = hh.x.ver.split(":")[0]; votos[pid] = (votos[pid] || 0) + 1; } });
  P.forEach(function (p) { p.macros.forEach(function (m) { if (macros.indexOf(m.n) >= 0) votos[p.id] = (votos[p.id] || 0) + 5; }); });
  var parecido = Object.keys(votos).sort(function (a, b) { return votos[b] - votos[a]; })[0];

  return { hits: hits, rangos: rangos, macros: macros, textos: textos, cols: cols, parecido: parecido };
}

function textoMarcado(txt, rangos) {
  rangos.sort(function (a, b) { return a.a - b.a || b.b - a.b; });
  var h = "", pos = 0;
  rangos.forEach(function (r) {
    if (r.a < pos) return;
    h += esc(txt.slice(pos, r.a)) + '<mark data-ir-frase="' + r.id + '" title="Ver la pieza">' + esc(txt.slice(r.a, r.b)) + "</mark>";
    pos = r.b;
  });
  return h + esc(txt.slice(pos));
}

var ultimoAnalisis = null;
function pintarAnalisis() {
  var caja = $("#analisis"); if (!caja) return;
  var txt = estado.texto || "";
  if (!txt.trim()) {
    caja.innerHTML = '<p class="vacio">Cuando pegués un enunciado, aquí aparecen las macros que pide, los textos que deben ir idénticos, las columnas que menciona y las piezas de código, en orden. Probá con uno de los ejemplos.</p>';
    ultimoAnalisis = null; return;
  }
  var r = analizar(txt); ultimoAnalisis = r;
  var h = '<div class="resumen">';
  h += '<div class="res-bloque"><h3>Macros que pide</h3>' + (r.macros.length ? r.macros.map(function (m) {
    return '<button type="button" class="chip cod-chip" data-esqueleto="' + esc(m) + '" title="Copiar Sub ' + esc(m) + '() … End Sub">' + esc(m) + "</button>";
  }).join("") : '<span class="tenue">No encontré nombres tipo CalcularParqueo.</span>') + "</div>";
  h += '<div class="res-bloque"><h3>Textos que deben ir idénticos</h3>' + (r.textos.length ? r.textos.map(function (t) {
    return '<button type="button" class="chip cod-chip" data-copiar-txt="' + esc(t) + '" title="Copiar">' + esc(t) + "</button>";
  }).join("") : '<span class="tenue">Ninguno entre comillas.</span>') + "</div>";
  h += '<div class="res-bloque"><h3>Columnas que menciona</h3>' + (r.cols.length ? r.cols.map(function (c) { return '<span class="chip quieto">' + esc(c) + "</span>"; }).join("") : '<span class="tenue">Ninguna.</span>') + "</div>";
  if (r.parecido) {
    var pp = P.filter(function (p) { return p.id === r.parecido; })[0];
    h += '<div class="res-bloque"><h3>Plantilla más parecida</h3><button type="button" class="btn" data-ir="' + pp.id + '">' + pp.id + " " + esc(pp.titulo) + "</button></div>";
  }
  h += "</div>";

  h += '<details class="marcado" open><summary>El enunciado con las frases reconocidas <span>' + r.hits.length + " piezas</span></summary><div class=\"marcado-txt\">" + textoMarcado(txt, r.rangos) + "</div></details>";

  if (r.hits.length) {
    h += '<div class="piezas-cab"><h3>Piezas, en el orden en que se arma la macro</h3><button type="button" class="btn prim" id="copiarPiezas">Copiar todas las piezas</button></div>';
    var cat = "";
    r.hits.forEach(function (hh) {
      if (hh.x.cat !== cat) { cat = hh.x.cat; h += '<div class="pieza-cat">' + esc(CATS.filter(function (c) { return c[0] === cat; })[0][1]) + "</div>"; }
      h += tarjetaPieza(hh.x, hh.frase, "a-");
    });
  }
  caja.innerHTML = h;
}

function tarjetaPieza(x, frase, pre) {
  var h = '<article class="pieza" id="' + pre + x.id + '"><div class="pieza-cab"><div><h4>' + esc(x.t) + "</h4>";
  if (frase) h += '<p class="pidio">Lo pide: «' + esc(frase) + "»</p>";
  if (x.src) h += '<p class="pidio otro">En tus materiales: ' + esc(x.src) + etiquetaFuera(x.src) + "</p>";
  h += '</div><div class="pieza-acc"><button type="button" class="btn" data-copiar-pieza="' + x.id + '">Copiar</button>';
  if (x.ver) {
    var v = x.ver.split(":");
    h += '<button type="button" class="btn" data-ver="' + x.ver + '" title="Abrir la macro completa">Ver en ' + v[0] + "</button>";
  }
  h += '</div></div><p class="pieza-q">' + esc(x.q) + "</p>";
  h += '<pre class="cod">' + htmlPieza(x.code) + "</pre>";
  if (x.tr) h += '<p class="pieza-tr"><b>Trampa:</b> ' + esc(x.tr) + "</p>";
  return h + "</article>";
}

function pintarCatalogo() {
  var caja = $("#listaCat"); if (!caja) return;
  var abiertos = $$("#listaCat details[open]").map(function (d) { return d.id; });
  var h = "";
  CATS.forEach(function (c) {
    h += '<h3 class="cat-tit" id="cat-' + c[0] + '">' + esc(c[1]) + "</h3>";
    FR.filter(function (x) { return x.cat === c[0]; }).forEach(function (x) {
      var id = "f-" + x.id;
      h += '<details class="frase" id="' + id + '"' + (abiertos.indexOf(id) >= 0 ? " open" : "") + "><summary><b>" + esc(x.t) + "</b><span>" + x.fr.slice(0, 3).map(function (f) { return "«" + esc(f) + "»"; }).join(" ") + "</span></summary>";
      h += '<div class="frase-cuerpo">' + tarjetaPieza(x, "", "c-") + "</div></details>";
    });
  });
  caja.innerHTML = h;
}

function piezaPorId(id) { return FR.filter(function (x) { return x.id === id; })[0]; }


/* =====================================================================
   FASE 4: Errores y checklist
   ===================================================================== */
var TIPOS_ERR = [["comp", "De compilación", "VBA no te deja correr: la línea sale marcada antes de empezar."],
  ["run", "En ejecución", "La macro arranca y se detiene con un número de error."],
  ["mudo", "Sin mensaje", "No sale ningún error, pero el resultado está mal."]];
function abrirErrores() {
  estado.vista = "errores";
  marcarPestana();
  if (location.hash !== "#errores") history.replaceState(null, "", "#errores");
  $("#cajaNombre").textContent = "#!";
  pintarNav(); pintarErrores(); panelInicio(); window.scrollTo(0, 0);
}
function pintarNavErrores() {
  var h = '<div class="nav-tit">Errores</div>';
  TIPOS_ERR.forEach(function (t) {
    h += '<button type="button" class="nav-m nav-c" data-ancla="err-' + t[0] + '">' + t[1] + ' <span class="cuenta">' + ERR.filter(function (e) { return e.tipo === t[0]; }).length + "</span></button>";
  });
  h += '<button type="button" class="nav-m nav-c" data-ancla="depurar">Cómo depurar</button>';
  h += '<div class="nav-sep"></div><div class="nav-tit">Checklist</div>';
  CHK.forEach(function (g, i) { h += '<button type="button" class="nav-m nav-c" data-ancla="chk-' + i + '">' + esc(g.t) + ' <span class="cuenta" data-avance="' + i + '"></span></button>'; });
  h += '<div class="nav-ayuda"><p>Escribí en el filtro el número del error o una palabra del mensaje, tal como sale en tu Excel en inglés.</p></div>';
  $("#nav").innerHTML = h;
  avanceChecklist();
}
function pintarErrores() {
  var h = '<header class="cabeza"><h1><span class="id">#!</span>Errores y checklist</h1>';
  h += "<p>Qué significa cada error, por qué pasa casi siempre y cómo arreglarlo. Al final, la lista de revisión para el examen.</p></header>";
  h += '<div class="filtro"><input id="filtroErr" type="search" placeholder="Número o palabra del mensaje: 91, 1004, not defined, ambiguous, #NAME…" autocomplete="off" spellcheck="false"></div>';
  h += '<div id="listaErr"></div>';
  h += '<section class="bloque" id="depurar"><h2>Cómo depurar <small>Teclas del editor de VBA</small></h2><table class="teclas">';
  DEP.forEach(function (d) { h += "<tr><td><kbd>" + esc(d[0]) + "</kbd></td><td>" + esc(d[1]) + "</td></tr>"; });
  h += "</table></section>";
  h += '<section class="bloque" id="checklist"><h2>Checklist del examen <small>Se guarda en este navegador</small> <button type="button" class="mini" id="reiniciarChk">Reiniciar</button></h2>';
  CHK.forEach(function (g, i) {
    var marcados = estado.lista[i] || [];
    h += '<h3 class="cat-tit" id="chk-' + i + '">' + esc(g.t) + "</h3><ul class=\"check\">";
    g.items.forEach(function (it, j) {
      h += '<li><label><input type="checkbox" data-chk="' + i + "-" + j + '"' + (marcados[j] ? " checked" : "") + "><span>" + esc(it) + "</span></label></li>";
    });
    h += "</ul>";
  });
  h += "</section>";
  $("#vista").innerHTML = h;
  pintarListaErr("");
}
function htmlError(e, abierto) {
  var h = '<details class="err" id="e-' + e.id + '"' + (abierto ? " open" : "") + "><summary>";
  h += '<span class="num">' + (e.num != null ? e.num : (e.tipo === "comp" ? "C" : "?")) + "</span><b>" + esc(e.msg) + "</b></summary>";
  h += '<div class="err-cuerpo"><p class="que">' + esc(e.es) + "</p>";
  h += '<div class="err-cols"><div><h3>Por qué pasa</h3><ul>' + e.causas.map(function (c) { return "<li>" + esc(c) + "</li>"; }).join("") + "</ul></div>";
  h += '<div><h3>Cómo arreglarlo</h3><ul>' + e.arreglo.map(function (c) { return "<li>" + esc(c) + "</li>"; }).join("") + "</ul></div></div>";
  if (e.code) h += bloqueCodigo(e.code);
  if (e.fn && e.fn.length) {
    h += '<div class="usos">';
    e.fn.forEach(function (k) { var r = ficha(k); if (r) h += '<button type="button" class="chip cod-chip" data-fn-chip="' + k + '">' + esc(r.f.n || k) + "</button>"; });
    h += "</div>";
  }
  return h + "</div></details>";
}
function pintarListaErr(q) {
  var nq = norm(q).trim(), h = "";
  TIPOS_ERR.forEach(function (t) {
    var items = ERR.filter(function (e) {
      if (e.tipo !== t[0]) return false;
      if (!nq) return true;
      var s = norm((e.num || "") + " " + e.msg + " " + e.es + " " + e.causas.join(" "));
      return nq.split(/\s+/).every(function (w) { return s.indexOf(w) >= 0; });
    });
    if (!items.length) return;
    h += '<section class="bloque" id="err-' + t[0] + '"><h2>' + t[1] + " <small>" + t[2] + "</small></h2>";
    items.forEach(function (e) { h += htmlError(e, !!nq); });
    h += "</section>";
  });
  $("#listaErr").innerHTML = h || '<p class="vacio">Ningún error con “' + esc(q) + '”. Probá con el número (91, 1004) o con una palabra del mensaje en inglés.</p>';
}
function avanceChecklist() {
  CHK.forEach(function (g, i) {
    var el = $('[data-avance="' + i + '"]'); if (!el) return;
    var n = (estado.lista[i] || []).filter(Boolean).length;
    el.textContent = n + "/" + g.items.length;
  });
}

/* ---------- Eventos ---------- */
document.addEventListener("click", function (e) {
  var t = e.target, el;
  if ((el = t.closest("[data-vista]"))) { if (el.dataset.vista === "problemas") abrir(estado.pid || P[0].id); else abrir(el.dataset.vista); return; }
  if ((el = t.closest("[data-uso]"))) { irUso(el.dataset.uso); return; }
  if ((el = t.closest("[data-fn-chip]"))) { mostrarFn(el.dataset.fnChip); return; }
  if (t.closest("#reiniciarChk")) { estado.lista = {}; store.set("checklist", {}); $$("[data-chk]").forEach(function (c) { c.checked = false; }); avanceChecklist(); toast("Checklist reiniciada"); return; }
  if (t.closest("#abrirTodas")) { var todas = $$("#listaFn details"), abrirT = todas.some(function (d) { return !d.open; }); todas.forEach(function (d) { d.open = abrirT; }); t.closest("#abrirTodas").textContent = abrirT ? "Cerrar todas" : "Abrir todas"; return; }
  if ((el = t.closest("[data-ir]"))) { abrir(el.dataset.ir); return; }
  if ((el = t.closest("[data-ejemplo]"))) { estado.texto = EJ[Number(el.dataset.ejemplo)].txt; store.set("texto", estado.texto); $("#texto").value = estado.texto; pintarAnalisis(); return; }
  if (t.closest("#limpiarTexto")) { estado.texto = ""; store.set("texto", ""); $("#texto").value = ""; pintarAnalisis(); $("#texto").focus(); return; }
  if ((el = t.closest("[data-ir-frase]"))) { scrollA("#a-" + el.dataset.irFrase); return; }
  if ((el = t.closest("[data-copiar-pieza]"))) { copiar(M.procesar(piezaPorId(el.dataset.copiarPieza).code, [], {}, estado.nivel).texto, "Pieza copiada"); return; }
  if ((el = t.closest("[data-esqueleto]"))) { copiar("Sub " + el.dataset.esqueleto + "()\n    \nEnd Sub", "Sub " + el.dataset.esqueleto + " copiado"); return; }
  if ((el = t.closest("[data-copiar-txt]"))) { copiar(el.dataset.copiarTxt, "Texto copiado"); return; }
  if (t.closest("#copiarPiezas") && ultimoAnalisis) {
    copiar(ultimoAnalisis.hits.map(function (hh) { return "' ---- " + hh.x.t + " ----\n" + M.procesar(hh.x.code, [], {}, estado.nivel).texto; }).join("\n\n"), "Piezas copiadas"); return;
  }
  if ((el = t.closest("[data-ver]"))) { var v = el.dataset.ver.split(":"); abrir(v[0]); setTimeout(function () { scrollA("#m-" + v[1]); }, 30); return; }
  if ((el = t.closest(".tok"))) { mostrarFn(el.dataset.fn); return; }
  if ((el = t.closest("[data-macro]"))) { scrollA("#m-" + el.dataset.macro); return; }
  if ((el = t.closest("[data-ancla]"))) { scrollA("#" + el.dataset.ancla); return; }
  if ((el = t.closest("[data-nivel]"))) { cambiarNivel(Number(el.dataset.nivel)); return; }
  if ((el = t.closest("[data-copiar]"))) {
    var p = problema(), m = p.macros.filter(function (x) { return x.n === el.dataset.copiar; })[0];
    copiar(M.procesar(m.code, p.params, valores(p), estado.nivel).texto, m.n + " copiada"); return;
  }
  if ((el = t.closest("[data-copiar-var]"))) {
    e.preventDefault();
    var pv = problema(), x = pv.variantes[Number(el.dataset.copiarVar)];
    copiar(M.procesar(x.b, pv.params, valores(pv), estado.nivel).texto, "Variante copiada"); return;
  }
  if (t.closest("#copiarModulo")) { var pm = problema(); copiar(M.modulo(pm, valores(pm), estado.nivel), "Módulo " + pm.id + " copiado"); return; }
  if (t.closest("#bajarBas")) { descargarBas(problema()); return; }
  if (t.closest("#restaurar")) {
    estado.valores[problema().id] = {}; store.set("valores", estado.valores);
    pintarProblema(); toast("Valores de la ayudantía restaurados"); return;
  }
  if (t.closest("[data-cerrar]")) { $("#panel").classList.remove("abierto"); panelInicio(); $$(".tok.activo,.ln.sel").forEach(function (x) { x.classList.remove("activo", "sel"); }); return; }
  if ((el = t.closest("[data-res]"))) { irA(resActuales[Number(el.dataset.res)].ir); $("#q").blur(); return; }
  if ((el = t.closest(".ln[data-linea]"))) { var pre = el.closest("[data-macro-cod]"); if (pre) mostrarLinea(pre.dataset.macroCod, Number(el.dataset.linea), el); return; }
  if (!t.closest(".formula")) $("#resultados").hidden = true;
});

function cambiarNivel(n) {
  estado.nivel = n; store.set("nivel", n);
  $$("[data-nivel]").forEach(function (b) { b.setAttribute("aria-pressed", String(Number(b.dataset.nivel) === n)); });
  if (estado.vista === "enunciado") { pintarAnalisis(); pintarCatalogo(); }
  else { pintarCodigo(); pintarVariantes(); }
}

var tParam;
document.addEventListener("input", function (e) {
  var t = e.target;
  if (t.dataset && t.dataset.k) {
    var p = problema(), x = p.params.filter(function (q) { return q.k === t.dataset.k; })[0];
    valores(p)[t.dataset.k] = t.value;
    t.closest(".celda").classList.toggle("cambio", t.value !== x.v);
    var eq = $('[data-eq="' + t.dataset.k + '"]'); if (eq) eq.textContent = "= " + M.colNum(t.value || x.v);
    clearTimeout(tParam);
    tParam = setTimeout(function () { store.set("valores", estado.valores); pintarCodigo(); pintarVariantes(); }, 150);
    return;
  }
  if (t.id === "notas") { estado.notas[problema().id] = t.value; store.set("notas", estado.notas); return; }
  if (t.id === "q") buscar(t.value);
  if (t.id === "filtroFn") pintarListaFn(t.value);
  if (t.id === "filtroErr") pintarListaErr(t.value);
  if (t.id === "texto") {
    estado.texto = t.value;
    clearTimeout(tParam);
    tParam = setTimeout(function () { store.set("texto", estado.texto); pintarAnalisis(); }, 250);
  }
});
document.addEventListener("change", function (e) {
  var t = e.target;
  if (t.dataset && t.dataset.check != null) {
    var id = problema().id, arr = estado.checks[id] || (estado.checks[id] = []);
    arr[Number(t.dataset.check)] = t.checked; store.set("checks", estado.checks);
  }
  if (t.dataset && t.dataset.chk) {
    var pz = t.dataset.chk.split("-"), g = estado.lista[pz[0]] || (estado.lista[pz[0]] = []);
    g[Number(pz[1])] = t.checked; store.set("checklist", estado.lista); avanceChecklist();
  }
});

$("#q").addEventListener("keydown", function (e) {
  var caja = $("#resultados");
  if (e.key === "Escape") { this.value = ""; caja.hidden = true; this.blur(); return; }
  if (caja.hidden || !resActuales.length) return;
  if (e.key === "ArrowDown" || e.key === "ArrowUp") {
    e.preventDefault();
    resSel = (resSel + (e.key === "ArrowDown" ? 1 : -1) + resActuales.length) % resActuales.length;
    $$(".res", caja).forEach(function (b, i) { b.classList.toggle("sel", i === resSel); if (i === resSel) b.scrollIntoView({ block: "nearest" }); });
  }
  if (e.key === "Enter") { e.preventDefault(); irA(resActuales[resSel].ir); this.blur(); }
});
$("#q").addEventListener("focus", function () { if (this.value) buscar(this.value); });

document.addEventListener("keydown", function (e) {
  var en = document.activeElement, escribiendo = en && (en.tagName === "INPUT" || en.tagName === "TEXTAREA");
  if (escribiendo || e.ctrlKey || e.metaKey || e.altKey) return;
  if (e.key === "/") { e.preventDefault(); $("#q").focus(); }
  else if (e.key === "Escape") { $("#panel").classList.remove("abierto"); }
  else if (e.key === "1" || e.key === "2" || e.key === "3") cambiarNivel(Number(e.key) - 1);
  else if (e.key === "c" || e.key === "C") { var p = problema(); copiar(M.modulo(p, valores(p), estado.nivel), "Módulo " + p.id + " copiado"); }
});

/* Tooltip de las líneas marcadas */
var tip = $("#tip");
document.addEventListener("mousemove", function (e) {
  var el = e.target.closest && e.target.closest("[data-tip]");
  if (!el || e.target.closest(".tok")) { tip.hidden = true; return; }
  tip.textContent = el.dataset.tip; tip.hidden = false;
  var x = Math.min(e.clientX + 14, window.innerWidth - tip.offsetWidth - 10);
  var y = e.clientY + 18; if (y + tip.offsetHeight > window.innerHeight - 10) y = e.clientY - tip.offsetHeight - 12;
  tip.style.left = x + "px"; tip.style.top = y + "px";
});
document.addEventListener("mouseleave", function () { tip.hidden = true; });

/* ---------- Inicio ---------- */
armarIndice();
var inicial = (location.hash || "").replace("#", "");
estado.pid = P[0].id;
abrir(inicial === "enunciado" || inicial === "funciones" || inicial === "errores" || P.some(function (p) { return p.id === inicial; }) ? inicial : P[0].id);
medir();
window.addEventListener("hashchange", function () {
  var id = location.hash.replace("#", "");
  if ((id === "enunciado" || id === "funciones" || id === "errores") && estado.vista !== id) abrir(id);
  else if ((id !== estado.pid || estado.vista !== "problemas") && P.some(function (p) { return p.id === id; })) abrir(id);
});
})();
