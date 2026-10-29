/* Motor del kit: no toca la página, solo transforma texto.
   Lo usa app.js en el navegador y construir.py (vía node) para generar los .bas. */
(function (raiz) {
"use strict";

/* "C" -> 3, "AA" -> 27 */
function colNum(letras) {
  var s = String(letras || "").toUpperCase().replace(/[^A-Z]/g, "");
  var n = 0;
  for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
  return n;
}

/* Valor final de un parámetro, ya listo para pegar en el código */
function valorDe(p, crudo) {
  var v = (crudo === undefined || crudo === null || String(crudo).trim() === "") ? p.v : String(crudo);
  if (p.t === "col") { v = v.toUpperCase().replace(/[^A-Z]/g, "") || p.v; return { txt: v, num: String(colNum(v)) }; }
  if (p.t === "num") { v = v.replace(",", ".").replace(/[^0-9.\-]/g, "") || p.v; return { txt: v }; }
  if (p.t === "low") { v = v.trim().toLowerCase(); }
  return { txt: v.replace(/"/g, '""') };
}

/* Posición del primer ' que está fuera de comillas (o -1) */
function inicioComentario(linea) {
  var dentro = false;
  for (var i = 0; i < linea.length; i++) {
    var c = linea[i];
    if (c === '"') dentro = !dentro;
    else if (c === "'" && !dentro) return i;
  }
  return -1;
}

/* Separa una línea de plantilla en código, comentario corto/largo y trampa */
function partir(linea) {
  var r = { codigo: linea, c1: "", c2: "", trampa: "", soloComentario: 0 };
  var t = linea.trim();
  var m = /^'([12]) (.*)$/.exec(t);
  if (m) { r.soloComentario = Number(m[1]); r.codigo = linea.slice(0, linea.indexOf("'")); r.c1 = m[2]; return r; }
  var pos = inicioComentario(linea);
  if (pos < 0) return r;
  var resto = linea.slice(pos);
  if (!/^'[12!] /.test(resto)) return r;          /* comentario normal: se deja tal cual */
  r.codigo = linea.slice(0, pos).replace(/\s+$/, "");
  var trozos = resto.split(/(?:^|\s)'([12!]) /);   /* ["", "1", "texto", "!", "texto"] */
  for (var i = 1; i < trozos.length; i += 2) {
    var tipo = trozos[i], txt = (trozos[i + 1] || "").trim();
    if (tipo === "1") r.c1 = txt; else if (tipo === "2") r.c2 = txt; else r.trampa = txt;
  }
  return r;
}

/* Aplica parámetros y nivel. nivel: 0 sin comentarios, 1 cortos, 2 completos.
   Devuelve { texto, lineas: [{ t, cambiar:[claves], trampa }] } */
function procesar(plantilla, params, valores, nivel) {
  var mapa = {};
  (params || []).forEach(function (p) { mapa[p.k] = p; });
  var salida = [];
  plantilla.split("\n").forEach(function (cruda) {
    var claves = [];
    var linea = cruda.replace(/\{\{([A-Z0-9_]+)(#?)\}\}/g, function (_, k, num) {
      var p = mapa[k];
      if (!p) return "{{" + k + num + "}}";
      if (claves.indexOf(k) < 0) claves.push(k);
      var v = valorDe(p, valores ? valores[k] : undefined);
      return num ? (v.num || v.txt) : v.txt;
    });
    var pz = partir(linea);
    if (pz.soloComentario) {
      if (nivel >= pz.soloComentario) salida.push({ t: pz.codigo + "' " + pz.c1, cambiar: [], trampa: "", comentario: true });
      return;
    }
    var txt = pz.codigo;
    var notas = [];
    if (nivel >= 1 && pz.c1) notas.push(pz.c1);
    if (nivel >= 2 && pz.trampa) notas.push("Trampa: " + pz.trampa);
    if (notas.length) txt += " ' " + notas.join(". ");
    salida.push({ t: txt, cambiar: claves, trampa: pz.trampa });
  });
  /* sin comentarios: quitar líneas en blanco dobles que quedan */
  if (nivel === 0) {
    salida = salida.filter(function (l, i, a) { return !(l.t.trim() === "" && i > 0 && a[i - 1].t.trim() === ""); });
  }
  return { texto: salida.map(function (l) { return l.t; }).join("\n"), lineas: salida };
}

/* Módulo completo listo para importar (.bas) */
function modulo(problema, valores, nivel) {
  var cuerpo = problema.macros.map(function (m) {
    return procesar(m.code, problema.params, valores, nivel).texto;
  }).join("\n\n");
  return "Option Explicit\n\n" + cuerpo + "\n";
}

function bas(problema, valores, nivel) {
  var txt = 'Attribute VB_Name = "' + problema.modulo + '"\n' + modulo(problema, valores, nivel);
  return txt.replace(/\r?\n/g, "\r\n");
}

/* El editor de VBA importa en Windows-1252: los acentos van como un byte */
function bytesAnsi(texto) {
  var out = new Uint8Array(texto.length);
  for (var i = 0; i < texto.length; i++) {
    var c = texto.charCodeAt(i);
    out[i] = c < 256 ? c : 63; /* "?" para lo que no exista */
  }
  return out;
}

/* ---------- Resaltado ---------- */
var CLAVES = ("sub function end if then else elseif for to next each in dim as set exit do loop while until with " +
  "select case is and or not new private public optional byval byref option explicit call true false nothing " +
  "long double string boolean variant date object worksheet range collection me step redim preserve like mod goto on error resume").split(" ");
var ES_CLAVE = {};
CLAVES.forEach(function (k) { ES_CLAVE[k] = 1; });

function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }

/* Devuelve HTML con spans. fn = diccionario de funciones (para hacerlas clicables) */
function resaltar(linea, fn) {
  var html = "", i = 0, n = linea.length;
  var pos = inicioComentario(linea);
  var codigo = pos < 0 ? linea : linea.slice(0, pos);
  var com = pos < 0 ? "" : linea.slice(pos);
  while (i < codigo.length) {
    var c = codigo[i];
    if (c === '"') {
      var j = codigo.indexOf('"', i + 1);
      while (j >= 0 && codigo[j + 1] === '"') j = codigo.indexOf('"', j + 2);
      if (j < 0) j = codigo.length - 1;
      html += '<span class="s">' + esc(codigo.slice(i, j + 1)) + "</span>";
      i = j + 1; continue;
    }
    if (/[A-Za-z_]/.test(c)) {
      var m = /^[A-Za-z_][A-Za-z0-9_]*/.exec(codigo.slice(i))[0];
      var punto = i > 0 && codigo[i - 1] === ".";
      var low = m.toLowerCase();
      var clave = null;
      if (fn) {
        if (punto && fn["." + low]) clave = "." + low;
        else if (!punto && fn[low]) clave = low;
      }
      var cls = (!punto && ES_CLAVE[low]) ? "k" : (clave ? "f" : "");
      var cont = esc(m);
      if (clave) html += '<button type="button" class="tok ' + cls + '" data-fn="' + clave + '">' + cont + "</button>";
      else html += cls ? '<span class="' + cls + '">' + cont + "</span>" : cont;
      i += m.length; continue;
    }
    if (/[0-9]/.test(c) && !(i > 0 && /[A-Za-z_]/.test(codigo[i - 1]))) {
      var mn = /^[0-9]+(\.[0-9]+)?/.exec(codigo.slice(i))[0];
      html += '<span class="n">' + mn + "</span>";
      i += mn.length; continue;
    }
    html += esc(c); i++;
  }
  if (com) html += '<span class="c">' + esc(com) + "</span>";
  return html;
}

var Motor = { colNum: colNum, procesar: procesar, modulo: modulo, bas: bas, bytesAnsi: bytesAnsi, resaltar: resaltar, esc: esc };
if (typeof module !== "undefined" && module.exports) module.exports = Motor;
else raiz.Motor = Motor;
})(this);
