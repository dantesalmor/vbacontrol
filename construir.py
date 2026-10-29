"""Construye el kit.

Uso (desde la carpeta kit-vba):
    python construir.py

Hace dos cosas:
1. Junta index.html + css + js + data en un solo archivo: Kit_VBA.html
   (el que se usa en el examen y se sube a GitHub Pages o Drive).
2. Si tenés Node.js instalado, genera los .bas de cada problema en vba/
   con los valores de la ayudantía y comentarios cortos.
   Sin Node, usá el botón "Descargar .bas" dentro del kit.
"""
import pathlib
import re
import shutil
import subprocess

RAIZ = pathlib.Path(__file__).parent


def leer(ruta):
    return (RAIZ / ruta).read_text(encoding="utf-8")


def construir_html():
    html = leer("index.html")

    def css(m):
        return "<style>\n" + leer(m.group(1)) + "\n</style>"

    def js(m):
        codigo = leer(m.group(1)).replace("</script", "<\\/script")
        return "<script>\n" + codigo + "\n</script>"

    html = re.sub(r'<link rel="stylesheet" href="([^"]+)">', css, html)
    html = re.sub(r'<script src="([^"]+)"></script>', js, html)
    destino = RAIZ / "Kit_VBA.html"
    destino.write_text(html, encoding="utf-8")
    print("Listo:", destino.name, f"({destino.stat().st_size // 1024} KB)")


GENERAR_BAS = r"""
global.window = {};
require(process.argv[1] + '/data/problemas.js');
const M = require(process.argv[1] + '/js/motor.js');
const fs = require('fs');
for (const p of window.PROBLEMAS) {
  fs.writeFileSync(process.argv[1] + '/vba/' + p.modulo + '.bas', Buffer.from(M.bytesAnsi(M.bas(p, {}, 1))));
  console.log('Listo: vba/' + p.modulo + '.bas');
}
"""


def construir_bas():
    if not shutil.which("node"):
        print("Node.js no está instalado: los .bas se descargan desde el kit.")
        return
    (RAIZ / "vba").mkdir(exist_ok=True)
    subprocess.run(["node", "-e", GENERAR_BAS, str(RAIZ.resolve())], check=True)


if __name__ == "__main__":
    construir_html()
    construir_bas()
