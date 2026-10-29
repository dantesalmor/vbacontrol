# Kit VBA para el Control 2

## Qué archivo abrir

- **Kit_VBA.html**: todo en un solo archivo. Es el que usás en el examen. Doble clic y funciona, con o sin internet.
- **index.html**: la misma herramienta, pero lee sus partes desde las carpetas. Sirve para editar el kit.

## Carpetas

| Carpeta o archivo | Qué tiene |
|---|---|
| `data/problemas.js` | Los 4 problemas: adaptador, plantillas, comprobaciones y variantes |
| `data/enunciado.js` | Las frases típicas de enunciado y su pieza de código (pestaña De enunciado a código) |
| `data/funciones.js` | Las fichas de la pestaña Funciones, que también se abren al hacer clic en el código |
| `data/errores.js` | El índice de errores, las teclas de depuración y la checklist |
| `js/motor.js` | Aplica el adaptador, los niveles de comentario y el resaltado |
| `js/app.js` | La interfaz |
| `css/estilos.css` | Cómo se ve |
| `vba/` | Los módulos `.bas` con los valores de la ayudantía |
| `construir.py` | Junta todo en `Kit_VBA.html` y regenera los `.bas` |

Después de editar algo en `data/`, `js/` o `css/`, corré `python construir.py` para actualizar `Kit_VBA.html`.

## Importar un .bas en Excel

1. Abrí el libro de datos y guardalo como `.xlsm`.
2. Alt + F11 para abrir el editor.
3. File → Import File… y elegí el `.bas`.
4. Aparece un módulo nuevo con todas las macros del problema.

El botón **Descargar .bas** del kit genera el archivo con los valores que tengas en el adaptador y el nivel de comentarios elegido.

## Publicarlo en GitHub Pages

1. Creá un repositorio público en github.com (por ejemplo `kit-vba`).
2. Subí el contenido de esta carpeta (Add file → Upload files).
3. En Settings → Pages, elegí la rama `main` y la carpeta `/ (root)`.
4. En uno o dos minutos queda en `https://TU-USUARIO.github.io/kit-vba/`.

Lo que guardás en el kit (adaptador, casillas marcadas y notas) queda en el navegador donde lo usaste. En otra computadora empieza limpio.
