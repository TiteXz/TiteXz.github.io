# TiteXz.github.io
# Revolut → EUR (web)

Versión web **estática** de la app de escritorio: convierte el informe fiscal de Revolut a
euros con los tipos oficiales del BCE. Es HTML + JavaScript puro (sin servidor, sin PHP), así
que **se puede alojar gratis en GitHub Pages**.

Todo ocurre en el navegador del usuario:
- Lee `.csv`, `.xlsx` y `.ods` con [SheetJS](https://sheetjs.com) (cargado por CDN).
- Pide los tipos EUR/USD a la **API SDMX del BCE** (`data-api.ecb.europa.eu`, que permite CORS),
  solo el rango de fechas del informe, con la regla de día hábil anterior.
- Por cada operación: `Cost basis EUR = Cost basis / FX(compra)`,
  `Gross proceeds EUR = Gross proceeds / FX(venta)` y `PnL EUR = proceeds − cost`
  (calculado sin redondear y redondeado al final, igual que el backend de escritorio).
- Genera y descarga el `.xlsx` en euros, con los tipos aplicados y los totales (ventas y
  dividendos) alineados en la última columna. Los dividendos (ya en EUR) se copian tal cual.
- **El archivo no se sube a ningún sitio**: se procesa localmente.

## Probar en local

No necesita compilarse. Como usa `fetch` y un CDN, ábrelo con un servidor local (no con
`file://`):

```bash
cd HaciendaWeb
python3 -m http.server 8000
# abre http://localhost:8000
```

## Publicar en GitHub Pages (github.io)

**Opción A — sitio de usuario** (`https://TU_USUARIO.github.io/`):
1. Crea un repo llamado exactamente `TU_USUARIO.github.io`.
2. Sube el **contenido de esta carpeta** (index.html, styles.css, app.js, .nojekyll) a la raíz
   de la rama `main`.
3. En *Settings → Pages*, fuente: `Deploy from a branch`, rama `main`, carpeta `/ (root)`.
4. En 1–2 min estará en `https://TU_USUARIO.github.io/`.

**Opción B — dentro de un repo cualquiera** (`https://TU_USUARIO.github.io/REPO/`):
1. Sube esta carpeta a un repo (puede ser el mismo del proyecto de escritorio).
2. *Settings → Pages* → rama `main`, carpeta `/ (root)` **o** `/docs` si la pones en `docs/`.
3. Quedará en `https://TU_USUARIO.github.io/REPO/`.

> Requiere internet en el navegador del usuario (para el CDN de SheetJS y la API del BCE).
> Ambos funcionan por HTTPS, compatibles con GitHub Pages.

## Archivos

```
index.html    estructura y carga de SheetJS
styles.css    tema oscuro (negros/grises + verde dinero), redondeado y minimal
app.js        lógica: leer → tipos del BCE → convertir → descargar xlsx
.nojekyll     evita el procesado Jekyll de GitHub Pages
```
