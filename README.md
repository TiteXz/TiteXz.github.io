# Revolut → EUR

Conversor web que transforma el **informe fiscal de Revolut** (acciones) a **euros**, usando
el tipo de cambio **oficial del Banco Central Europeo** correspondiente a cada fecha.

Es una página **100 % estática** (HTML + CSS + JavaScript, sin servidor ni base de datos), por
lo que puede alojarse gratis en **GitHub Pages**. Todo el procesamiento ocurre en el navegador:
**tu archivo nunca se sube a ningún sitio**.

---

## ¿Para qué sirve?

Revolut entrega el informe de compraventa de acciones en **dólares (USD)**, pero la declaración
de la renta en España debe presentarse en **euros**, y cada operación tiene que convertirse con
el tipo de cambio del día correspondiente. Hacerlo a mano es tedioso y propenso a errores.

Esta herramienta automatiza ese trabajo:

1. Cargas el archivo que te da Revolut (`.csv`, `.xlsx` o `.ods`).
2. La página descarga del BCE los tipos EUR/USD de las fechas que necesita.
3. Convierte cada compra y cada venta a euros y calcula el resultado (PnL).
4. Te descarga un `.xlsx` nuevo, ya en euros, con los totales listos.

---

## Cómo funciona (por dentro)

Todo el flujo vive en [`app.js`](app.js) y se ejecuta en el navegador del usuario.

### 1. Lectura del archivo
- Usa **[SheetJS](https://sheetjs.com)** (cargado por CDN) para leer `.csv`, `.xlsx` y `.ods`.
- Detecta automáticamente la **fila de cabeceras** por sus columnas (`date acquired`,
  `date sold`, `symbol`…), sin depender del título `Income from Sells`. Opcionalmente localiza
  el bloque **`Other income & fees`** (dividendos y comisiones, que ya vienen en EUR).
- Corrige texto mal codificado (*mojibake*, p. ej. `€` guardado como `â‚¬`) y normaliza fechas
  en cualquier formato (ISO, `dd/mm/aaaa`, número de serie de Excel…).
- Valida que estén todas las columnas requeridas; si falta alguna, avisa.

### 2. Tipos de cambio del BCE
- Pide los tipos a la **API SDMX del BCE** (`data-api.ecb.europa.eu`), que permite CORS y por
  tanto funciona sin servidor intermedio.
- Solo descarga el **rango de fechas del informe** (con un margen previo de 15 días).
- Aplica la **regla del día hábil anterior**: si una fecha cae en fin de semana o festivo (sin
  cotización), retrocede hasta encontrar el último día con tipo publicado (hasta 10 días atrás).

### 3. Conversión
Por cada operación, dividiendo entre el tipo USD/EUR de la fecha correspondiente:

```
Cost basis EUR     = Cost basis     / FX(fecha de compra)
Gross proceeds EUR = Gross proceeds / FX(fecha de venta)
PnL EUR            = Gross proceeds EUR − Cost basis EUR
```

> El PnL se calcula con los importes **sin redondear** y se redondea solo al final (2 decimales),
> igual que el backend de la app de escritorio, para evitar descuadres por redondeo.
>
> Las operaciones que ya están en **EUR** se copian tal cual, sin tocar el tipo de cambio.

### 4. Generación del resultado
- Construye un `.xlsx` con las mismas columnas del original más **`FX compra`** y **`FX venta`**
  (el tipo aplicado a cada operación, para que sea auditable).
- Añade el **TOTAL PnL EUR** de las ventas y el **TOTAL Net Amount EUR** de los dividendos,
  alineados en la última columna.
- Descarga el archivo como `NOMBRE_ORIGINAL_EUR.xlsx`.

---

## Probar en local

No necesita compilarse. Como usa `fetch` y un CDN, **no funciona abriéndolo con `file://`**;
ábrelo con un servidor local:

```bash
python3 -m http.server 8000
# abre http://localhost:8000
```

---

## Publicar en GitHub Pages

**Opción A — sitio de usuario** (`https://TU_USUARIO.github.io/`):
1. Crea un repo llamado exactamente `TU_USUARIO.github.io`.
2. Sube el contenido de esta carpeta (`index.html`, `styles.css`, `app.js`) a la raíz de `main`.
3. En *Settings → Pages*: fuente `Deploy from a branch`, rama `main`, carpeta `/ (root)`.
4. En 1–2 min estará publicado.

**Opción B — dentro de un repo cualquiera** (`https://TU_USUARIO.github.io/REPO/`):
1. Sube esta carpeta a un repo.
2. *Settings → Pages* → rama `main`, carpeta `/ (root)` (o `/docs` si la pones en `docs/`).

> Requiere conexión a internet en el navegador (para el CDN de SheetJS y la API del BCE), ambos
> por HTTPS y compatibles con GitHub Pages.

---

## Estructura del proyecto

```
index.html    Interfaz: zona de arrastre, botón y carga de SheetJS
styles.css    Tema oscuro (negros/grises + verde), minimalista y responsive
app.js        Lógica: leer → tipos del BCE → convertir → descargar .xlsx
```

---

## Privacidad

El archivo se procesa **íntegramente en tu navegador**. Lo único que sale a la red son las
peticiones a la API pública del BCE para obtener los tipos de cambio (solo fechas, ningún dato
de tus operaciones). Tu informe nunca se envía a ningún servidor.
