# FILES.md — Estructura del proyecto

```
etiquetas-qr/
├── label_maker.py            # Interfaz gráfica (Tkinter). Punto de entrada.
├── label_generator.py        # Motor de dibujo: etiqueta individual + hoja completa.
├── printer_labels.py         # Salida: impresión en Windows y exportación a PDF.
├── requirements_labels.txt   # Dependencias de Python.
├── FILES.md                  # Este documento.
├── README.md                 # Documentación del repositorio.
└── datos/                    # Datos de usuario y estado de la aplicación.
    ├── Etiquetas.xlsx        # Plantilla de ejemplo: cómo deben ir los datos.
    ├── excel_labels.txt      # Ruta del último Excel abierto.
    ├── labels.txt            # Configuración guardada automáticamente.
    └── icono.png             # (opcional) Icono de la ventana.
```

---

## Módulos de código

### `label_maker.py`
Aplicación de escritorio en Tkinter. Es el fichero que se ejecuta.

- Carga el Excel (por diálogo o el último usado).
- Panel izquierdo: medidas de hoja, disposición de la rejilla, tamaño del QR y una
  tabla por registro con **cantidad**, **texto adicional** y **tamaño de letra**.
- Panel derecho: vista previa (se genera en un hilo aparte para no bloquear la ventana).
- Botones: Cargar Excel, Vista Previa, Guardar PDF, Imprimir.
- Guarda la configuración en `datos/labels.txt` en cada pulsación de tecla, al
  redimensionar la ventana y al cerrarla.

Detecta si corre empaquetado (`sys.frozen`), así que la carpeta `datos/` se busca junto al
`.exe` cuando se distribuye con PyInstaller y junto al `.py` cuando se ejecuta desde código.

### `label_generator.py`
Sin dependencias de interfaz; se puede usar como librería.

- `LabelConfig` — *dataclass* con toda la configuración (mm, columna del QR, textos y
  tamaños por registro).
- `mm_to_px()` — conversión a píxeles a 300 DPI.
- `generar_etiqueta_con_qr()` — dibuja una etiqueta: genera el QR, parte el texto en
  líneas según el ancho disponible y centra el bloque QR + texto. `show_border=True`
  añade un recuadro azul que **solo** se usa en la vista previa.
- `generar_hoja_etiquetas()` — repite cada registro según su cantidad y las coloca en la
  rejilla de la hoja.

### `printer_labels.py`
- `imprimir_hoja_etiquetas()` — envía la hoja a la impresora mediante `win32print` /
  `win32ui`. **Solo Windows.**
- `exportar_hoja_pdf()` — genera el PDF con ReportLab. Multiplataforma.

### `requirements_labels.txt`
`pandas`, `openpyxl`, `Pillow`, `qrcode`, `reportlab` y `pywin32` (este último solo se
instala en Windows gracias al marcador `sys_platform == 'win32'`).

---

## Carpeta `datos/`

La aplicación crea esta carpeta automáticamente si no existe (`os.makedirs(..., exist_ok=True)`).

### `Etiquetas.xlsx`
Plantilla de ejemplo que enseña **cómo deben ir los datos**. Contiene dos hojas:

**Hoja 1 — `Etiquetas`** (es la única que lee el programa):

| Codigo   | Descripcion                        | Lote      | Ubicacion |
|----------|------------------------------------|-----------|-----------|
| ART-0001 | Tornillo hexagonal M8 x 40 zincado | L2026-014 | A-01-03   |
| ART-0002 | Tuerca autoblocante M8 DIN 985     | L2026-014 | A-01-04   |
| ART-0003 | Arandela plana M8 DIN 125          | L2026-021 | A-02-01   |
| ART-0004 | Brida nylon 200 x 4,8 mm negra     | L2026-007 | B-04-02   |
| ART-0005 | Cinta aislante 19 mm x 20 m        | L2026-007 | B-04-05   |

**Hoja 2 — `Instrucciones`**: las reglas de uso, fuera de la hoja de datos para que no
se lean como registros vacíos.

Reglas que hay que respetar:

1. La fila 1 son las cabeceras. Se pueden renombrar y se pueden añadir columnas, pero
   ninguna cabecera puede quedar vacía.
2. Cada fila a partir de la 2 es un registro, es decir, una etiqueta (o varias, según la
   cantidad que se indique en la aplicación).
3. La **primera columna** se usa como identificador del registro en la lista de
   cantidades de la interfaz. Si supera 25 caracteres se muestra recortada.
4. La columna que se codifica en el QR se elige en la aplicación; por defecto toma la
   primera.
5. En «Texto adicional» se pueden insertar valores del Excel entre llaves:
   `{Descripcion} - Lote {Lote}`.
6. Sin filas vacías entre registros, y los datos siempre en la primera hoja del libro.
7. Todo se lee como texto (`dtype=str`), así que los ceros a la izquierda se conservan.

### `excel_labels.txt`
Una sola línea: la ruta absoluta del último Excel cargado. Al arrancar, la aplicación la
lee y reabre ese fichero si sigue existiendo.

```
C:\Etiquetas\datos\Etiquetas.xlsx
```

### `labels.txt`
Configuración persistente, en formato `CLAVE=VALOR`, una por línea. Se reescribe entero
cada vez que se guarda.

```
SHEET_WIDTH=210
SHEET_HEIGHT=297
LABELS_PER_ROW=2
LABELS_PER_COL=4
GAP_MM=2
MARGIN_LEFT=10
MARGIN_TOP=10
LABEL_WIDTH=90
LABEL_HEIGHT=60
QR_COLUMN=Codigo
QR_SIZE=18
QUANTITIES=2|1|1|4|1
EXTRA_TEXTS={Descripcion}|{Descripcion} - Lote {Lote}||{Ubicacion}|
FONT_SIZES=12|10|12|14|12
WINDOW_WIDTH=1200
WINDOW_HEIGHT=700
```

| Clave | Significado |
|---|---|
| `SHEET_WIDTH` / `SHEET_HEIGHT` | Tamaño de la hoja en mm (210 × 297 = A4). |
| `LABELS_PER_ROW` / `LABELS_PER_COL` | Rejilla de etiquetas. Su producto es la capacidad de la hoja. |
| `GAP_MM` | Separación entre etiquetas, en mm. |
| `MARGIN_LEFT` / `MARGIN_TOP` | Márgenes de la hoja, en mm. |
| `LABEL_WIDTH` / `LABEL_HEIGHT` | Tamaño de cada etiqueta, en mm. |
| `QR_COLUMN` | Nombre de la columna del Excel que se codifica en el QR. |
| `QR_SIZE` | Lado del QR, en mm. |
| `QUANTITIES` | Copias por registro, separadas por `|`, en el orden de las filas del Excel. |
| `EXTRA_TEXTS` | Texto adicional por registro, separado por `|`. Admite `{columna}`. Un campo vacío = sin texto. |
| `FONT_SIZES` | Tamaño de letra por registro, separado por `|`. |
| `WINDOW_WIDTH` / `WINDOW_HEIGHT` | Última geometría de la ventana. |

Las tres claves con `|` dependen del Excel cargado, así que se aplican **después** de
abrirlo. En los textos, las barras verticales y las contrabarras se escapan como `\|` y
`\\`.

Si el fichero no existe, la aplicación arranca con los valores por defecto y lo crea en
el primer guardado. Borrarlo equivale a restablecer la configuración.

### `icono.png`
Opcional. Si está presente, se usa como icono de la ventana y de la barra de tareas. Si
falta, el programa continúa sin él.

---

## Flujo de datos

```
Etiquetas.xlsx ──> pandas.DataFrame ──> LabelConfig ──> generar_hoja_etiquetas()
                                             ▲                     │
                                             │                     ├──> vista previa (con borde)
                                    labels.txt (persistencia)      ├──> PDF (ReportLab)
                                                                   └──> impresora (win32print)
```
