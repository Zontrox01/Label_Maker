# Label_Maker
Aplicación de escritorio que convierte una hoja de Excel en una plancha de etiquetas con
código QR, lista para imprimir o exportar a PDF.

Pensada para almacén, laboratorio e inventario: se parte de un listado de artículos, se
elige qué columna va codificada en el QR, cuántas copias hace falta de cada uno y qué
texto acompaña a cada etiqueta.

![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![Licencia](https://img.shields.io/badge/licencia-MIT-green)
![Plataforma](https://img.shields.io/badge/plataforma-Windows%20%7C%20PDF%20multiplataforma-lightgrey)

---

## Características

- **Lee cualquier Excel.** Los nombres de las columnas los pones tú; la aplicación se
  adapta a lo que encuentre.
- **Cantidad por registro.** Dos etiquetas de un artículo y siete de otro en la misma hoja.
- **Texto adicional por registro,** con variables del propio Excel:
  `{Descripcion} - Lote {Lote}`.
- **Tamaño de letra independiente** para cada registro.
- **Hoja totalmente configurable:** tamaño de página, rejilla, márgenes, separación,
  tamaño de la etiqueta y del QR, todo en milímetros.
- **Vista previa** con el contorno de cada etiqueta marcado (el contorno no se imprime).
- **Salida a PDF** con ReportLab o **impresión directa** en Windows.
- **Todo se guarda solo.** Medidas, cantidades y textos se recuperan en el siguiente
  arranque, igual que el último Excel abierto.
- Generación a **300 DPI**, que es lo que hace falta para que un QR pequeño se lea bien.

---

## Requisitos

- Python 3.9 o superior
- Windows para imprimir directamente. En Linux y macOS todo lo demás funciona; solo se
  pierde el botón de impresión (queda la exportación a PDF).

## Instalación

```bash
git clone https://github.com/<usuario>/etiquetas-qr.git
cd etiquetas-qr

python -m venv .venv
# Windows
.venv\Scripts\activate
# Linux / macOS
source .venv/bin/activate

pip install -r requirements_labels.txt
```

## Uso

```bash
python label_maker.py
```

1. **Cargar Excel** y elegir el fichero. Puedes empezar por `datos/Etiquetas.xlsx`, que
   viene de ejemplo.
2. Ajustar las medidas de la hoja y de la etiqueta en el panel izquierdo.
3. Elegir la **columna para el QR**.
4. Para cada registro, indicar cantidad, texto adicional y tamaño de letra.
5. **Vista Previa** para comprobar el resultado.
6. **Guardar PDF** o **Imprimir**.

### Cómo deben ir los datos

La primera hoja del Excel, cabeceras en la fila 1 y un registro por fila:

| Codigo   | Descripcion                        | Lote      | Ubicacion |
|----------|------------------------------------|-----------|-----------|
| ART-0001 | Tornillo hexagonal M8 x 40 zincado | L2026-014 | A-01-03   |
| ART-0002 | Tuerca autoblocante M8 DIN 985     | L2026-014 | A-01-04   |

- Sin filas vacías intercaladas.
- La primera columna se usa como identificador en la lista de cantidades.
- Todo se lee como texto, así que los ceros a la izquierda no se pierden.

En «Texto adicional» cualquier `{Cabecera}` se sustituye por el valor de esa columna en
esa fila. Con la tabla de arriba, `{Descripcion} ({Ubicacion})` produce
`Tornillo hexagonal M8 x 40 zincado (A-01-03)`.

En `FILES.md` está el detalle completo de la plantilla y de los ficheros de configuración.

---

## Estructura

```
label-maker/
├── label_maker.py            # Interfaz gráfica. Punto de entrada.
├── label_generator.py        # Motor de dibujo de etiquetas y hojas.
├── printer_labels.py         # Impresión (Windows) y exportación a PDF.
├── requirements_labels.txt   # Dependencias.
├── FILES.md                  # Descripción detallada de cada fichero.
├── README.md
└── datos/
    ├── Etiquetas.xlsx        # Plantilla de ejemplo.
    ├── excel_labels.txt      # Ruta del último Excel abierto.
    ├── labels.txt            # Configuración guardada.
    └── icono.png             # Icono de la ventana (opcional).
```

`label_generator.py` no depende de Tkinter, así que se puede usar como librería:

```python
import pandas as pd
from label_generator import LabelConfig, generar_hoja_etiquetas

df = pd.read_excel("datos/Etiquetas.xlsx", dtype=str).fillna("")

cfg = LabelConfig(
    labels_per_row=2,
    labels_per_col=4,
    label_width_mm=90,
    label_height_mm=60,
    qr_column="Codigo",
    qr_size_mm=18,
    textos_por_registro=["{Descripcion}"] * len(df),
    tamanios_por_registro=[12] * len(df),
)

hoja = generar_hoja_etiquetas(df, cfg, cantidades_por_registro=[1] * len(df))
hoja.save("etiquetas.png", dpi=(300, 300))
```

---

## Notas y limitaciones

- **Una hoja por tirada.** Si el total de etiquetas supera `filas × columnas`, la
  aplicación avisa e imprime solo las que caben. La paginación automática está pendiente.
- La impresión usa la API de Windows (`pywin32`); en otros sistemas hay que pasar por el PDF.
- El programa busca las fuentes Arial y DejaVu Sans; si no encuentra ninguna, cae en la
  fuente por defecto de Pillow, que ignora el tamaño de letra.
- Corrección de errores del QR en nivel **M** (recupera hasta un 15 % de daño). Para
  etiquetas expuestas a roce o suciedad conviene subirla a **Q** o **H** en
  `generar_etiqueta_con_qr()`.
- Antes de una tirada grande, imprime una hoja de prueba y mide con una regla: algunos
  drivers reescalan la página y desplazan la rejilla.

## Empaquetado (opcional)

```bash
pip install pyinstaller
pyinstaller --noconsole --name EtiquetasQR label_maker.py
```

La carpeta `datos/` se busca junto al ejecutable, así que hay que copiarla al lado del
`.exe` generado.

## Licencia

MIT. Consulta el fichero `LICENSE`.
