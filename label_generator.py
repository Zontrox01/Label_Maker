import os
from PIL import Image, ImageDraw, ImageFont
import qrcode
from dataclasses import dataclass
import pandas as pd

DPI = 300

@dataclass
class LabelConfig:
    # Configuración de la hoja
    sheet_width_mm: float = 210
    sheet_height_mm: float = 297
    
    # Diseño de etiquetas
    labels_per_row: int = 2
    labels_per_col: int = 4
    gap_mm: float = 2
    margin_left_mm: float = 10
    margin_top_mm: float = 10
    label_width_mm: float = 90
    label_height_mm: float = 60
    
    # Contenido
    qr_column: str = ""
    qr_size_mm: float = 30
    
    # Textos y tamaños por registro
    textos_por_registro: list = None  # Lista de textos adicionales por cada registro
    tamanios_por_registro: list = None  # Lista de tamaños de letra por cada registro

def mm_to_px(mm_val, dpi=DPI):
    return int(mm_val / 25.4 * dpi)

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

def get_font(size, bold=False):
    font_name = "arialbd.ttf" if bold else "arial.ttf"
    for nombre in [font_name, "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf", "arial.ttf"]:
        try:
            return ImageFont.truetype(nombre, size)
        except:
            pass
    return ImageFont.load_default()

def generar_etiqueta_con_qr(texto_qr, texto_extra, qr_size_mm, label_width_px, label_height_px, show_border=False, font_size=12):
    img = Image.new("RGB", (label_width_px, label_height_px), color="white")
    draw = ImageDraw.Draw(img)
    
    qr_size_px = mm_to_px(qr_size_mm)
    separacion = 10  # Píxeles de separación entre QR y texto
    
    # Pre-calcular dimensiones del texto si existe
    lines = []
    total_text_height = 0
    max_text_width = 0
    
    if texto_extra and texto_extra.strip():
        font = get_font(font_size)
        max_width = label_width_px - 20
        
        # Dividir texto en líneas
        current_line = ""
        for word in texto_extra.split():
            test_line = current_line + " " + word if current_line else word
            bbox = draw.textbbox((0, 0), test_line, font=font)
            if bbox[2] - bbox[0] <= max_width:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)
        
        # Calcular altura total del texto y ancho máximo
        line_height = font_size + 8
        total_text_height = len(lines) * line_height
        for line in lines:
            bbox = draw.textbbox((0, 0), line, font=font)
            max_text_width = max(max_text_width, bbox[2] - bbox[0])
    
    # Calcular dimensiones del conjunto (QR + texto)
    if texto_qr and texto_extra:
        # Caso: Hay QR y texto
        bloque_width = max(qr_size_px, max_text_width)
        bloque_height = qr_size_px + separacion + total_text_height
    elif texto_qr:
        # Caso: Solo QR
        bloque_width = qr_size_px
        bloque_height = qr_size_px
    elif texto_extra:
        # Caso: Solo texto
        bloque_width = max_text_width
        bloque_height = total_text_height
    else:
        # Caso: Vacío - no hacer nada
        if show_border:
            draw.rectangle([(0, 0), (label_width_px-1, label_height_px-1)], outline="#0066CC", width=2)
        return img
    
    # Centrar el bloque completo
    bloque_x = (label_width_px - bloque_width) // 2
    bloque_y = (label_height_px - bloque_height) // 2
    
    # Dibujar QR (si existe)
    if texto_qr:
        qr = qrcode.QRCode(
            version=None,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=3,
            border=1,
        )
        qr.add_data(texto_qr)
        qr.make(fit=True)
        
        qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")
        qr_img = qr_img.resize((qr_size_px, qr_size_px), Image.NEAREST)
        
        # Posicionar QR dentro del bloque centrado
        qr_x = bloque_x + (bloque_width - qr_size_px) // 2
        qr_y = bloque_y
        img.paste(qr_img, (qr_x, qr_y))
    
    # Dibujar texto (si existe)
    if texto_extra and texto_extra.strip() and lines:
        # Posición Y del texto (después del QR)
        text_y = bloque_y + (qr_size_px if texto_qr else 0) + (separacion if texto_qr else 0)
        
        # Dibujar cada línea centrada dentro del bloque
        for i, line in enumerate(lines):
            bbox = draw.textbbox((0, 0), line, font=font)
            line_width = bbox[2] - bbox[0]
            text_x = bloque_x + (bloque_width - line_width) // 2
            draw.text((text_x, text_y + i * line_height), line, font=font, fill="black")
    
    # Dibujar borde SOLO si show_border es True (solo para vista previa)
    if show_border:
        draw.rectangle([(0, 0), (label_width_px-1, label_height_px-1)], outline="#0066CC", width=2)
    
    return img

def generar_hoja_etiquetas(df, config: LabelConfig, cantidades_por_registro, preview_mode=False):
    """Genera una hoja completa con todas las etiquetas"""
    
    # Convertir medidas a píxeles
    sheet_width_px = mm_to_px(config.sheet_width_mm)
    sheet_height_px = mm_to_px(config.sheet_height_mm)
    label_width_px = mm_to_px(config.label_width_mm)
    label_height_px = mm_to_px(config.label_height_mm)
    gap_px = mm_to_px(config.gap_mm)
    margin_left_px = mm_to_px(config.margin_left_mm)
    margin_top_px = mm_to_px(config.margin_top_mm)
    
    # Crear imagen de la hoja (fondo blanco)
    hoja = Image.new("RGB", (sheet_width_px, sheet_height_px), color="white")
    
    # Generar todas las etiquetas
    todas_las_etiquetas = []
    
    # Preparar listas de textos y tamaños por registro
    textos_por_registro = config.textos_por_registro if config.textos_por_registro else []
    tamanios_por_registro = config.tamanios_por_registro if config.tamanios_por_registro else []
    
    for idx, (_, row) in enumerate(df.iterrows()):
        cantidad = cantidades_por_registro[idx] if idx < len(cantidades_por_registro) else 0
        
        if cantidad <= 0:
            continue
        
        # Obtener texto para QR
        texto_qr = str(row[config.qr_column]) if config.qr_column in df.columns else ""
        
        # Obtener texto adicional para este registro (si existe)
        texto_extra = ""
        if idx < len(textos_por_registro):
            texto_extra_raw = textos_por_registro[idx]
            # Procesar variables en el texto adicional {columna}
            if texto_extra_raw:
                texto_extra = texto_extra_raw
                for col in df.columns:
                    texto_extra = texto_extra.replace(f"{{{col}}}", str(row[col]))
        
        # Obtener tamaño de letra para este registro
        font_size = 12
        if idx < len(tamanios_por_registro):
            try:
                font_size = int(tamanios_por_registro[idx])
            except:
                font_size = 12
        
        # Generar las etiquetas necesarias para este registro
        for _ in range(cantidad):
            etiqueta = generar_etiqueta_con_qr(
                texto_qr, 
                texto_extra,
                config.qr_size_mm,
                label_width_px,
                label_height_px,
                show_border=preview_mode,
                font_size=font_size
            )
            todas_las_etiquetas.append(etiqueta)
    
    # Colocar etiquetas en la hoja
    columna_actual = 0
    fila_actual = 0
    
    for etiqueta in todas_las_etiquetas:
        # Calcular posición
        x = margin_left_px + columna_actual * (label_width_px + gap_px)
        y = margin_top_px + fila_actual * (label_height_px + gap_px)
        
        # Verificar si cabe en la hoja actual
        if x + label_width_px > sheet_width_px:
            columna_actual = 0
            fila_actual += 1
            x = margin_left_px + columna_actual * (label_width_px + gap_px)
            y = margin_top_px + fila_actual * (label_height_px + gap_px)
        
        if y + label_height_px > sheet_height_px:
            # No cabe más en esta hoja (en una implementación real, crearíamos nueva hoja)
            break
        
        # Pegar etiqueta en la hoja
        hoja.paste(etiqueta, (x, y))
        
        # Avanzar a la siguiente posición
        columna_actual += 1
        if columna_actual >= config.labels_per_row:
            columna_actual = 0
            fila_actual += 1
    
    return hoja