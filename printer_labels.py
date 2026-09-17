import tempfile
import win32print
import win32ui
from PIL import ImageWin

def imprimir_hoja_etiquetas(img_hoja, nombre_impresora=None):
    """Imprime la hoja de etiquetas"""
    printer_name = nombre_impresora or win32print.GetDefaultPrinter()
    
    # Guardar imagen temporalmente
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        img_path = tmp.name
        img_hoja.save(img_path, dpi=(300, 300))
    
    try:
        # Configurar impresión
        hprinter = win32print.OpenPrinter(printer_name)
        try:
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)
            hdc.StartDoc("Hoja de Etiquetas")
            hdc.StartPage()
            
            # Obtener dimensiones de la página
            printer_width = hdc.GetDeviceCaps(110)  # HORZRES
            printer_height = hdc.GetDeviceCaps(111)  # VERTRES
            
            # Escalar imagen para que quepa en la página
            img = ImageWin.Image(0, 0, img_path)
            img.draw(hdc.GetHandleOutput(), (0, 0, printer_width, printer_height))
            
            hdc.EndPage()
            hdc.EndDoc()
            hdc.DeleteDC()
        finally:
            win32print.ClosePrinter(hprinter)
    finally:
        # Limpiar archivo temporal
        import os
        os.unlink(img_path)

def exportar_hoja_pdf(img_hoja, output_path, config):
    """Exporta la hoja de etiquetas a PDF usando ReportLab"""
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        img_path = tmp.name
        img_hoja.save(img_path, dpi=(300, 300))
    
    try:
        c = canvas.Canvas(output_path, pagesize=(config.sheet_width_mm * mm, 
                                                 config.sheet_height_mm * mm))
        c.drawImage(img_path, 0, 0, 
                   width=config.sheet_width_mm * mm,
                   height=config.sheet_height_mm * mm)
        c.save()
    finally:
        import os
        os.unlink(img_path)