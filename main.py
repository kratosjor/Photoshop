import tkinter as tk
from tkinter import messagebox, filedialog
import win32com.client
import os
import sys

# Para compatibilidad con PyInstaller
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Rutas de recursos
RUTA_PLANTILLA = resource_path("plantilla.psd")
RUTA_ABR = resource_path("Trane_Brushes.abr")

# Capas por categoría
capas_por_categoria = {
    "Floorplan": {
        "Walls": "Walls",
        "Floor": "Floor",
        "Lower_Level": None,
        "1_line_ Details": "Stairs/Elevetors/End Caps"
    },
    "Ductwork": {
        "2px_Duct": "2px brush",
        "4px_Duct": "Ductwork",
        "2px_Duct_EXH": None
    },
    "Diffusers": {
        "Diffusers": "Circle Diffuser",
        "Diffusers_EXH": None
    },
    "Equipment": {
        "GreenUnit_AHU_RTU": "Green Unit Horizontal",
        "Blue_Unit": "Equipment Box",
        "Orange_Unit": "Orange Unit",
        "Magenta_Unit_EF": "EF"
    },
    "Keymaps & Menus": {
        "Floor_keymap": None,
        "Border_keymap": None,
        "Selected Area": None,
        "floor_menu": None,
        "Border": None,
        "Non_Selectable": None,
    }
}

def verificar_o_cargar_brushes(psApp, archivo_abr, brush_ejemplo="Walls"):
    jsx_code = f"""
    var brushName = "{brush_ejemplo}";
    var abrPath = "{archivo_abr.replace("\\\\", "/").replace("\\", "/")}";
    var brushExists = true;

    try {{
        var desc = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putName(charIDToTypeID("Brsh"), brushName);
        desc.putReference(charIDToTypeID("null"), ref);
        executeAction(charIDToTypeID("slct"), desc, DialogModes.NO);
    }} catch(e) {{
        brushExists = false;
    }}

    if (!brushExists) {{
        var abrFile = new File(abrPath);
        if (abrFile.exists) {{
            app.load(abrFile);
            alert("✅ Brushes importados automáticamente desde: " + abrPath);
        }} else {{
            alert("❌ Archivo ABR no encontrado: " + abrPath);
        }}
    }}
    """
    psApp.DoJavaScript(jsx_code)

def duplicar_capa_y_brush(nombre_capa, brush_name):
    try:
        psApp = win32com.client.Dispatch("Photoshop.Application")
        psApp.Visible = True

        if psApp.Documents.Count == 0:
            messagebox.showerror("Error", "No hay documento abierto para importar capas.")
            return

        docDestino = psApp.ActiveDocument
        display_dialog_backup = psApp.DisplayDialogs
        psApp.DisplayDialogs = 3  # No diálogos

        plantilla_doc = psApp.Open(RUTA_PLANTILLA)

        try:
            capa = plantilla_doc.ArtLayers.Item(nombre_capa)
            capa.Duplicate(docDestino, 2)
        except Exception:
            messagebox.showerror("Error", f"No se encontró la capa '{nombre_capa}' en plantilla.")
            plantilla_doc.Close(2)
            psApp.DisplayDialogs = display_dialog_backup
            return

        plantilla_doc.Close(2)
        psApp.DisplayDialogs = display_dialog_backup

        if brush_name:
            jsx_code = f"""
            var brushName = "{brush_name}";
            var desc = new ActionDescriptor();
            var ref = new ActionReference();
            ref.putName(charIDToTypeID("Brsh"), brushName);
            desc.putReference(charIDToTypeID("null"), ref);
            try {{
                executeAction(charIDToTypeID("slct"), desc, DialogModes.NO);
            }} catch(e) {{
                alert("❌ No se pudo seleccionar el pincel: " + brushName);
            }}
            """
            psApp.DoJavaScript(jsx_code)

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo conectar o manipular Photoshop:\n{e}")

def crear_floor_section():
    try:
        psApp = win32com.client.Dispatch("Photoshop.Application")
        psApp.Visible = True

        if psApp.Documents.Count == 0:
            messagebox.showerror("Error", "Abre o crea un documento para insertar grupos.")
            return

        jsx_code = """
        var doc = app.activeDocument;
        var mainGroup = doc.layerSets.add();
        mainGroup.name = "Floor Section";

        var hvacGroup = mainGroup.layerSets.add();
        hvacGroup.name = "HVAC";

        var archGroup = mainGroup.layerSets.add();
        archGroup.name = "ARCH";
        """
        psApp.DoJavaScript(jsx_code)

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo crear los grupos:\n{e}")

def importar_imagenes_como_capas():
    try:
        psApp = win32com.client.Dispatch("Photoshop.Application")
        psApp.Visible = True

        if psApp.Documents.Count == 0:
            messagebox.showerror("Error", "Debes tener un documento abierto para importar imágenes.")
            return

        archivos = filedialog.askopenfilenames(
            title="Selecciona las imágenes a importar",
            filetypes=[("Imágenes", "*.jpg *.png *.jpeg *.psd *.tif *.bmp"), ("Todos los archivos", "*.*")]
        )

        if not archivos:
            return

        doc_actual = psApp.ActiveDocument

        for ruta in archivos:
            imagen_doc = psApp.Open(ruta)
            imagen_doc.ArtLayers[0].Duplicate(doc_actual, 2)
            imagen_doc.Close(2)

        messagebox.showinfo("Importación completa", f"{len(archivos)} imagen(es) importadas como capas.")

    except Exception as e:
        messagebox.showerror("Error", f"No se pudieron importar imágenes:\n{e}")

# GUI
root = tk.Tk()
root.title("Importar capas desde plantilla.psd")

# Mensaje inicial
messagebox.showinfo(
    "Antes de comenzar",
    "💾 Antes de comenzar a trabajar en el nuevo PSD, asegúrate de haber guardado el archivo, oni-chan."
)

# Conectar a Photoshop y configurar perfiles de color
try:
    psApp = win32com.client.Dispatch("Photoshop.Application")
    psApp.Visible = True

    # Desactiva advertencias de perfil incrustado y asigna sRGB automáticamente
    psApp.DoJavaScript("""
    app.colorSettings = app.colorSettings;
    app.colorSettings.missingProfiles = false;
    app.colorSettings.profileMismatch = false;
    app.colorSettings.askWhenOpening = false;
    app.colorSettings.askWhenPasting = false;
    app.colorSettings.RGBProfile = "sRGB IEC61966-2.1";
    """)

    verificar_o_cargar_brushes(psApp, RUTA_ABR, brush_ejemplo="Walls")

except Exception as e:
    messagebox.showerror("Error", f"No se pudo conectar a Photoshop:\n{e}")

# Botón para crear grupos
tk.Button(root, text="Crear grupo 'Floor Section' con HVAC y ARCH", command=crear_floor_section).pack(padx=20, pady=(10, 5))

# Botón para importar imágenes
tk.Button(root, text="Importar imágenes como capas", command=importar_imagenes_como_capas).pack(padx=20, pady=5)

# Menús desplegables por categoría
for categoria, capas in capas_por_categoria.items():
    contenedor = tk.Frame(root)
    contenedor.pack(fill="x", padx=10, pady=5)

    botones_frame = tk.Frame(contenedor)
    botones_frame.pack_forget()

    def crear_toggle(frame=botones_frame, texto=categoria):
        visible = [False]
        def toggle():
            if visible[0]:
                frame.pack_forget()
                boton.config(text=f"▶ {texto}")
            else:
                frame.pack(fill="x")
                boton.config(text=f"▼ {texto}")
            visible[0] = not visible[0]
        return toggle

    boton = tk.Button(contenedor, text=f"▶ {categoria}", anchor="w", command=crear_toggle())
    boton.pack(fill="x")

    for capa_nombre, brush_name in capas.items():
        btn = tk.Button(botones_frame, text=capa_nombre, command=lambda cn=capa_nombre, bn=brush_name: duplicar_capa_y_brush(cn, bn))
        btn.pack(anchor="w", padx=20, pady=2)

root.mainloop()
