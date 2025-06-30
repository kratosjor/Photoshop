import tkinter as tk
from tkinter import messagebox
import win32com.client
import os
import sys

# Ruta relativa a plantilla.psd compatible con PyInstaller
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # Cuando está empaquetado
    except AttributeError:
        base_path = os.path.abspath(".")  # Cuando se ejecuta como .py
    return os.path.join(base_path, relative_path)

# Ruta plantilla PSD
RUTA_PLANTILLA = resource_path("plantilla.psd")

# Capas agrupadas por categoría, con brush opcional
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

def duplicar_capa_y_brush(nombre_capa, brush_name):
    try:
        psApp = win32com.client.Dispatch("Photoshop.Application")
        psApp.Visible = True

        if psApp.Documents.Count == 0:
            messagebox.showerror("Error", "No hay documento abierto para importar capas.")
            return

        docDestino = psApp.ActiveDocument

        display_dialog_backup = psApp.DisplayDialogs
        psApp.DisplayDialogs = 3  # NEVER

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

# GUI Tkinter
root = tk.Tk()
root.title("Importar capas desde plantilla.psd")

botones_frame = tk.Frame(root)
botones_creados = False
botones_visibles = False

def toggle_botones():
    global botones_creados, botones_visibles

    if not botones_creados:
        for categoria, capas in capas_por_categoria.items():
            lf = tk.LabelFrame(botones_frame, text=categoria, padx=10, pady=5, font=('Segoe UI', 10, 'bold'))
            for capa_nombre, brush_name in capas.items():
                btn = tk.Button(
                    lf,
                    text=capa_nombre,
                    command=lambda cn=capa_nombre, bn=brush_name: duplicar_capa_y_brush(cn, bn)
                )
                btn.pack(anchor="w", padx=10, pady=2)
            lf.pack(fill="x", padx=10, pady=8)
        botones_creados = True

    if not botones_visibles:
        botones_frame.pack(pady=10)
        toggle_btn.config(text="Ocultar capas disponibles")
        botones_visibles = True
    else:
        botones_frame.pack_forget()
        toggle_btn.config(text="Mostrar capas disponibles")
        botones_visibles = False

toggle_btn = tk.Button(root, text="Mostrar capas disponibles", command=toggle_botones)
toggle_btn.pack(padx=20, pady=20)

root.mainloop()
