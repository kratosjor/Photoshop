import tkinter as tk
from tkinter import messagebox, filedialog
import win32com.client
import os
import sys
import datetime
import shutil
from tkinter import simpledialog


#######################################


# Para compatibilidad con PyInstaller
#buscara los archivos templates necesarios
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Rutas de recursos
RUTA_PLANTILLA = resource_path("plantilla.psd")
RUTA_ABR = resource_path("Trane_Brushes.abr")
RUTA_SENSORES = resource_path("sensores")

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
        # Conectar a Photoshop
        psApp = win32com.client.Dispatch("Photoshop.Application")
        psApp.Visible = True

        if psApp.Documents.Count == 0:
            messagebox.showerror("Error", "Abre o crea un documento en Photoshop antes de continuar.")
            return

        # Código JSX para crear los grupos
        jsx_code = """
        var doc = app.activeDocument;
        var mainGroup = doc.layerSets.add();
        mainGroup.name = "Floor Section";

        var archGroup = mainGroup.layerSets.add();
        archGroup.name = "ARCH";

        var hvacGroup = mainGroup.layerSets.add();
        hvacGroup.name = "HVAC";

        
        """
        psApp.DoJavaScript(jsx_code)

        messagebox.showinfo("Éxito", "Se creó el grupo 'Floor Section' con subgrupos 'HVAC' y 'ARCH'.")

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

#########################################################################
# Exportar a HTML con opciones
#########################################################################

def exportar_completo(tipo_html="SC", exportar_html=True):
    try:
        import tkinter as tk
        from tkinter import simpledialog, messagebox
        import win32com.client
        import os, shutil, datetime

        RUTA_SENSORES = resource_path("sensores")
        psApp = win32com.client.Dispatch("Photoshop.Application")

        if psApp.Documents.Count == 0:
            messagebox.showerror("Error", "No hay documento abierto.")
            return

        doc = psApp.ActiveDocument
        if not doc.FullName:
            messagebox.showerror("Error", "Debes guardar el documento PSD primero.")
            return

        ruta_psd = doc.FullName
        carpeta_base = os.path.dirname(ruta_psd)
        fecha_actual = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")

        # Detectar grupo seleccionado
        nombre_grupo_original = psApp.DoJavaScript("""
        var sel = app.activeDocument.activeLayer;
        while (sel.parent != app.activeDocument) {
            sel = sel.parent;
        }
        sel.name;
        """)

        # Pedir nombre de exportación
        root = tk.Tk()
        root.withdraw()
        nombre_export = simpledialog.askstring(
            "Nombre de exportación",
            f"Grupo seleccionado: {nombre_grupo_original}\n\nIngrese nombre para exportar (ej: piso_1_cliente):"
        )
        if not nombre_export:
            messagebox.showwarning("Cancelado", "Exportación cancelada por el usuario.")
            return
        nombre_export = nombre_export.strip()

        # Crear carpetas base según tipo
        nombre_carpeta_base = f"floorplans_{tipo_html}"
        carpeta_floorplans = os.path.join(carpeta_base, nombre_carpeta_base)

        # Crear subcarpetas HVAC y No-HVAC
        carpeta_hvac = os.path.join(carpeta_floorplans, f"{nombre_export}_HVAC")
        carpeta_arch = os.path.join(carpeta_floorplans, f"{nombre_export}_No-HVAC")

        # Crear carpetas si no existen
        os.makedirs(carpeta_hvac, exist_ok=True)
        os.makedirs(carpeta_arch, exist_ok=True)


        options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
        options.Format = 13
        options.PNG8 = False
        options.Transparency = True
        options.Quality = 100

        def crear_html(ruta_destino, nombre_imagen):
            sensores_html = """
            <img class="sensor" src="Temp_Sensor.png" style="top:203px;left:33px;"/>
            <img class="sensor" src="Sensor.png" style="top:179px;left:33px;z-index:101;"/>
            <img class="sensor" src="C02_Sensor.png" style="top:254px;left:33px;z-index:102;"/>
            <img class="sensor" src="Humidity_Sensor.png" style="top:227px;left:33px;z-index:103;"/>
            """
            nav_overlay = '<div class="nav-overlay">[ES Navigation Controls]</div>' if tipo_html == "ES" else ""
            html = f"""<!DOCTYPE html>
<html>
<head>
    <title>{nombre_export} - {tipo_html}</title>
    <meta name="CGFVersion" content="4.0"/>
    <meta name="Description" content="Exported Graphic"/>
    <meta name="Template" content="false"/>
    <meta name="CreationTime" content="{fecha_actual}"/>
    <meta name="ModificationTime" content="{fecha_actual}"/>
    <meta name="{tipo_html}Version" content=""/>
    <meta name="TGEVersion" content="Centralized Services Graphics"/>
    <meta name="ScrollLength" content="788"/>
    <meta name="ScrollWidth" content="1613"/>
    <style>
        body {{ margin: 0; padding: 0; background-color: #333; overflow: hidden; }}
        .hvac-container {{ position: relative; height: 788px; width: 1613px; }}
        .hvac-main {{ position: absolute; top: 0; left: 0; z-index: 12; }}
        .sensor {{ position: absolute; height: 22px; width: 22px; z-index: 100; }}
        .nav-overlay {{ position: absolute; top: 10px; right: 10px; z-index: 200; }}
    </style>
</head>
<body>
    <div class="hvac-container">
        <img class="hvac-main" src="{nombre_imagen}" alt="HVAC System"/>
        {sensores_html}
        {nav_overlay}
    </div>
</body>
</html>"""
            with open(os.path.join(ruta_destino, f"{nombre_export}_{tipo_html}.html"), "w", encoding="utf-8") as f:
                f.write(html)

        # Export HVAC (HVAC + ARCH visibles)
        psApp.DoJavaScript(f"""
        var doc = app.activeDocument;
        for (var i = 0; i < doc.layerSets.length; i++) {{
            doc.layerSets[i].visible = false;
        }}
        var grupo = doc.layerSets.getByName("{nombre_grupo_original}");
        grupo.visible = true;
        for (var j = 0; j < grupo.layerSets.length; j++) {{
            grupo.layerSets[j].visible = true;
        }}
        """)
        ruta_hvac_img = os.path.join(carpeta_hvac, f"{nombre_export}.png")
        doc.Export(ExportIn=ruta_hvac_img, ExportAs=2, Options=options)
        crear_html(carpeta_hvac, os.path.basename(ruta_hvac_img))

        # Export ARCH (solo ARCH visible)
        psApp.DoJavaScript(f"""
        var doc = app.activeDocument;
        for (var i = 0; i < doc.layerSets.length; i++) {{
            doc.layerSets[i].visible = false;
        }}
        var grupo = doc.layerSets.getByName("{nombre_grupo_original}");
        grupo.visible = true;
        for (var j = 0; j < grupo.layerSets.length; j++) {{
            grupo.layerSets[j].visible = false;
            if (grupo.layerSets[j].name == "ARCH") {{
                grupo.layerSets[j].visible = true;
            }}
        }}
        """)
        ruta_arch_img = os.path.join(carpeta_arch, f"{nombre_export}_ARCH.png")
        doc.Export(ExportIn=ruta_arch_img, ExportAs=2, Options=options)
        crear_html(carpeta_arch, os.path.basename(ruta_arch_img))

        # Copiar sensores
        sensores = ["Temp_Sensor.png", "Sensor.png", "C02_Sensor.png", "Humidity_Sensor.png"]
        for carpeta_dest in [carpeta_hvac, carpeta_arch]:
            for sensor in sensores:
                origen = os.path.join(RUTA_SENSORES, sensor)
                if os.path.exists(origen):
                    shutil.copy2(origen, os.path.join(carpeta_dest, sensor))

        # Restaurar visibilidad
        psApp.DoJavaScript("""
        var doc = app.activeDocument;
        for (var i = 0; i < doc.layerSets.length; i++) {
            doc.layerSets[i].visible = true;
        }
        """)

        # Mensaje final
        messagebox.showinfo("Exportación finalizada",
            f"Exportación completa:\n\nHVAC: {carpeta_hvac}\nNo-HVAC: {carpeta_arch}"
        )

    except Exception as e:
        messagebox.showerror("Error", f"Error durante la exportación:\n\n{str(e)}")


###########################################################################
#exportacion isolate
##$###########################################################################
def exportar_isolate_folder(tipo_html="SC"):
    try:
        psApp = win32com.client.Dispatch("Photoshop.Application")

        if psApp.Documents.Count == 0:
            messagebox.showerror("Error", "No hay documentos abiertos.")
            return

        doc = psApp.ActiveDocument
        if not doc.FullName:
            messagebox.showerror("Error", "Debes guardar el documento PSD primero.")
            return

        ruta_psd = doc.FullName
        carpeta_base = os.path.dirname(ruta_psd)
        fecha_actual = datetime.datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")

        # Pedir nombre de exportación
        root_tmp = tk.Tk()
        root_tmp.withdraw()
        nombre_export = simpledialog.askstring(
            "Nombre de exportación",
            f"Ingrese nombre para exportar (modo {tipo_html}):"
        )
        root_tmp.destroy()

        if not nombre_export:
            messagebox.showwarning("Cancelado", "Exportación cancelada por el usuario.")
            return
        nombre_export = nombre_export.strip()

        # Carpeta base para exportar (similar a exportar_completo)
        nombre_carpeta_base = f"floorplans_{tipo_html}"
        carpeta_floorplans = os.path.join(carpeta_base, nombre_carpeta_base)

        # Crear carpeta de exportación
        carpeta_export = os.path.join(carpeta_floorplans, nombre_export)
        os.makedirs(carpeta_export, exist_ok=True)

        options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
        options.Format = 13  # PNG
        options.PNG8 = False
        options.Transparency = True
        options.Quality = 100

        # Ruta PNG
        ruta_png = os.path.join(carpeta_export, f"{nombre_export}.png")

        # Exportar la vista actual (sin tocar visibilidad)
        doc.Export(ExportIn=ruta_png, ExportAs=2, Options=options)

        # Función para crear el HTML con sensores
        def crear_html(ruta_destino, nombre_imagen):
            sensores_html = """
            <img class="sensor" src="Temp_Sensor.png" style="top:203px;left:33px;"/>
            <img class="sensor" src="Sensor.png" style="top:179px;left:33px;z-index:101;"/>
            <img class="sensor" src="C02_Sensor.png" style="top:254px;left:33px;z-index:102;"/>
            <img class="sensor" src="Humidity_Sensor.png" style="top:227px;left:33px;z-index:103;"/>
            """
            nav_overlay = '<div class="nav-overlay">[ES Navigation Controls]</div>' if tipo_html == "ES" else ""
            html = f"""<!DOCTYPE html>
<html>
<head>
    <title>{nombre_export} - {tipo_html}</title>
    <meta name="CGFVersion" content="4.0"/>
    <meta name="Description" content="Exported Graphic"/>
    <meta name="Template" content="false"/>
    <meta name="CreationTime" content="{fecha_actual}"/>
    <meta name="ModificationTime" content="{fecha_actual}"/>
    <meta name="{tipo_html}Version" content=""/>
    <meta name="TGEVersion" content="Centralized Services Graphics"/>
    <meta name="ScrollLength" content="788"/>
    <meta name="ScrollWidth" content="1613"/>
    <style>
        body {{ margin: 0; padding: 0; background-color: #333; overflow: hidden; }}
        .hvac-container {{ position: relative; height: 788px; width: 1613px; }}
        .hvac-main {{ position: absolute; top: 0; left: 0; z-index: 12; }}
        .sensor {{ position: absolute; height: 22px; width: 22px; z-index: 100; }}
        .nav-overlay {{ position: absolute; top: 10px; right: 10px; z-index: 200; }}
    </style>
</head>
<body>
    <div class="hvac-container">
        <img class="hvac-main" src="{nombre_imagen}" alt="HVAC System"/>
        {sensores_html}
        {nav_overlay}
    </div>
</body>
</html>"""
            with open(os.path.join(ruta_destino, f"{nombre_export}_{tipo_html}.html"), "w", encoding="utf-8") as f:
                f.write(html)

        # Copiar sensores
        sensores = ["Temp_Sensor.png", "Sensor.png", "C02_Sensor.png", "Humidity_Sensor.png"]
        for sensor in sensores:
            origen = os.path.join(RUTA_SENSORES, sensor)
            if os.path.exists(origen):
                shutil.copy2(origen, os.path.join(carpeta_export, sensor))

        # Crear el archivo HTML
        crear_html(carpeta_export, os.path.basename(ruta_png))

        messagebox.showinfo("Exportación finalizada",
            f"Exportación isolate completa en:\n{carpeta_export}")

    except Exception as e:
        messagebox.showerror("Error", f"Error durante la exportación isolate:\n{str(e)}")



#######################################################################
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

#Boton para exportar a HTML
# Crear un frame para los botones de exportación
export_frame = tk.Frame(root)  
export_frame.pack(padx=20, pady=(5, 10))

tk.Button(export_frame, text="Exportar SC", 
          command=lambda: exportar_completo("SC", True)).pack(side="left", padx=5)

tk.Button(export_frame, text="Exportar ES", 
          command=lambda: exportar_completo("ES", True)).pack(side="left", padx=5)


# Frame para botones Isolate
isolate_frame = tk.Frame(root)
isolate_frame.pack(padx=20, pady=5)

tk.Button(isolate_frame, text="Isolate SC", command=lambda: exportar_isolate_folder("SC")).pack(side="left", padx=5)
tk.Button(isolate_frame, text="Isolate ES", command=lambda: exportar_isolate_folder("ES")).pack(side="left", padx=5)


#ESTETICA HERRAMIENTA
root.configure(bg="#2c3e50")
root.option_add("*Font", "Helvetica 12")
root.option_add("*Button*highlightBackground", "#2c3e50")
root.option_add("*Button*highlightColor", "#2980b9")



root.mainloop()
