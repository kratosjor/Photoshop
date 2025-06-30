#target photoshop

// Ruta a la plantilla
var rutaPlantilla = "C:/Users/Jordan/Desktop/Cursos Programacion/Photoshop/plantilla.psd";
var archivo = new File(rutaPlantilla);

// Documento destino (activo)
var docDestino = app.activeDocument;

// Variable para el documento de plantilla
var docPlantilla = null;

// Verificar que el archivo existe y abrirlo
if (archivo.exists) {
    docPlantilla = app.open(archivo);
} else {
    alert("❌ Archivo plantilla no encontrado:\n" + rutaPlantilla);
    throw new Error("Archivo plantilla no encontrado");
}

// Crear ventana ScriptUI
var win = new Window("palette", "Selecciona capas para agregar", undefined);
win.orientation = "column";

// Obtener capas planas de la plantilla
var capas = docPlantilla.artLayers;

// Crear un botón por cada capa en la plantilla
for (var i = 0; i < capas.length; i++) {
    (function(i) {
        var capa = capas[i];
        var btn = win.add("button", undefined, capa.name);
        btn.onClick = function() {
            // Duplicar capa al documento destino
            capa.duplicate(docDestino);
            alert("Capa '" + capa.name + "' agregada.");
        };
    })(i);
}

// Botón para cerrar ventana y cerrar plantilla
var cerrarBtn = win.add("button", undefined, "Cerrar");
cerrarBtn.onClick = function() {
    docPlantilla.close(SaveOptions.DONOTSAVECHANGES);
    win.close();
};

win.center();
win.show();
