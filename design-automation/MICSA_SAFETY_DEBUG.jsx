/**
 * MICSA SAFETY - VERSIÓN DEBUG
 * Simplificada para identificar el problema
 */

#target illustrator

alert("PASO 1: Script iniciado");

try {
    // Crear documento
    var doc = app.documents.length > 0 ? app.activeDocument : app.documents.add();
    alert("PASO 2: Documento creado/abierto");

    // Crear colores simples
    var Navy = new CMYKColor();
    Navy.cyan = 100;
    Navy.magenta = 85;
    Navy.yellow = 40;
    Navy.black = 35;

    var White = new RGBColor();
    White.red = 255;
    White.green = 255;
    White.blue = 255;

    alert("PASO 3: Colores creados");

    // Crear UN SOLO artboard de prueba
    var ab = doc.artboards.add([0, 0, 595, -842]);
    ab.name = "TEST";
    doc.artboards.setActiveArtboardIndex(0);

    alert("PASO 4: Artboard creado");

    // Intentar agregar un rectángulo simple
    var rect = ab.pathItems.rectangle(0, 0, 595, 842);
    rect.filled = true;
    rect.fillColor = White;

    alert("PASO 5: Rectángulo blanco agregado");

    // Intentar agregar un rectángulo Navy
    var rect2 = ab.pathItems.rectangle(0, 0, 595, 100);
    rect2.filled = true;
    rect2.fillColor = Navy;

    alert("PASO 6: Rectángulo Navy agregado");

    // Intentar guardar
    var file = new File(Folder.desktop + "/MICSA_TEST.ai");
    doc.saveAs(file);

    alert("PASO 7: ÉXITO - Documento guardado en Desktop");

} catch(e) {
    alert("ERROR en paso anterior:\n" + e.message + "\n\nLínea: " + e.line);
}
