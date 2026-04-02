/**
 * MICSA SAFETY CORPORATE IDENTITY SYSTEM v3.0
 * Corregido para Adobe Illustrator 2025+ ExtendScript
 */

#target illustrator

// ============================================================================
// DOCUMENTO CMYK
// ============================================================================
var doc = app.documents.add(DocumentColorSpace.CMYK);
var layer = doc.activeLayer;

// ============================================================================
// COLORES CMYK (todo CMYK para consistencia de impresi\u00f3n)
// ============================================================================
function cmyk(c, m, y, k) {
    var col = new CMYKColor();
    col.cyan = c;
    col.magenta = m;
    col.yellow = y;
    col.black = k;
    return col;
}

var Navy = cmyk(100, 85, 40, 35);
var Gold = cmyk(0, 20, 100, 0);
var White = cmyk(0, 0, 0, 0);
var DarkGray = cmyk(0, 0, 0, 77);
var LightGray = cmyk(0, 0, 0, 7);
var MidGray = cmyk(0, 0, 0, 22);
var FaintGray = cmyk(0, 0, 0, 30);

// ============================================================================
// ARTBOARDS (espaciados horizontalmente)
// ============================================================================
var pieces = [
    { name: "01-Letterhead", w: 595, h: 842 },
    { name: "02-BusinessCard", w: 255, h: 142 },
    { name: "03-Envelope", w: 623, h: 312 },
    { name: "04-Folder", w: 680, h: 907 },
    { name: "05-Credential", w: 241, h: 153 },
    { name: "06-Report", w: 595, h: 842 },
    { name: "07-Manual", w: 595, h: 842 },
    { name: "08-Seal", w: 283, h: 283 },
    { name: "09-Signature", w: 510, h: 227 },
    { name: "10-Proposal", w: 595, h: 842 }
];

var gap = 50;
var curX = 0;
var origins = [];

for (var i = 0; i < pieces.length; i++) {
    var p = pieces[i];
    // Illustrator: [left, top, right, bottom] — Y sube
    var abRect = [curX, p.h, curX + p.w, 0];

    if (i === 0) {
        doc.artboards[0].artboardRect = abRect;
        doc.artboards[0].name = p.name;
    } else {
        doc.artboards.add(abRect);
        doc.artboards[doc.artboards.length - 1].name = p.name;
    }

    origins.push({ x: curX, y: p.h });
    curX += p.w + gap;
}

// ============================================================================
// UTILIDADES DE DIBUJO
// Coordenadas relativas al top-left del artboard activo
// (x crece a la derecha, y crece hacia abajo)
// ============================================================================
var _ox = 0, _oy = 0;

function setAB(idx) {
    _ox = origins[idx].x;
    _oy = origins[idx].y;
    doc.artboards.setActiveArtboardIndex(idx);
}

// Rect\u00e1ngulo: (x, y) desde esquina superior-izquierda del artboard
function R(x, y, w, h, fill, stroke, sw) {
    var r = layer.pathItems.rectangle(_oy - y, _ox + x, w, h);
    if (fill) {
        r.filled = true;
        r.fillColor = fill;
    } else {
        r.filled = false;
    }
    if (stroke) {
        r.stroked = true;
        r.strokeColor = stroke;
        r.strokeWidth = sw || 1;
    } else {
        r.stroked = false;
    }
    return r;
}

// Texto
function T(str, x, y, sz, font, color) {
    var tf = layer.textFrames.add();
    tf.contents = str;
    tf.left = _ox + x;
    tf.top = _oy - y;
    var ca = tf.textRange.characterAttributes;
    ca.size = sz;
    ca.fillColor = color;
    try {
        ca.textFont = app.textFonts.getByName(font);
    } catch (e) {
        try {
            ca.textFont = app.textFonts.getByName("ArialMT");
        } catch (e2) { }
    }
    return tf;
}

// Elipse centrada en (cx, cy) relativo al artboard
function EL(cx, cy, w, h, fill, stroke, sw) {
    var e = layer.pathItems.ellipse(_oy - cy + h / 2, _ox + cx - w / 2, w, h);
    if (fill) {
        e.filled = true;
        e.fillColor = fill;
    } else {
        e.filled = false;
    }
    if (stroke) {
        e.stroked = true;
        e.strokeColor = stroke;
        e.strokeWidth = sw || 1;
    } else {
        e.stroked = false;
    }
    return e;
}

// ============================================================================
// 01 — LETTERHEAD (Hoja membretada A4)
// ============================================================================
setAB(0);
R(0, 0, 595, 99, Navy);
T("MICSA", 42, 25, 32, "Montserrat-Bold", White);
T("SAFETY DIVISION", 42, 60, 11, "OpenSans-Regular", Gold);
T("www.micsasafety.com.mx", 440, 18, 9, "OpenSans-Regular", White);
T("contacto@micsasafety.com.mx", 440, 35, 9, "OpenSans-Regular", White);
T("+52 (844) 637-0000", 440, 52, 9, "OpenSans-Regular", White);
R(0, 99, 595, 2, Gold);
R(0, 101, 595, 20, LightGray);
T("Monclova, Coahuila  \u00B7  Fecha: ____  \u00B7  Folio: ____", 42, 107, 8, "OpenSans-Regular", DarkGray);
R(20, 121, 3, 630, Gold);
T("ASUNTO:", 42, 138, 8, "Montserrat-Bold", DarkGray);
R(42, 150, 360, 1, DarkGray);
R(42, 165, 400, 400, null, MidGray, 1);
T("[Contenido de la carta]", 52, 340, 9, "OpenSans-Regular", FaintGray);
// Footer
R(0, 751, 595, 91, Navy);
R(0, 751, 595, 2, Gold);
T("Monclova, Coahuila  \u00B7  RFC: MSD-000000-XXX", 42, 762, 7, "OpenSans-Regular", White);
T("No cuidamos instalaciones. Controlamos operaciones.", 42, 778, 7, "Montserrat-Bold", Gold);
T("Disciplina  \u00B7  Criterio  \u00B7  Control Operativo", 42, 792, 6, "OpenSans-Regular", White);

// ============================================================================
// 02 — BUSINESS CARD (Tarjeta de presentaci\u00f3n)
// ============================================================================
setAB(1);
R(0, 0, 57, 142, Navy);
T("MICSA", 71, 16, 14, "Montserrat-Bold", Navy);
T("SAFETY DIVISION", 71, 35, 7, "OpenSans-Regular", Gold);
T("contacto@micsasafety.com.mx", 71, 71, 6, "OpenSans-Regular", DarkGray);
T("+52 (844) 637-0000", 71, 82, 6, "OpenSans-Regular", DarkGray);
T("www.micsasafety.com.mx", 71, 93, 6, "OpenSans-Regular", DarkGray);
T("Gerardo Guzm\u00E1n Alvarado", 71, 112, 7, "Montserrat-Bold", DarkGray);
T("Director de Seguridad", 71, 124, 6, "OpenSans-Regular", Gold);

// ============================================================================
// 03 — ENVELOPE (Sobre)
// ============================================================================
setAB(2);
R(0, 0, 623, 57, Navy);
T("MICSA SAFETY DIVISION", 28, 15, 12, "Montserrat-Bold", White);
T("www.micsasafety.com.mx", 460, 15, 7, "OpenSans-Regular", Gold);
T("[Destinatario]", 57, 90, 9, "OpenSans-Regular", DarkGray);
T("[Direcci\u00F3n]", 57, 106, 9, "OpenSans-Regular", DarkGray);
T("[Ciudad, Estado, C\u00F3digo Postal]", 57, 122, 9, "OpenSans-Regular", DarkGray);
R(0, 304, 623, 8, Gold);

// ============================================================================
// 04 — FOLDER (Carpeta)
// ============================================================================
setAB(3);
R(0, 0, 680, 907, Navy);
R(0, 85, 680, 5, Gold);
R(0, 766, 680, 5, Gold);
T("MICSA SAFETY DIVISION", 57, 32, 18, "Montserrat-Bold", White);
T("Sistema Integral de Seguridad", 57, 60, 10, "OpenSans-Regular", Gold);
T("No cuidamos instalaciones.", 57, 120, 12, "Montserrat-Bold", Gold);
T("Controlamos operaciones.", 57, 142, 12, "Montserrat-Bold", Gold);
T("www.micsasafety.com.mx  |  +52 (844) 637-0000", 57, 880, 8, "OpenSans-Regular", White);

// ============================================================================
// 05 — CREDENTIAL (Credencial)
// ============================================================================
setAB(4);
R(8, 8, 71, 93, MidGray, DarkGray, 1);
T("[Foto]", 25, 48, 8, "OpenSans-Regular", DarkGray);
R(85, 8, 148, 105, Navy);
T("Gerardo Guzm\u00E1n Alvarado", 92, 24, 9, "Montserrat-Bold", White);
T("Director de Seguridad", 92, 40, 7, "OpenSans-Regular", Gold);
T("ID: MICSA-0001", 92, 58, 7, "OpenSans-Regular", White);
T("Experiencia: 5 a\u00F1os", 92, 72, 7, "OpenSans-Regular", White);
T("Fuerza Civil | Protecci\u00F3n Civil", 92, 84, 6, "OpenSans-Regular", White);
T("V\u00E1lida hasta: 2027", 92, 96, 6, "OpenSans-Regular", Gold);

// ============================================================================
// 06 — REPORT (Reporte de Seguridad)
// ============================================================================
setAB(5);
R(0, 0, 595, 141, Navy);
T("REPORTE DE SEGURIDAD", 57, 45, 24, "Montserrat-Bold", White);
T("MICSA Safety Division", 57, 90, 12, "OpenSans-Regular", Gold);
R(0, 141, 595, 5, Gold);
T("Per\u00EDodo: ____________________", 57, 172, 10, "OpenSans-Regular", DarkGray);
T("Sitio: ____________________", 57, 190, 10, "OpenSans-Regular", DarkGray);
T("Clasificaci\u00F3n: Confidencial", 57, 208, 9, "OpenSans-Regular", Gold);
R(57, 230, 480, 450, null, MidGray, 1);
T("[Contenido del Reporte]", 200, 430, 10, "OpenSans-Regular", FaintGray);
T("MICSA Safety Division  |  Confidencial  |  2026", 57, 810, 8, "OpenSans-Regular", DarkGray);

// ============================================================================
// 07 — MANUAL (Manual Operativo)
// ============================================================================
setAB(6);
R(0, 0, 595, 842, Navy);
R(0, 230, 595, 11, Gold);
T("MANUAL OPERATIVO", 57, 120, 28, "Montserrat-Bold", White);
T("MICSA SAFETY DIVISION", 57, 165, 14, "OpenSans-Regular", Gold);
T("No cuidamos instalaciones.", 57, 265, 12, "Montserrat-Bold", Gold);
T("Controlamos operaciones.", 57, 288, 12, "Montserrat-Bold", Gold);
T("Versi\u00F3n: 1.0  |  A\u00F1o: 2026", 57, 762, 9, "OpenSans-Regular", White);
T("Clasificaci\u00F3n: Confidencial", 57, 780, 9, "OpenSans-Regular", White);

// ============================================================================
// 08 — SEAL (Sello corporativo)
// ============================================================================
setAB(7);
EL(141.5, 141.5, 170, 170, Navy, Gold, 3);
T("MICSA", 101, 124, 14, "Montserrat-Bold", White);
T("SAFETY", 103, 146, 10, "OpenSans-Regular", Gold);
T("2026", 117, 164, 8, "OpenSans-Regular", White);

// ============================================================================
// 09 — SIGNATURE (Bloque de firma autorizada)
// ============================================================================
setAB(8);
R(0, 0, 510, 34, Navy);
T("FIRMA AUTORIZADA", 14, 10, 9, "Montserrat-Bold", White);
R(0, 34, 510, 2, Gold);
R(14, 70, 160, 1, DarkGray);
T("Firma", 14, 76, 8, "OpenSans-Regular", DarkGray);
T("Gerardo Guzm\u00E1n Alvarado", 14, 100, 8, "OpenSans-Regular", DarkGray);
T("Director de Seguridad", 14, 116, 8, "OpenSans-Regular", DarkGray);
T("Fecha: ____________________", 14, 134, 8, "OpenSans-Regular", DarkGray);
R(340, 55, 156, 156, MidGray, DarkGray, 1);
T("[SELLO]", 388, 128, 10, "OpenSans-Regular", DarkGray);

// ============================================================================
// 10 — PROPOSAL (Propuesta Comercial)
// ============================================================================
setAB(9);
R(0, 0, 113, 842, Navy);
R(108, 0, 5, 842, Gold);
T("PROPUESTA", 170, 180, 26, "Montserrat-Bold", Navy);
T("COMERCIAL", 170, 218, 26, "Montserrat-Bold", Navy);
T("SISTEMA INTEGRAL DE", 170, 265, 14, "OpenSans-Regular", Gold);
T("SEGURIDAD", 170, 286, 14, "OpenSans-Regular", Gold);
T("MICSA SAFETY DIVISION", 170, 320, 12, "Montserrat-Bold", Navy);
T("Cliente: ____________________", 170, 370, 10, "OpenSans-Regular", DarkGray);
T("Proyecto: ____________________", 170, 390, 10, "OpenSans-Regular", DarkGray);
T("Fecha: ____________________", 170, 410, 10, "OpenSans-Regular", DarkGray);
R(0, 771, 595, 2, Gold);
T("No cuidamos instalaciones. Controlamos operaciones.", 170, 784, 11, "Montserrat-Bold", Navy);
T("www.micsasafety.com.mx  |  +52 (844) 637-0000", 170, 804, 8, "OpenSans-Regular", DarkGray);

// ============================================================================
// GUARDAR
// ============================================================================
try {
    var saveOpts = new IllustratorSaveOptions();
    var file = new File(Folder.desktop + "/MICSA_CORPORATE_IDENTITY.ai");
    doc.saveAs(file, saveOpts);
    alert("\u2713 COMPLETO \u2014 10 artboards generados:\n\n" +
          "1. Letterhead (Membretada)\n" +
          "2. Business Card (Tarjeta)\n" +
          "3. Envelope (Sobre)\n" +
          "4. Folder (Carpeta)\n" +
          "5. Credential (Credencial)\n" +
          "6. Report (Reporte)\n" +
          "7. Manual Operativo\n" +
          "8. Seal (Sello)\n" +
          "9. Signature (Firma)\n" +
          "10. Proposal (Propuesta)\n\n" +
          "Colores: Navy Blue + Gold (CMYK)\n" +
          "Guardado en Desktop.");
} catch (e) {
    alert("Documento generado con 10 artboards.\n" +
          "Gu\u00E1rdalo manualmente: File > Save As\n\n" +
          "Detalle: " + e.message);
}
