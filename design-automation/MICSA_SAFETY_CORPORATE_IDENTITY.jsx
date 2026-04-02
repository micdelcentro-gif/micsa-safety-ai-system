/**
 * MICSA SAFETY CORPORATE IDENTITY SYSTEM GENERATOR v1.0
 * Generador automático de 10 piezas de identidad corporativa en Adobe Illustrator
 *
 * Color Palette: Navy Blue (#1A2B4A) + Gold (#E8A800)
 * Typography: Montserrat Bold (titles), Open Sans (body)
 * Director: Gerardo Guzmán Alvarado
 *
 * INSTRUCTIONS FOR USE:
 * 1. Open Adobe Illustrator 2025 or later
 * 2. Go to File > Scripts > Other Scripts
 * 3. Navigate to this file (MICSA_SAFETY_CORPORATE_IDENTITY.jsx)
 * 4. Click Open or Execute
 * 5. Wait 30-60 seconds for all 10 artboards to be generated
 * 6. Save the document as needed
 */

#target illustrator

// Verificar si Illustrator está disponible
if (app.documents.length === 0) {
    var doc = app.documents.add();
} else {
    var doc = app.activeDocument;
}

// ============================================================================
// CONFIGURACIÓN: COLORES Y TIPOGRAFÍA
// ============================================================================

var config = {
    colors: {
        navyBlue: {
            name: "Navy Blue",
            r: 26,
            g: 43,
            b: 74,
            c: 100,
            m: 85,
            y: 40,
            k: 35
        },
        gold: {
            name: "Gold",
            r: 232,
            g: 168,
            b: 0,
            c: 0,
            m: 20,
            y: 100,
            k: 0
        },
        white: { r: 255, g: 255, b: 255 },
        darkGray: { r: 58, g: 58, b: 58 }
    },
    fonts: {
        title: "Montserrat Bold",
        body: "Open Sans"
    },
    director: "Gerardo Guzmán Alvarado",
    motto: "No cuidamos instalaciones. Controlamos operaciones.",
    discipline: "Disciplina  ·  Criterio  ·  Control Operativo"
};

// ============================================================================
// UTILIDADES: CONVERSIÓN DE COLORES Y ELEMENTOS
// ============================================================================

function mmToPoints(mm) {
    return mm * 2.834645669;
}

function createCMYKColor(c, m, y, k) {
    var color = new CMYKColor();
    color.cyan = c;
    color.magenta = m;
    color.yellow = y;
    color.black = k;
    return color;
}

function createRGBColor(r, g, b) {
    var color = new RGBColor();
    color.red = r;
    color.green = g;
    color.blue = b;
    return color;
}

function addRectangle(artboard, x, y, width, height, fillColor, strokeColor, strokeWidth) {
    // En Illustrator: rectangle(top, left, width, height)
    // top = posición Y desde arriba, left = posición X desde izquierda
    var rect = artboard.pathItems.rectangle(y + height, x, width, height);
    if (fillColor) {
        rect.filled = true;
        rect.fillColor = fillColor;
    } else {
        rect.filled = false;
    }
    if (strokeColor) {
        rect.stroked = true;
        rect.strokeColor = strokeColor;
        if (strokeWidth) rect.strokeWidth = strokeWidth;
    } else {
        rect.stroked = false;
    }
    return rect;
}

function addText(artboard, text, x, y, fontSize, fontName, fillColor) {
    var textFrame = artboard.textFrames.add();
    textFrame.contents = text;
    textFrame.left = x;
    textFrame.top = y;

    textFrame.textRange.font = app.fonts.getByName(fontName);
    textFrame.textRange.fontSize = fontSize;
    textFrame.textRange.fillColor = fillColor;

    return textFrame;
}

// ============================================================================
// DEFINICIONES DE ARTBOARDS (10 PIEZAS)
// ============================================================================

var artboards = [
    {
        name: "01 - Letterhead (Hoja Membretada)",
        width: mmToPoints(210),
        height: mmToPoints(297),
        description: "A4 Letterhead with header, contact info, and footer"
    },
    {
        name: "02 - Business Card (Tarjeta de Presentación)",
        width: mmToPoints(90),
        height: mmToPoints(50),
        description: "Standard business card format"
    },
    {
        name: "03 - Envelope (Sobre)",
        width: mmToPoints(220),
        height: mmToPoints(110),
        description: "C5 Envelope format"
    },
    {
        name: "04 - Folder (Carpeta)",
        width: mmToPoints(240),
        height: mmToPoints(320),
        description: "Corporate folder design"
    },
    {
        name: "05 - ID Credential (Credencial)",
        width: mmToPoints(85),
        height: mmToPoints(54),
        description: "ID card format with photo area"
    },
    {
        name: "06 - Report Format (Reporte)",
        width: mmToPoints(210),
        height: mmToPoints(297),
        description: "A4 Report cover page"
    },
    {
        name: "07 - Manual Cover (Portada Manual)",
        width: mmToPoints(210),
        height: mmToPoints(297),
        description: "A4 Manual cover design"
    },
    {
        name: "08 - Corporate Seal (Sello)",
        width: mmToPoints(100),
        height: mmToPoints(100),
        description: "Circular corporate seal"
    },
    {
        name: "09 - Signature Block (Bloque de Firma)",
        width: mmToPoints(180),
        height: mmToPoints(80),
        description: "Executive signature block"
    },
    {
        name: "10 - Proposal Cover (Portada Propuesta)",
        width: mmToPoints(210),
        height: mmToPoints(297),
        description: "A4 Proposal cover design"
    }
];

// ============================================================================
// CREAR ARTBOARDS Y DISEÑOS
// ============================================================================

for (var i = 0; i < artboards.length; i++) {
    var artboardDef = artboards[i];
    var artboard = doc.artboards.add([0, 0, artboardDef.width, -artboardDef.height]);
    artboard.name = artboardDef.name;
    doc.artboards.setActiveArtboardIndex(i);

    // Fondo blanco
    var background = addRectangle(
        artboard,
        0, 0,
        artboardDef.width,
        artboardDef.height,
        createRGBColor(255, 255, 255),
        null
    );

    // Implementación según tipo de pieza
    switch (i) {
        case 0: // Letterhead
            designLetterhead(artboard, artboardDef);
            break;
        case 1: // Business Card
            designBusinessCard(artboard, artboardDef);
            break;
        case 2: // Envelope
            designEnvelope(artboard, artboardDef);
            break;
        case 3: // Folder
            designFolder(artboard, artboardDef);
            break;
        case 4: // ID Credential
            designCredential(artboard, artboardDef);
            break;
        case 5: // Report Format
            designReport(artboard, artboardDef);
            break;
        case 6: // Manual Cover
            designManualCover(artboard, artboardDef);
            break;
        case 7: // Corporate Seal
            designSeal(artboard, artboardDef);
            break;
        case 8: // Signature Block
            designSignatureBlock(artboard, artboardDef);
            break;
        case 9: // Proposal Cover
            designProposalCover(artboard, artboardDef);
            break;
    }
}

// ============================================================================
// DISEÑOS ESPECÍFICOS POR PIEZA
// ============================================================================

function designLetterhead(artboard, artboardDef) {
    var navyColor = createCMYKColor(config.colors.navyBlue.c, config.colors.navyBlue.m, config.colors.navyBlue.y, config.colors.navyBlue.k);
    var goldColor = createCMYKColor(config.colors.gold.c, config.colors.gold.m, config.colors.gold.y, config.colors.gold.k);
    var darkGrayColor = createRGBColor(58, 58, 58);
    var whiteColor = createRGBColor(255, 255, 255);

    // Header Block (Navy Blue)
    var headerHeight = mmToPoints(35);
    addRectangle(artboard, 0, 0, artboardDef.width, headerHeight, navyColor, null);

    // Company Name - MICSA
    addText(artboard, "MICSA", mmToPoints(15), mmToPoints(12), 32, config.fonts.title, whiteColor);

    // Subtitle - SAFETY DIVISION
    addText(artboard, "SAFETY DIVISION", mmToPoints(15), mmToPoints(21), 11, config.fonts.body, goldColor);

    // Contact Info (Right aligned)
    var contactX = mmToPoints(185);
    addText(artboard, "www.micsasafety.com.mx", contactX, mmToPoints(10), 9, config.fonts.body, whiteColor);
    addText(artboard, "contacto@micsasafety.com.mx", contactX, mmToPoints(16), 9, config.fonts.body, whiteColor);
    addText(artboard, "+52 (844) 637-0000", contactX, mmToPoints(22), 9, config.fonts.body, whiteColor);

    // Divider line
    addRectangle(artboard, 0, headerHeight, artboardDef.width, mmToPoints(1), goldColor, null);

    // Meta Info Band (Light Gray)
    var metaHeight = mmToPoints(7);
    var metaY = headerHeight + mmToPoints(1);
    addRectangle(artboard, 0, metaY, artboardDef.width, metaHeight, createRGBColor(238, 238, 238), null);
    addText(artboard, "Monclova, Coahuila, México  ·  Fecha: ____________________  ·  Folio: ____________", mmToPoints(15), metaY + mmToPoints(1), 8, config.fonts.body, darkGrayColor);

    // Left margin accent
    addRectangle(artboard, mmToPoints(7), metaY + metaHeight, mmToPoints(1), mmToPoints(240), goldColor, null);

    // Subject label
    addText(artboard, "ASUNTO:", mmToPoints(15), metaY + metaHeight + mmToPoints(8), 8, config.fonts.title, darkGrayColor);

    // Content area (dashed box placeholder)
    var contentY = metaY + metaHeight + mmToPoints(18);
    var contentHeight = mmToPoints(140);
    addRectangle(artboard, mmToPoints(15), contentY, mmToPoints(180), contentHeight, null, createRGBColor(200, 200, 200), 0.5);
    addText(artboard, "[Contenido de la carta]", mmToPoints(18), contentY + mmToPoints(5), 9, config.fonts.body, createRGBColor(180, 180, 180));

    // Signature section
    var signatureY = contentY + contentHeight + mmToPoints(10);
    addRectangle(artboard, mmToPoints(15), signatureY, mmToPoints(80), mmToPoints(1), darkGrayColor, null);
    addText(artboard, "Nombre y Firma Autorizada", mmToPoints(15), signatureY + mmToPoints(2), 7, config.fonts.title, darkGrayColor);
    addText(artboard, "MICSA Safety Division", mmToPoints(15), signatureY + mmToPoints(5), 8, config.fonts.body, darkGrayColor);

    // Position signature
    addRectangle(artboard, mmToPoints(130), signatureY, mmToPoints(65), mmToPoints(1), darkGrayColor, null);
    addText(artboard, "Cargo / Posición", mmToPoints(130), signatureY + mmToPoints(2), 7, config.fonts.title, darkGrayColor);
    addText(artboard, "Gerardo Guzmán Alvarado", mmToPoints(130), signatureY + mmToPoints(5), 8, config.fonts.body, darkGrayColor);

    // Footer Block (Navy Blue)
    var footerY = artboardDef.height - mmToPoints(32);
    addRectangle(artboard, 0, footerY, artboardDef.width, mmToPoints(32), navyColor, null);
    addRectangle(artboard, 0, footerY, artboardDef.width, mmToPoints(1), goldColor, null);

    // Footer text
    addText(artboard, "Monclova, Coahuila, México  ·  RFC: MSD-000000-XXX", mmToPoints(15), footerY + mmToPoints(2), 7, config.fonts.body, whiteColor);
    addText(artboard, config.motto, mmToPoints(15), footerY + mmToPoints(6), 7, config.fonts.title, goldColor);
    addText(artboard, config.discipline, mmToPoints(15), footerY + mmToPoints(10), 6, config.fonts.body, whiteColor);
    addText(artboard, "Pág. 1 / 1", artboardDef.width - mmToPoints(15), footerY + mmToPoints(10), 6, config.fonts.body, whiteColor);
}

function designBusinessCard(artboard, artboardDef) {
    var navyColor = createCMYKColor(config.colors.navyBlue.c, config.colors.navyBlue.m, config.colors.navyBlue.y, config.colors.navyBlue.k);
    var goldColor = createCMYKColor(config.colors.gold.c, config.colors.gold.m, config.colors.gold.y, config.colors.gold.k);
    var whiteColor = createRGBColor(255, 255, 255);
    var darkGrayColor = createRGBColor(58, 58, 58);

    // Navy Blue strip on left
    addRectangle(artboard, 0, 0, mmToPoints(20), artboardDef.height, navyColor, null);

    // Company name
    addText(artboard, "MICSA", mmToPoints(25), mmToPoints(5), 14, config.fonts.title, navyColor);
    addText(artboard, "SAFETY DIVISION", mmToPoints(25), mmToPoints(10), 7, config.fonts.body, goldColor);

    // Contact info
    addText(artboard, "contacto@micsasafety.com.mx", mmToPoints(25), mmToPoints(25), 6, config.fonts.body, darkGrayColor);
    addText(artboard, "+52 (844) 637-0000", mmToPoints(25), mmToPoints(28), 6, config.fonts.body, darkGrayColor);
    addText(artboard, "www.micsasafety.com.mx", mmToPoints(25), mmToPoints(31), 6, config.fonts.body, darkGrayColor);

    // Position
    addText(artboard, "Gerardo Guzmán Alvarado", mmToPoints(25), mmToPoints(38), 7, config.fonts.title, darkGrayColor);
    addText(artboard, "Director de Seguridad", mmToPoints(25), mmToPoints(42), 6, config.fonts.body, goldColor);

    // Motto at bottom
    addText(artboard, "Disciplina · Criterio · Control", mmToPoints(25), mmToPoints(47), 5, config.fonts.body, navyColor);
}

function designEnvelope(artboard, artboardDef) {
    var navyColor = createCMYKColor(config.colors.navyBlue.c, config.colors.navyBlue.m, config.colors.navyBlue.y, config.colors.navyBlue.k);
    var goldColor = createCMYKColor(config.colors.gold.c, config.colors.gold.m, config.colors.gold.y, config.colors.gold.k);
    var whiteColor = createRGBColor(255, 255, 255);

    // Header strip
    addRectangle(artboard, 0, 0, artboardDef.width, mmToPoints(20), navyColor, null);

    // Logo area
    addText(artboard, "MICSA SAFETY DIVISION", mmToPoints(10), mmToPoints(4), 12, config.fonts.title, whiteColor);
    addText(artboard, "www.micsasafety.com.mx", artboardDef.width - mmToPoints(40), mmToPoints(4), 7, config.fonts.body, goldColor);

    // Address area
    var addressY = mmToPoints(30);
    addText(artboard, "[Recipient Name / Nombre del Destinatario]", mmToPoints(20), addressY, 9, config.fonts.body, createRGBColor(58, 58, 58));
    addText(artboard, "[Address / Dirección]", mmToPoints(20), addressY + mmToPoints(5), 9, config.fonts.body, createRGBColor(58, 58, 58));
    addText(artboard, "[City, State, Postal Code / Ciudad, Estado, Código Postal]", mmToPoints(20), addressY + mmToPoints(10), 9, config.fonts.body, createRGBColor(58, 58, 58));

    // Bottom accent
    addRectangle(artboard, 0, artboardDef.height - mmToPoints(3), artboardDef.width, mmToPoints(3), goldColor, null);
}

function designFolder(artboard, artboardDef) {
    var navyColor = createCMYKColor(config.colors.navyBlue.c, config.colors.navyBlue.m, config.colors.navyBlue.y, config.colors.navyBlue.k);
    var goldColor = createCMYKColor(config.colors.gold.c, config.colors.gold.m, config.colors.gold.y, config.colors.gold.k);
    var whiteColor = createRGBColor(255, 255, 255);

    // Large Navy Blue background
    addRectangle(artboard, 0, 0, artboardDef.width, artboardDef.height, navyColor, null);

    // Gold accent lines
    addRectangle(artboard, 0, mmToPoints(30), artboardDef.width, mmToPoints(2), goldColor, null);
    addRectangle(artboard, 0, artboardDef.height - mmToPoints(50), artboardDef.width, mmToPoints(2), goldColor, null);

    // Company info
    addText(artboard, "MICSA SAFETY DIVISION", mmToPoints(20), mmToPoints(10), 18, config.fonts.title, whiteColor);
    addText(artboard, "Sistema Integral de Seguridad Patrimonial e Industrial", mmToPoints(20), mmToPoints(18), 10, config.fonts.body, goldColor);

    // Motto
    addText(artboard, config.motto, mmToPoints(20), artboardDef.height - mmToPoints(45), 12, config.fonts.title, goldColor);

    // Contact at bottom
    addText(artboard, "www.micsasafety.com.mx  |  +52 (844) 637-0000", mmToPoints(20), artboardDef.height - mmToPoints(10), 8, config.fonts.body, whiteColor);
}

function designCredential(artboard, artboardDef) {
    var navyColor = createCMYKColor(config.colors.navyBlue.c, config.colors.navyBlue.m, config.colors.navyBlue.y, config.colors.navyBlue.k);
    var goldColor = createCMYKColor(config.colors.gold.c, config.colors.gold.m, config.colors.gold.y, config.colors.gold.k);
    var whiteColor = createRGBColor(255, 255, 255);
    var darkGrayColor = createRGBColor(58, 58, 58);

    // Photo area placeholder
    addRectangle(artboard, mmToPoints(3), mmToPoints(3), mmToPoints(25), mmToPoints(33), createRGBColor(200, 200, 200), darkGrayColor, 0.5);
    addText(artboard, "[Foto]", mmToPoints(10), mmToPoints(16), 8, config.fonts.body, darkGrayColor);

    // Navy Blue info block
    addRectangle(artboard, mmToPoints(30), mmToPoints(3), mmToPoints(52), mmToPoints(37), navyColor, null);

    // Name
    addText(artboard, "Gerardo Guzmán Alvarado", mmToPoints(32), mmToPoints(8), 9, config.fonts.title, whiteColor);

    // Position
    addText(artboard, "Director de Seguridad", mmToPoints(32), mmToPoints(13), 7, config.fonts.body, goldColor);

    // Employee ID
    addText(artboard, "ID: MICSA-0001", mmToPoints(32), mmToPoints(20), 7, config.fonts.body, whiteColor);

    // Experience
    addText(artboard, "Experiencia: 5 años", mmToPoints(32), mmToPoints(24), 7, config.fonts.body, whiteColor);
    addText(artboard, "Fuerza Civil | Protección Civil", mmToPoints(32), mmToPoints(27), 6, config.fonts.body, whiteColor);

    // Valid dates
    addText(artboard, "Válida hasta: 2027", mmToPoints(32), mmToPoints(33), 6, config.fonts.body, goldColor);
}

function designReport(artboard, artboardDef) {
    var navyColor = createCMYKColor(config.colors.navyBlue.c, config.colors.navyBlue.m, config.colors.navyBlue.y, config.colors.navyBlue.k);
    var goldColor = createCMYKColor(config.colors.gold.c, config.colors.gold.m, config.colors.gold.y, config.colors.gold.k);
    var whiteColor = createRGBColor(255, 255, 255);
    var darkGrayColor = createRGBColor(58, 58, 58);

    // Header
    addRectangle(artboard, 0, 0, artboardDef.width, mmToPoints(50), navyColor, null);
    addText(artboard, "REPORTE DE SEGURIDAD", mmToPoints(20), mmToPoints(15), 24, config.fonts.title, whiteColor);
    addText(artboard, "MICSA Safety Division", mmToPoints(20), mmToPoints(30), 12, config.fonts.body, goldColor);

    // Gold line
    addRectangle(artboard, 0, mmToPoints(50), artboardDef.width, mmToPoints(2), goldColor, null);

    // Report title area
    addText(artboard, "Período: ____________________", mmToPoints(20), mmToPoints(60), 10, config.fonts.body, darkGrayColor);
    addText(artboard, "Sitio: ____________________", mmToPoints(20), mmToPoints(65), 10, config.fonts.body, darkGrayColor);
    addText(artboard, "Clasificación: Confidencial", mmToPoints(20), mmToPoints(70), 9, config.fonts.body, goldColor);

    // Content placeholder
    var contentY = mmToPoints(80);
    addRectangle(artboard, mmToPoints(20), contentY, mmToPoints(170), mmToPoints(160), null, createRGBColor(200, 200, 200), 0.5);
    addText(artboard, "[Contenido del Reporte]", mmToPoints(25), contentY + mmToPoints(80), 10, config.fonts.body, createRGBColor(180, 180, 180));

    // Footer
    var footerY = artboardDef.height - mmToPoints(15);
    addText(artboard, "MICSA Safety Division  |  Confidencial  |  2026", mmToPoints(20), footerY, 8, config.fonts.body, darkGrayColor);
    addText(artboard, "Pág. 1 / 1", artboardDef.width - mmToPoints(20), footerY, 8, config.fonts.body, darkGrayColor);
}

function designManualCover(artboard, artboardDef) {
    var navyColor = createCMYKColor(config.colors.navyBlue.c, config.colors.navyBlue.m, config.colors.navyBlue.y, config.colors.navyBlue.k);
    var goldColor = createCMYKColor(config.colors.gold.c, config.colors.gold.m, config.colors.gold.y, config.colors.gold.k);
    var whiteColor = createRGBColor(255, 255, 255);

    // Full Navy Blue background
    addRectangle(artboard, 0, 0, artboardDef.width, artboardDef.height, navyColor, null);

    // Gold accent band
    addRectangle(artboard, 0, mmToPoints(80), artboardDef.width, mmToPoints(4), goldColor, null);

    // Title
    addText(artboard, "MANUAL OPERATIVO", mmToPoints(20), mmToPoints(40), 28, config.fonts.title, whiteColor);
    addText(artboard, "MICSA SAFETY DIVISION", mmToPoints(20), mmToPoints(55), 14, config.fonts.body, goldColor);

    // Motto
    addText(artboard, config.motto, mmToPoints(20), mmToPoints(90), 12, config.fonts.title, goldColor);

    // Version info
    addText(artboard, "Versión: 1.0  |  Año: 2026", mmToPoints(20), artboardDef.height - mmToPoints(30), 9, config.fonts.body, whiteColor);
    addText(artboard, "Clasificación: Confidencial", mmToPoints(20), artboardDef.height - mmToPoints(25), 9, config.fonts.body, whiteColor);
    addText(artboard, "Aprobado por: Gerardo Guzmán Alvarado", mmToPoints(20), artboardDef.height - mmToPoints(20), 9, config.fonts.body, goldColor);
}

function designSeal(artboard, artboardDef) {
    var navyColor = createCMYKColor(config.colors.navyBlue.c, config.colors.navyBlue.m, config.colors.navyBlue.y, config.colors.navyBlue.k);
    var goldColor = createCMYKColor(config.colors.gold.c, config.colors.gold.m, config.colors.gold.y, config.colors.gold.k);
    var whiteColor = createRGBColor(255, 255, 255);

    // Create circular seal with Navy Blue
    var centerX = artboardDef.width / 2;
    var centerY = artboardDef.height / 2;
    var radius = mmToPoints(40);

    // Outer circle (Navy)
    var circle = artboard.pathItems.ellipse(centerY + radius, centerX - radius, radius * 2, radius * 2);
    circle.filled = true;
    circle.fillColor = navyColor;
    circle.stroked = true;
    circle.strokeColor = goldColor;
    circle.strokeWidth = 2;

    // Inner circle (Gold ring)
    var innerCircle = artboard.pathItems.ellipse(centerY + radius - mmToPoints(3), centerX - radius + mmToPoints(3), radius * 1.8, radius * 1.8);
    innerCircle.filled = false;
    innerCircle.stroked = true;
    innerCircle.strokeColor = goldColor;
    innerCircle.strokeWidth = 1;

    // Center text
    addText(artboard, "MICSA", centerX - mmToPoints(15), centerY - mmToPoints(5), 11, config.fonts.title, whiteColor);
    addText(artboard, "SAFETY", centerX - mmToPoints(15), centerY + mmToPoints(2), 9, config.fonts.body, goldColor);
    addText(artboard, "2026", centerX - mmToPoints(10), centerY + mmToPoints(8), 8, config.fonts.body, whiteColor);
}

function designSignatureBlock(artboard, artboardDef) {
    var navyColor = createCMYKColor(config.colors.navyBlue.c, config.colors.navyBlue.m, config.colors.navyBlue.y, config.colors.navyBlue.k);
    var goldColor = createCMYKColor(config.colors.gold.c, config.colors.gold.m, config.colors.gold.y, config.colors.gold.k);
    var darkGrayColor = createRGBColor(58, 58, 58);
    var whiteColor = createRGBColor(255, 255, 255);

    // Navy Blue banner at top
    addRectangle(artboard, 0, 0, artboardDef.width, mmToPoints(12), navyColor, null);
    addText(artboard, "FIRMA AUTORIZADA / AUTHORIZED SIGNATURE", mmToPoints(5), mmToPoints(2), 9, config.fonts.title, whiteColor);

    // Gold separator
    addRectangle(artboard, 0, mmToPoints(12), artboardDef.width, mmToPoints(1), goldColor, null);

    // Signature line
    addRectangle(artboard, mmToPoints(5), mmToPoints(22), mmToPoints(50), mmToPoints(0.5), darkGrayColor, null);
    addText(artboard, "Firma / Signature", mmToPoints(5), mmToPoints(24), 8, config.fonts.body, darkGrayColor);

    // Name
    addText(artboard, "Nombre: Gerardo Guzmán Alvarado", mmToPoints(5), mmToPoints(30), 8, config.fonts.body, darkGrayColor);

    // Position and date
    addText(artboard, "Cargo: Director de Seguridad", mmToPoints(5), mmToPoints(35), 8, config.fonts.body, darkGrayColor);
    addText(artboard, "Fecha: ____________________", mmToPoints(5), mmToPoints(40), 8, config.fonts.body, darkGrayColor);

    // Seal placeholder
    addRectangle(artboard, mmToPoints(120), mmToPoints(18), mmToPoints(55), mmToPoints(55), createRGBColor(230, 230, 230), darkGrayColor, 0.5);
    addText(artboard, "[Sello / Seal]", mmToPoints(135), mmToPoints(40), 10, config.fonts.body, darkGrayColor);
}

function designProposalCover(artboard, artboardDef) {
    var navyColor = createCMYKColor(config.colors.navyBlue.c, config.colors.navyBlue.m, config.colors.navyBlue.y, config.colors.navyBlue.k);
    var goldColor = createCMYKColor(config.colors.gold.c, config.colors.gold.m, config.colors.gold.y, config.colors.gold.k);
    var whiteColor = createRGBColor(255, 255, 255);
    var darkGrayColor = createRGBColor(58, 58, 58);

    // Left Navy Blue stripe
    addRectangle(artboard, 0, 0, mmToPoints(40), artboardDef.height, navyColor, null);

    // Gold accent
    addRectangle(artboard, mmToPoints(38), 0, mmToPoints(2), artboardDef.height, goldColor, null);

    // Title area
    addText(artboard, "PROPUESTA COMERCIAL", mmToPoints(60), mmToPoints(60), 26, config.fonts.title, navyColor);
    addText(artboard, "SISTEMA INTEGRAL DE SEGURIDAD", mmToPoints(60), mmToPoints(75), 14, config.fonts.body, goldColor);

    // Company info
    addText(artboard, "MICSA SAFETY DIVISION", mmToPoints(60), mmToPoints(95), 12, config.fonts.title, navyColor);

    // Project details placeholder
    var detailsY = mmToPoints(115);
    addText(artboard, "Cliente: ____________________", mmToPoints(60), detailsY, 10, config.fonts.body, darkGrayColor);
    addText(artboard, "Proyecto: ____________________", mmToPoints(60), detailsY + mmToPoints(8), 10, config.fonts.body, darkGrayColor);
    addText(artboard, "Fecha: ____________________", mmToPoints(60), detailsY + mmToPoints(16), 10, config.fonts.body, darkGrayColor);

    // Footer
    var footerY = artboardDef.height - mmToPoints(25);
    addRectangle(artboard, 0, footerY, artboardDef.width, mmToPoints(1), goldColor, null);
    addText(artboard, config.motto, mmToPoints(60), footerY + mmToPoints(3), 11, config.fonts.title, navyColor);
    addText(artboard, "www.micsasafety.com.mx  |  +52 (844) 637-0000", mmToPoints(60), footerY + mmToPoints(10), 8, config.fonts.body, darkGrayColor);
}

// ============================================================================
// GUARDAR DOCUMENTO
// ============================================================================

var targetFolder = Folder.desktop;
var documentName = "MICSA_SAFETY_CORPORATE_IDENTITY_10PIECES.ai";
var filePath = new File(targetFolder + "/" + documentName);

try {
    doc.saveAs(filePath, { compatibility: DocumentAIVersion.Illustrator2025 });
    alert("✓ ÉXITO\n\nDocumento creado con 10 artboards:\n" +
          "1. Letterhead (Hoja Membretada)\n" +
          "2. Business Card (Tarjeta)\n" +
          "3. Envelope (Sobre)\n" +
          "4. Folder (Carpeta)\n" +
          "5. ID Credential (Credencial)\n" +
          "6. Report Format (Reporte)\n" +
          "7. Manual Cover (Manual)\n" +
          "8. Corporate Seal (Sello)\n" +
          "9. Signature Block (Firma)\n" +
          "10. Proposal Cover (Propuesta)\n\n" +
          "Colores: Navy Blue + Gold\n" +
          "Guardado en: " + targetFolder.name);
} catch (e) {
    alert("✗ ERROR\n\nNo se pudo guardar el documento.\n\nDetalle: " + e.message);
}
