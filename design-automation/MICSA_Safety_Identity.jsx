/**
 * MICSA SAFETY DIVISION â€” Identity System Generator
 * ExtendScript (.jsx) para Adobe Illustrator 2025
 *
 * Genera: 13 artboards + exporta PDF por artboard + PNG por artboard
 *
 * USO: Archivo > Scripts > Otro script... > seleccionar este archivo
 *      O doble clic en el archivo con Illustrator abierto
 */

// â”€â”€â”€ CONFIGURACIÃ“N GLOBAL â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

var BASE_PATH  = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\MICSA_Safety\\";
var LOGO_FILE  = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\adobe\\ID-5608-20260329T034218Z-1-001\\ID-5608\\LOGO 7\\editable\\editable.ai";

// Dimensiones originales del logo editable.ai (500Ã—300 pt â†’ 176.4Ã—105.8 mm)
var LOGO_ORI_W = 500;
var LOGO_ORI_H = 300;

// Colores CMYK â€” Paleta MICSA Safety Division (Negro + Plateado + Amarillo)
// Basado en LOGO 7: escudo plateado, texto amarillo, fondo negro
var COLOR = {
    navy:     [60, 40, 40, 100],  // Negro Rico (Rich Black) â€” fondos grandes
    gold:     [ 0, 15,100,  0],   // Amarillo MICSA Safety (del logo real) â€” acentos vivos
    white:    [ 0,  0,  0,  0],   // Blanco
    darkGray: [ 0,  0,  0, 30]    // Plateado/Gris frÃ­o â€” lÃ­neas y detalles plateados
};

// Medidas en puntos (1 mm = 2.834645669 pt)
var MM = 2.834645669;
var PT = {
    a4w:    210 * MM,   // 595.276
    a4h:    297 * MM,   // 841.890
    cr85w:  88.9 * MM,  // 251.906
    cr85h:  50.8 * MM,  // 144.000
    cr80w:  85.6 * MM,  // 242.685
    cr80h:  53.98 * MM, // 153.055
    dlw:    110 * MM,   // 311.811
    dlh:    220 * MM,   // 623.622
    m2:     20  * MM,   // 56.693  (margen 2cm)
    m15:    15  * MM,   // 42.520  (margen 1.5cm)
    m1:     10  * MM,   // 28.346  (margen 1cm)
    m05:    5   * MM,   // 14.173  (margen 0.5cm)
    m03:    3   * MM,   // 8.504   (margen 0.3cm)
    m025:   2.5 * MM,   // 7.087   (margen 0.25cm)
    sello:  30  * MM,   // 85.039  (sello 30mm)
    firmaw: 80  * MM,   // 226.772 (firma 80mm)
    firmah: 25  * MM    // 70.866  (firma 25mm)
};

// TipografÃ­a
var FONT = {
    bold:    "Montserrat-Bold",
    semi:    "Montserrat-SemiBold",
    regular: "OpenSans-Regular",
    light:   "OpenSans-Light"
};

// DefiniciÃ³n de artboards
var ARTBOARDS = [
    { id: 1,  name: "01_Hoja_Membretada",          w: PT.a4w,   h: PT.a4h,   label: "Hoja Membretada A4"       },
    { id: 2,  name: "02_Tarjeta_Frente",            w: PT.cr85w, h: PT.cr85h, label: "Tarjeta Frente CR85"      },
    { id: 3,  name: "02_Tarjeta_Reverso",           w: PT.cr85w, h: PT.cr85h, label: "Tarjeta Reverso CR85"     },
    { id: 4,  name: "03_Sobre_Corporativo",         w: PT.dlw,   h: PT.dlh,   label: "Sobre DL"                 },
    { id: 5,  name: "04_Carpeta_Institucional",     w: PT.a4w,   h: PT.a4h,   label: "Carpeta A4"               },
    { id: 6,  name: "05_Credencial_Empleado_Frente",w: PT.cr80w, h: PT.cr80h, label: "Credencial Frente CR80"   },
    { id: 7,  name: "05_Credencial_Empleado_Reverso",w:PT.cr80w, h: PT.cr80h, label: "Credencial Reverso CR80"  },
    { id: 8,  name: "06_Bitacora_Operativa",        w: PT.a4w,   h: PT.a4h,   label: "BitÃ¡cora A4"              },
    { id: 9,  name: "07_Portada_Manual",            w: PT.a4w,   h: PT.a4h,   label: "Portada Manual A4"        },
    { id: 10, name: "08_Sello_Corporativo",         w: PT.sello, h: PT.sello, label: "Sello 30Ã—30mm"            },
    { id: 11, name: "09_Firma_Institucional",       w: PT.firmaw,h: PT.firmah, label: "Firma 80Ã—25mm"           },
    { id: 12, name: "10_Portada_Propuesta",         w: PT.a4w,   h: PT.a4h,   label: "Portada Propuesta A4"     },
    { id: 13, name: "MOCKUP_Composicion_Final",     w: PT.a4w * 2, h: PT.a4h, label: "Mockup ComposiciÃ³n Final" }
];

// â”€â”€â”€ UTILIDADES â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function makeCMYK(c, m, y, k) {
    var color = new CMYKColor();
    color.cyan    = c;
    color.magenta = m;
    color.yellow  = y;
    color.black   = k;
    return color;
}

function makeRect(layer, x, y, w, h, cmykArr, noStroke) {
    var rect = layer.pathItems.rectangle(y, x, w, h);
    rect.filled = true;
    rect.fillColor = makeCMYK(cmykArr[0], cmykArr[1], cmykArr[2], cmykArr[3]);
    if (noStroke) {
        rect.stroked = false;
    }
    return rect;
}

function makeLine(layer, x1, y1, x2, y2, cmykArr, weight) {
    var line = layer.pathItems.add();
    line.setEntirePath([[x1, -y1], [x2, -y2]]);
    line.stroked = true;
    line.strokeColor = makeCMYK(cmykArr[0], cmykArr[1], cmykArr[2], cmykArr[3]);
    line.strokeWidth = weight || 1;
    line.filled = false;
    return line;
}

function makeText(layer, content, x, y, size, fontName, cmykArr, width) {
    var tf;
    if (width) {
        tf = layer.textFrames.areaText(
            layer.pathItems.rectangle(y, x, width, size * 1.4)
        );
    } else {
        tf = layer.textFrames.pointText([x, -y]);
    }
    tf.contents = content;
    var attr = tf.textRange.characterAttributes;
    try { attr.textFont = app.textFonts.getByName(fontName); } catch(e) {
        try { attr.textFont = app.textFonts.getByName("Arial-BoldMT"); } catch(e2) {}
    }
    attr.size = size;
    attr.fillColor = makeCMYK(cmykArr[0], cmykArr[1], cmykArr[2], cmykArr[3]);
    return tf;
}

function makeCircle(layer, cx, cy, r, strokeArr, fillArr, strokeW) {
    var circle = layer.pathItems.ellipse(cy - r, cx - r, r * 2, r * 2);
    if (fillArr) {
        circle.filled = true;
        circle.fillColor = makeCMYK(fillArr[0], fillArr[1], fillArr[2], fillArr[3]);
    } else {
        circle.filled = false;
    }
    if (strokeArr) {
        circle.stroked = true;
        circle.strokeColor = makeCMYK(strokeArr[0], strokeArr[1], strokeArr[2], strokeArr[3]);
        circle.strokeWidth = strokeW || 2;
    } else {
        circle.stroked = false;
    }
    return circle;
}

// Coordenada Y en Illustrator: Y=0 es TOP del artboard, crece hacia abajo
// makeRect(layer, x, y, w, h) => x=izq, y=top(negativo en AI coords)
// La API de Illustrator usa: rectangle(top, left, width, height) donde top es negativo

function makeRectAI(layer, left, top, w, h, cmykArr) {
    // top, left en coordenadas visuales (0,0 = esquina superior izq del artboard)
    var rect = layer.pathItems.rectangle(-top, left, w, h);
    rect.filled = true;
    rect.fillColor = makeCMYK(cmykArr[0], cmykArr[1], cmykArr[2], cmykArr[3]);
    rect.stroked = false;
    return rect;
}

function makeTextAI(layer, content, left, top, size, fontName, cmykArr) {
    var tf = layer.textFrames.pointText([left, -top]);
    tf.contents = content;
    var attr = tf.textRange.characterAttributes;
    try { attr.textFont = app.textFonts.getByName(fontName); } catch(e) {}
    attr.size = size;
    attr.fillColor = makeCMYK(cmykArr[0], cmykArr[1], cmykArr[2], cmykArr[3]);
    return tf;
}

// Dibujar lÃ­nea separadora dorada horizontal
function goldLine(layer, left, top, width, weight) {
    var path = layer.pathItems.add();
    path.setEntirePath([[left, -top], [left + width, -top]]);
    path.stroked = true;
    path.strokeColor = makeCMYK(COLOR.gold[0], COLOR.gold[1], COLOR.gold[2], COLOR.gold[3]);
    path.strokeWidth = weight || 1.5;
    path.filled = false;
    return path;
}

// â”€â”€â”€ COLOCAR LOGO (editable.ai) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
/**
 * Coloca el logo editable.ai en la capa "Logo" del artboard activo.
 *
 * @param {Layer}  layer      - capa destino (L.Logo)
 * @param {number} left       - posiciÃ³n X desde borde izq del artboard (pt)
 * @param {number} top        - posiciÃ³n Y desde borde sup del artboard (pt)
 * @param {number} targetW    - ancho deseado en pt (alto se calcula por proporciÃ³n)
 * @param {string} [version]  - "color" (default) | "white"
 *                              Si "white" se intenta invertir colores vÃ­a opacity trick;
 *                              para resultado exacto reemplazar con logo blanco real.
 */
function placeLogo(layer, left, top, targetW, version) {
    var logoFile = new File(LOGO_FILE);
    if (!logoFile.exists) {
        // Si el archivo no existe, dibujar placeholder
        var ph = layer.pathItems.rectangle(-top, left, targetW, targetW * (LOGO_ORI_H / LOGO_ORI_W));
        ph.filled = false;
        ph.stroked = true;
        ph.strokeColor = makeCMYK(COLOR.gold[0], COLOR.gold[1], COLOR.gold[2], COLOR.gold[3]);
        ph.strokeWidth = 0.5;
        return ph;
    }

    var placed = layer.placedItems.add();
    placed.file = logoFile;

    // Calcular escala proporcional
    var scale = (targetW / LOGO_ORI_W) * 100;  // porcentaje

    // Reposicionar: en AI el origen del placed item es su esquina inferior-izq
    // Debemos colocarlo en (left, -top) en coordenadas del artboard
    placed.left   = left;
    placed.top    = -top;  // negativo porque AI usa Y invertida

    // Escalar manteniendo proporciones
    placed.resize(scale, scale);

    // VersiÃ³n blanca: reducir opacity o aplicar modo Multiply sobre fondo oscuro
    if (version === "white") {
        // Nota: para logo blanco real, sustituir LOGO_FILE por versiÃ³n blanca
        // Como workaround, usar opacity + blend mode Screen sobre fondo navy
        placed.opacity = 100;
        placed.blendingMode = BlendModes.SCREEN;
    }

    return placed;
}

// â”€â”€â”€ CREACIÃ“N DE CAPAS ESTÃNDAR â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function createLayers(doc) {
    // Illustrator crea capas en orden inverso, la Ãºltima queda arriba
    var layers = {};
    var names = ["Fondo", "Logo", "Elementos_Graficos", "Tipografia", "Guias", "Notas"];
    for (var i = 0; i < names.length; i++) {
        try {
            layers[names[i]] = doc.layers.getByName(names[i]);
        } catch(e) {
            layers[names[i]] = doc.layers.add();
            layers[names[i]].name = names[i];
        }
    }
    return layers;
}

// â”€â”€â”€ ARTBOARD 1: HOJA MEMBRETADA â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildHojaMembretada(doc, ab) {
    var W = PT.a4w, H = PT.a4h;
    var m = PT.m2;
    var L = createLayers(doc);

    // Header navy 2.5cm
    makeRectAI(L.Fondo, 0, 0, W, 25 * MM, COLOR.navy);
    // Footer navy 1cm
    makeRectAI(L.Fondo, 0, H - 10 * MM, W, 10 * MM, COLOR.navy);

    // LÃ­nea dorada bajo header
    goldLine(L.Elementos_Graficos, 0, 25 * MM, W, 2);
    // LÃ­nea dorada sobre footer
    goldLine(L.Elementos_Graficos, 0, H - 10 * MM, W, 2);

    // Logo en header (ancho 55mm, modo white sobre navy)
    placeLogo(L.Logo, m, 3 * MM, 55 * MM, "white");

    // Textos del header
    makeTextAI(L.Tipografia, "MICSA SAFETY DIVISION", m + 57*MM, 7 * MM, 14, FONT.bold, COLOR.gold);
    makeTextAI(L.Tipografia, "Seguridad Industrial y Operativa", m + 57*MM, 15 * MM, 8, FONT.regular, COLOR.white);

    // Ãrea de contenido - guÃ­a de mÃ¡rgenes (caja invisible)
    var contentBox = L.Guias.pathItems.rectangle(-(25 * MM + 5 * MM), m, W - m * 2, H - 35 * MM - m);
    contentBox.filled = false;
    contentBox.stroked = true;
    contentBox.strokeColor = makeCMYK(0, 0, 0, 20);
    contentBox.strokeWidth = 0.5;
    contentBox.strokeDashes = [4, 4];

    // Folio / nÃºmero de pÃ¡gina (placeholder)
    makeTextAI(L.Tipografia, "Folio: ___________", m, H - 7 * MM, 7, FONT.light, COLOR.white);
    makeTextAI(L.Tipografia, "PÃ¡gina 1 de 1", W - m - 50 * MM, H - 7 * MM, 7, FONT.light, COLOR.white);

    // LÃ­nea de firma al pie del contenido
    goldLine(L.Elementos_Graficos, m, H - 40 * MM, 60 * MM, 0.75);
    makeTextAI(L.Tipografia, "Firma Autorizada", m, H - 37 * MM, 7, FONT.light, COLOR.darkGray);
}

// â”€â”€â”€ ARTBOARD 2: TARJETA FRENTE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildTarjetaFrente(doc, ab) {
    var W = PT.cr85w, H = PT.cr85h;
    var m = PT.m03;
    var L = createLayers(doc);

    // Fondo navy puro
    makeRectAI(L.Fondo, 0, 0, W, H, COLOR.navy);

    // Banda dorada izquierda (4mm)
    makeRectAI(L.Fondo, 0, 0, 4 * MM, H, COLOR.gold);

    // Logo en tarjeta frente (ancho 30mm, modo white sobre navy)
    placeLogo(L.Logo, 10 * MM, 4 * MM, 30 * MM, "white");

    // Nombre empresa (queda debajo del logo)
    makeTextAI(L.Tipografia, "MICSA", 10 * MM, 22 * MM, 10, FONT.bold, COLOR.white);
    makeTextAI(L.Tipografia, "SAFETY DIVISION", 10 * MM, 29 * MM, 6, FONT.semi, COLOR.gold);

    // LÃ­nea separadora
    goldLine(L.Elementos_Graficos, 10 * MM, 24 * MM, W - 14 * MM, 0.5);

    // Datos del titular (placeholders)
    makeTextAI(L.Tipografia, "Nombre del Titular", 10 * MM, 30 * MM, 9, FONT.semi, COLOR.white);
    makeTextAI(L.Tipografia, "Cargo / PosiciÃ³n", 10 * MM, 37 * MM, 7, FONT.regular, COLOR.gold);

    // Datos de contacto
    makeTextAI(L.Tipografia, "contacto@micsa.mx", 10 * MM, 43 * MM, 6, FONT.light, COLOR.white);
    makeTextAI(L.Tipografia, "www.micsa.mx", 10 * MM, 48 * MM, 6, FONT.light, COLOR.white);
}

// â”€â”€â”€ ARTBOARD 3: TARJETA REVERSO â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildTarjetaReverso(doc, ab) {
    var W = PT.cr85w, H = PT.cr85h;
    var m = PT.m03;
    var L = createLayers(doc);

    // Fondo blanco
    makeRectAI(L.Fondo, 0, 0, W, H, COLOR.white);

    // LÃ­nea dorada superior
    goldLine(L.Elementos_Graficos, 0, 3 * MM, W, 1.5);
    // LÃ­nea dorada inferior
    goldLine(L.Elementos_Graficos, 0, H - 3 * MM, W, 1.5);

    // Ãrea QR placeholder (cuadro)
    var qrBox = L.Elementos_Graficos.pathItems.rectangle(-(H/2 - 15*MM), W/2 - 15*MM, 30*MM, 30*MM);
    qrBox.filled = false;
    qrBox.stroked = true;
    qrBox.strokeColor = makeCMYK(COLOR.darkGray[0],COLOR.darkGray[1],COLOR.darkGray[2],COLOR.darkGray[3]);
    qrBox.strokeWidth = 0.5;
    makeTextAI(L.Tipografia, "QR", W/2 - 3*MM, H/2 + 2*MM, 7, FONT.light, COLOR.darkGray);

    makeTextAI(L.Tipografia, "Seguridad Industrial", m, H - 8 * MM, 6, FONT.light, COLOR.darkGray);
    makeTextAI(L.Tipografia, "www.micsa.mx", W - 40*MM, H - 8 * MM, 6, FONT.regular, COLOR.navy);
}

// â”€â”€â”€ ARTBOARD 4: SOBRE DL â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildSobre(doc, ab) {
    var W = PT.dlw, H = PT.dlh;
    var L = createLayers(doc);

    // Panel navy izquierdo (4cm)
    makeRectAI(L.Fondo, 0, 0, 40 * MM, H, COLOR.navy);
    // Fondo blanco derecho
    makeRectAI(L.Fondo, 40 * MM, 0, W - 40 * MM, H, COLOR.white);

    // LÃ­nea dorada vertical separadora
    var vline = L.Elementos_Graficos.pathItems.add();
    vline.setEntirePath([[40 * MM, 0], [40 * MM, -H]]);
    vline.stroked = true;
    vline.strokeColor = makeCMYK(COLOR.gold[0],COLOR.gold[1],COLOR.gold[2],COLOR.gold[3]);
    vline.strokeWidth = 1.5;
    vline.filled = false;

    // Logo en panel navy del sobre (ancho 30mm)
    placeLogo(L.Logo, 5 * MM, 15 * MM, 30 * MM, "white");
    makeTextAI(L.Tipografia, "www.micsa.mx", 8 * MM, 25 * MM + 30*MM*(LOGO_ORI_H/LOGO_ORI_W), 6, FONT.light, COLOR.white);

    // Ãrea destinatario
    makeTextAI(L.Tipografia, "Para:", 50 * MM, 30 * MM, 7, FONT.light, COLOR.darkGray);
    makeTextAI(L.Tipografia, "Nombre del Destinatario", 50 * MM, 38 * MM, 10, FONT.semi, COLOR.navy);
    makeTextAI(L.Tipografia, "Cargo / Empresa", 50 * MM, 47 * MM, 8, FONT.regular, COLOR.darkGray);
    makeTextAI(L.Tipografia, "DirecciÃ³n LÃ­nea 1", 50 * MM, 55 * MM, 8, FONT.regular, COLOR.darkGray);
    makeTextAI(L.Tipografia, "Ciudad, Estado CP", 50 * MM, 63 * MM, 8, FONT.regular, COLOR.darkGray);
}

// â”€â”€â”€ ARTBOARD 5: CARPETA INSTITUCIONAL â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildCarpeta(doc, ab) {
    var W = PT.a4w, H = PT.a4h;
    var L = createLayers(doc);

    // Fondo blanco
    makeRectAI(L.Fondo, 0, 0, W, H, COLOR.white);

    // Diagonal navy (triÃ¡ngulo inferior-derecho ~45Â°)
    var diag = L.Fondo.pathItems.add();
    diag.setEntirePath([
        [0, -H],
        [W, -H],
        [W, -(H - W)]
    ]);
    diag.closed = true;
    diag.filled = true;
    diag.fillColor = makeCMYK(COLOR.navy[0],COLOR.navy[1],COLOR.navy[2],COLOR.navy[3]);
    diag.stroked = false;

    // LÃ­nea dorada diagonal
    var diagLine = L.Elementos_Graficos.pathItems.add();
    diagLine.setEntirePath([[0, -H], [W, -(H - W)]]);
    diagLine.stroked = true;
    diagLine.strokeColor = makeCMYK(COLOR.gold[0],COLOR.gold[1],COLOR.gold[2],COLOR.gold[3]);
    diagLine.strokeWidth = 2;
    diagLine.filled = false;

    // Header
    makeRectAI(L.Fondo, 0, 0, W, 20 * MM, COLOR.navy);
    goldLine(L.Elementos_Graficos, 0, 20 * MM, W, 2);
    // Logo carpeta en header
    placeLogo(L.Logo, PT.m05, 2 * MM, 45 * MM, "white");
    makeTextAI(L.Tipografia, "MICSA SAFETY DIVISION", PT.m05 + 47*MM, 7 * MM, 12, FONT.bold, COLOR.white);
    makeTextAI(L.Tipografia, "Seguridad Industrial y Operativa", PT.m05 + 47*MM, 15 * MM, 6, FONT.regular, COLOR.gold);

    // Texto en Ã¡rea navy diagonal
    makeTextAI(L.Tipografia, "micsa.mx", W - 60*MM, H - 15*MM, 10, FONT.light, COLOR.white);
}

// â”€â”€â”€ ARTBOARD 6: CREDENCIAL FRENTE â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildCredencialFrente(doc, ab) {
    var W = PT.cr80w, H = PT.cr80h;
    var m = PT.m025;
    var L = createLayers(doc);

    // Fondo blanco
    makeRectAI(L.Fondo, 0, 0, W, H, COLOR.white);

    // Barra navy superior (12mm)
    makeRectAI(L.Fondo, 0, 0, W, 12 * MM, COLOR.navy);
    goldLine(L.Elementos_Graficos, 0, 12 * MM, W, 1.5);

    // Logo credencial (ancho 22mm en el header de 12mm de alto)
    placeLogo(L.Logo, m, 1.5 * MM, 22 * MM, "white");
    makeTextAI(L.Tipografia, "MICSA SAFETY", m + 24*MM, 5 * MM, 6, FONT.bold, COLOR.white);
    makeTextAI(L.Tipografia, "CREDENCIAL DE IDENTIFICACIÃ“N", m + 24*MM, 10 * MM, 4, FONT.light, COLOR.gold);

    // Placeholder foto (rectÃ¡ngulo)
    makeRectAI(L.Elementos_Graficos, m, 14 * MM, 22 * MM, 28 * MM, COLOR.darkGray);
    makeTextAI(L.Tipografia, "FOTO", m + 6*MM, 29 * MM, 7, FONT.light, COLOR.white);

    // Datos del empleado
    makeTextAI(L.Tipografia, "NOMBRE COMPLETO", 28 * MM, 18 * MM, 8, FONT.bold, COLOR.navy);
    makeTextAI(L.Tipografia, "Cargo / Departamento", 28 * MM, 25 * MM, 6, FONT.regular, COLOR.darkGray);

    goldLine(L.Elementos_Graficos, 28 * MM, 28 * MM, W - m - 28*MM, 0.5);

    makeTextAI(L.Tipografia, "ID: MSI-0000", 28 * MM, 32 * MM, 6, FONT.semi, COLOR.navy);
    makeTextAI(L.Tipografia, "Vigencia: 2026", 28 * MM, 37 * MM, 6, FONT.light, COLOR.darkGray);

    // Barra navy inferior (6mm)
    makeRectAI(L.Fondo, 0, H - 6 * MM, W, 6 * MM, COLOR.navy);
    makeTextAI(L.Tipografia, "www.micsa.mx", m, H - 3 * MM, 5, FONT.light, COLOR.gold);
}

// â”€â”€â”€ ARTBOARD 7: CREDENCIAL REVERSO â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildCredencialReverso(doc, ab) {
    var W = PT.cr80w, H = PT.cr80h;
    var m = PT.m025;
    var L = createLayers(doc);

    // Fondo dorado
    makeRectAI(L.Fondo, 0, 0, W, H, COLOR.gold);

    // Header navy
    makeRectAI(L.Fondo, 0, 0, W, 10 * MM, COLOR.navy);
    // Logo credencial reverso (sobre header navy 10mm)
    placeLogo(L.Logo, m, 0.5 * MM, 20 * MM, "white");
    makeTextAI(L.Tipografia, "MICSA SAFETY DIVISION", m + 22*MM, 4 * MM, 5.5, FONT.bold, COLOR.white);
    makeTextAI(L.Tipografia, "Datos de Emergencia", m + 22*MM, 8.5 * MM, 4.5, FONT.light, COLOR.gold);

    // Contenido reverso
    makeTextAI(L.Tipografia, "En caso de accidente notificar a:", m, 14 * MM, 6, FONT.semi, COLOR.navy);
    makeTextAI(L.Tipografia, "Tel. Emergencias: 800-MICSA-01", m, 20 * MM, 6, FONT.regular, COLOR.navy);
    makeTextAI(L.Tipografia, "Tipo de Sangre: ________", m, 26 * MM, 6, FONT.regular, COLOR.navy);
    makeTextAI(L.Tipografia, "Alergias: ______________", m, 32 * MM, 6, FONT.regular, COLOR.navy);
    makeTextAI(L.Tipografia, "Contacto: ______________", m, 38 * MM, 6, FONT.regular, COLOR.navy);

    goldLine(L.Elementos_Graficos, m, 42 * MM, W - m * 2, 0.5);

    // CÃ³digo de barras placeholder
    makeRectAI(L.Elementos_Graficos, W/2 - 20*MM, 44 * MM, 40 * MM, 8 * MM, COLOR.navy);
    makeTextAI(L.Tipografia, "MSI 0000 00000", W/2 - 15*MM, 53 * MM, 5, FONT.light, COLOR.navy);
}

// â”€â”€â”€ ARTBOARD 8: BITÃCORA OPERATIVA â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildBitacora(doc, ab) {
    var W = PT.a4w, H = PT.a4h;
    var m = PT.m15;
    var L = createLayers(doc);

    // Header navy
    makeRectAI(L.Fondo, 0, 0, W, 22 * MM, COLOR.navy);
    goldLine(L.Elementos_Graficos, 0, 22 * MM, W, 2);
    placeLogo(L.Logo, m, 2 * MM, 40 * MM, "white");
    makeTextAI(L.Tipografia, "BITÃCORA OPERATIVA", m + 42*MM, 8 * MM, 14, FONT.bold, COLOR.white);
    makeTextAI(L.Tipografia, "MICSA Safety Division", m + 42*MM, 17 * MM, 7, FONT.light, COLOR.gold);

    // Info cabecera
    var infoY = 28 * MM;
    makeTextAI(L.Tipografia, "Fecha: _________________", m, infoY, 8, FONT.regular, COLOR.navy);
    makeTextAI(L.Tipografia, "Turno: _________________", W/2, infoY, 8, FONT.regular, COLOR.navy);
    goldLine(L.Elementos_Graficos, m, infoY + 5*MM, W - m*2, 0.5);

    // Tabla de 18 filas
    var tableTop = 38 * MM;
    var rowH = 11 * MM;
    var cols = [m, m + 15*MM, m + 55*MM, m + 95*MM, m + 135*MM, W - m];
    var headers = ["No.", "Hora", "Actividad / Evento", "Responsable", "Observaciones"];

    // Header de tabla
    makeRectAI(L.Fondo, m, tableTop, W - m*2, rowH, COLOR.navy);
    for (var c = 0; c < headers.length; c++) {
        makeTextAI(L.Tipografia, headers[c], cols[c] + 1*MM, tableTop + 4*MM, 6, FONT.semi, COLOR.white);
    }
    goldLine(L.Elementos_Graficos, m, tableTop + rowH, W - m*2, 1);

    // 18 filas de datos
    for (var r = 0; r < 18; r++) {
        var rowTop = tableTop + rowH + r * rowH;
        var bg = (r % 2 === 1) ? COLOR.darkGray : COLOR.white;
        // Fondo alternado suave
        if (r % 2 === 1) {
            makeRectAI(L.Fondo, m, rowTop, W - m*2, rowH, [28, 20, 18, 10]);
        }
        // LÃ­nea inferior de fila
        goldLine(L.Elementos_Graficos, m, rowTop + rowH, W - m*2, 0.25);
        // NÃºmero de fila
        makeTextAI(L.Tipografia, String(r + 1), cols[0] + 1*MM, rowTop + 4*MM, 6, FONT.light, COLOR.navy);
        // LÃ­neas verticales entre columnas
        for (var cv = 1; cv < cols.length - 1; cv++) {
            var vl = L.Elementos_Graficos.pathItems.add();
            vl.setEntirePath([[cols[cv], -rowTop], [cols[cv], -(rowTop + rowH)]]);
            vl.stroked = true;
            vl.strokeColor = makeCMYK(COLOR.gold[0],COLOR.gold[1],COLOR.gold[2],COLOR.gold[3]);
            vl.strokeWidth = 0.25;
            vl.filled = false;
        }
    }

    // Footer
    makeRectAI(L.Fondo, 0, H - 10*MM, W, 10*MM, COLOR.navy);
    makeTextAI(L.Tipografia, "Supervisado por: _______________________", m, H - 6*MM, 7, FONT.regular, COLOR.white);
    makeTextAI(L.Tipografia, "Firma: ___________", W - m - 60*MM, H - 6*MM, 7, FONT.regular, COLOR.gold);
}

// â”€â”€â”€ ARTBOARD 9: PORTADA MANUAL â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildPortadaManual(doc, ab) {
    var W = PT.a4w, H = PT.a4h;
    var m = PT.m2;
    var L = createLayers(doc);

    // 60% inferior navy
    makeRectAI(L.Fondo, 0, 0, W, H * 0.4, COLOR.white);
    makeRectAI(L.Fondo, 0, H * 0.4, W, H * 0.6, COLOR.navy);

    // LÃ­nea dorada en la uniÃ³n
    goldLine(L.Elementos_Graficos, 0, H * 0.4, W, 3);

    // TÃ­tulo
    makeTextAI(L.Tipografia, "MANUAL DE", m, H * 0.35 - 30*MM, 28, FONT.bold, COLOR.navy);
    makeTextAI(L.Tipografia, "SEGURIDAD", m, H * 0.35 - 16*MM, 36, FONT.bold, COLOR.navy);
    makeTextAI(L.Tipografia, "OPERATIVA", m, H * 0.35 - 0*MM, 36, FONT.bold, COLOR.gold);

    // Logo portada manual (en Ã¡rea blanca, encima del tÃ­tulo)
    placeLogo(L.Logo, m, H * 0.4 - 22*MM, 70 * MM, "color");

    // SubtÃ­tulo en Ã¡rea navy
    makeTextAI(L.Tipografia, "MICSA Safety Division", m, H * 0.4 + 15*MM, 14, FONT.semi, COLOR.white);
    makeTextAI(L.Tipografia, "EdiciÃ³n 2026 â€” VersiÃ³n 1.0", m, H * 0.4 + 26*MM, 9, FONT.light, COLOR.gold);

    goldLine(L.Elementos_Graficos, m, H * 0.4 + 32*MM, 80*MM, 1);

    makeTextAI(L.Tipografia, "Uso Interno â€” Confidencial", m, H - 15*MM, 7, FONT.light, COLOR.white);
    makeTextAI(L.Tipografia, "Rev. 01 | 2026-03-29", W - m - 60*MM, H - 15*MM, 7, FONT.light, COLOR.gold);
}

// â”€â”€â”€ ARTBOARD 10: SELLO CORPORATIVO â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildSello(doc, ab) {
    var W = PT.sello, H = PT.sello;
    var cx = W / 2, cy = H / 2;
    var r = W / 2 - 1*MM;
    var L = createLayers(doc);

    // CÃ­rculo exterior (borde dorado)
    makeCircle(L.Elementos_Graficos, cx, cy, r, COLOR.gold, null, 2);
    // CÃ­rculo interior (borde dorado delgado)
    makeCircle(L.Elementos_Graficos, cx, cy, r - 3*MM, COLOR.gold, null, 0.5);
    // Fondo interior navy
    makeCircle(L.Fondo, cx, cy, r - 1.5*MM, null, COLOR.navy, 0);

    // Texto central
    makeTextAI(L.Tipografia, "MICSA", cx - 8*MM, cy - 2*MM, 8, FONT.bold, COLOR.white);
    makeTextAI(L.Tipografia, "SAFETY", cx - 9*MM, cy + 5*MM, 7, FONT.semi, COLOR.gold);

    // Texto circular superior (placeholder texto arco)
    makeTextAI(L.Tipografia, "â€¢ SEGURIDAD INDUSTRIAL â€¢", 1.5*MM, cy - r + 5*MM, 4, FONT.regular, COLOR.gold);
    makeTextAI(L.Tipografia, "â€¢ MEXICO â€¢ 2026 â€¢", 3*MM, cy + r - 6*MM, 4, FONT.regular, COLOR.gold);
}

// â”€â”€â”€ ARTBOARD 11: FIRMA INSTITUCIONAL â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildFirma(doc, ab) {
    var W = PT.firmaw, H = PT.firmah;
    var L = createLayers(doc);

    // Fondo blanco
    makeRectAI(L.Fondo, 0, 0, W, H, COLOR.white);

    // Logo real desde editable.ai (ocupa todo el alto de la firma 25mm â†’ ancho proporcional ~42mm)
    var logoFirmaW = H * (LOGO_ORI_W / LOGO_ORI_H);  // proporcional al alto
    placeLogo(L.Logo, 0, 0, logoFirmaW, "color");

    // LÃ­nea dorada separadora
    var logoFW = H * (LOGO_ORI_W / LOGO_ORI_H);  // ancho real del logo
    var vSep = L.Elementos_Graficos.pathItems.add();
    vSep.setEntirePath([[logoFW + 3*MM, -2*MM], [logoFW + 3*MM, -(H - 2*MM)]]);
    vSep.stroked = true;
    vSep.strokeColor = makeCMYK(COLOR.gold[0],COLOR.gold[1],COLOR.gold[2],COLOR.gold[3]);
    vSep.strokeWidth = 1;
    vSep.filled = false;

    // Texto firma
    makeTextAI(L.Tipografia, "MICSA Safety Division", logoFW + 6*MM, 7*MM, 8, FONT.bold, COLOR.navy);
    makeTextAI(L.Tipografia, "Seguridad Industrial y Operativa", logoFW + 6*MM, 13*MM, 6, FONT.regular, COLOR.darkGray);
    makeTextAI(L.Tipografia, "contacto@micsa.mx  |  www.micsa.mx", logoFW + 6*MM, 19*MM, 5.5, FONT.light, COLOR.gold);
    makeTextAI(L.Tipografia, "Tel. 800-MICSA-01", logoFW + 6*MM, 23.5*MM, 5.5, FONT.light, COLOR.darkGray);
}

// â”€â”€â”€ ARTBOARD 12: PORTADA PROPUESTA â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildPortadaPropuesta(doc, ab) {
    var W = PT.a4w, H = PT.a4h;
    var m = PT.m2;
    var L = createLayers(doc);

    // Fondo blanco
    makeRectAI(L.Fondo, 0, 0, W, H, COLOR.white);

    // Diagonal navy (esquina inferior derecha)
    var diagPath = L.Fondo.pathItems.add();
    diagPath.setEntirePath([
        [W * 0.45, -H],
        [W, -H],
        [W, -(H * 0.55)]
    ]);
    diagPath.closed = true;
    diagPath.filled = true;
    diagPath.fillColor = makeCMYK(COLOR.navy[0],COLOR.navy[1],COLOR.navy[2],COLOR.navy[3]);
    diagPath.stroked = false;

    // LÃ­nea dorada diagonal
    var dLine = L.Elementos_Graficos.pathItems.add();
    dLine.setEntirePath([[W * 0.45, -H], [W, -(H * 0.55)]]);
    dLine.stroked = true;
    dLine.strokeColor = makeCMYK(COLOR.gold[0],COLOR.gold[1],COLOR.gold[2],COLOR.gold[3]);
    dLine.strokeWidth = 2.5;
    dLine.filled = false;

    // Header
    makeRectAI(L.Fondo, 0, 0, W, 18*MM, COLOR.navy);
    goldLine(L.Elementos_Graficos, 0, 18*MM, W, 2);
    // Logo en header propuesta
    placeLogo(L.Logo, m, 1.5*MM, 38*MM, "white");
    makeTextAI(L.Tipografia, "MICSA SAFETY DIVISION", m + 40*MM, 7*MM, 11, FONT.bold, COLOR.white);
    makeTextAI(L.Tipografia, "Seguridad Industrial y Operativa", m + 40*MM, 14*MM, 5.5, FONT.light, COLOR.gold);

    // TÃ­tulo propuesta
    makeTextAI(L.Tipografia, "PROPUESTA", m, 50*MM, 36, FONT.bold, COLOR.navy);
    makeTextAI(L.Tipografia, "TÃ‰CNICO-COMERCIAL", m, 70*MM, 22, FONT.semi, COLOR.darkGray);
    goldLine(L.Elementos_Graficos, m, 78*MM, 100*MM, 2);

    makeTextAI(L.Tipografia, "Para: Cliente / Empresa", m, 88*MM, 11, FONT.regular, COLOR.navy);
    makeTextAI(L.Tipografia, "Ref: PROP-2026-001", m, 98*MM, 9, FONT.light, COLOR.darkGray);
    makeTextAI(L.Tipografia, "Fecha: 2026-03-29", m, 107*MM, 9, FONT.light, COLOR.darkGray);

    // Confidencial
    makeTextAI(L.Tipografia, "CONFIDENCIAL", m, H - 20*MM, 8, FONT.semi, COLOR.navy);
    makeTextAI(L.Tipografia, "Este documento es propiedad de MICSA Safety Division", m, H - 14*MM, 6, FONT.light, COLOR.darkGray);
}

// â”€â”€â”€ ARTBOARD 13: MOCKUP COMPOSICIÃ“N FINAL â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildMockup(doc, ab) {
    var W = PT.a4w * 2, H = PT.a4h;
    var L = createLayers(doc);

    // Fondo gris oscuro (mesa de trabajo)
    makeRectAI(L.Fondo, 0, 0, W, H, [0, 0, 0, 85]);

    // TÃ­tulo
    makeTextAI(L.Tipografia, "MICSA SAFETY DIVISION â€” Sistema de Identidad Visual v1.0",
               20*MM, 15*MM, 10, FONT.bold, COLOR.gold);
    goldLine(L.Elementos_Graficos, 20*MM, 20*MM, W - 40*MM, 0.5);

    // Miniaturas de cada pieza (placeholder rectangles en grid)
    var piezas = [
        {n:"Hoja Membretada", x:20*MM, y:25*MM, w:40*MM, h:56*MM},
        {n:"Tarjeta Frente",  x:70*MM, y:25*MM, w:40*MM, h:22*MM},
        {n:"Tarjeta Reverso", x:70*MM, y:52*MM, w:40*MM, h:22*MM},
        {n:"Sobre DL",        x:120*MM,y:25*MM, w:30*MM, h:60*MM},
        {n:"Carpeta",         x:160*MM,y:25*MM, w:40*MM, h:56*MM},
        {n:"Credencial F",    x:210*MM,y:25*MM, w:35*MM, h:22*MM},
        {n:"Credencial R",    x:250*MM,y:25*MM, w:35*MM, h:22*MM},
        {n:"BitÃ¡cora",        x:20*MM, y:90*MM, w:40*MM, h:56*MM},
        {n:"Manual",          x:70*MM, y:90*MM, w:40*MM, h:56*MM},
        {n:"Sello",           x:120*MM,y:90*MM, w:25*MM, h:25*MM},
        {n:"Firma",           x:155*MM,y:95*MM, w:55*MM, h:18*MM},
        {n:"Propuesta",       x:220*MM,y:90*MM, w:40*MM, h:56*MM}
    ];

    for (var p = 0; p < piezas.length; p++) {
        var pi = piezas[p];
        makeRectAI(L.Elementos_Graficos, pi.x, pi.y, pi.w, pi.h, COLOR.navy);
        // Borde dorado
        var borderRect = L.Elementos_Graficos.pathItems.rectangle(-pi.y, pi.x, pi.w, pi.h);
        borderRect.filled = false;
        borderRect.stroked = true;
        borderRect.strokeColor = makeCMYK(COLOR.gold[0],COLOR.gold[1],COLOR.gold[2],COLOR.gold[3]);
        borderRect.strokeWidth = 0.5;
        makeTextAI(L.Tipografia, pi.n, pi.x + 1*MM, pi.y + pi.h + 4*MM, 5, FONT.light, COLOR.white);
    }

    makeTextAI(L.Tipografia, "13 piezas  |  CMYK 300dpi  |  ISO Coated v2",
               20*MM, H - 10*MM, 7, FONT.light, COLOR.gold);
}

// â”€â”€â”€ BUILDER PRINCIPAL â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

var BUILDERS = [
    buildHojaMembretada,
    buildTarjetaFrente,
    buildTarjetaReverso,
    buildSobre,
    buildCarpeta,
    buildCredencialFrente,
    buildCredencialReverso,
    buildBitacora,
    buildPortadaManual,
    buildSello,
    buildFirma,
    buildPortadaPropuesta,
    buildMockup
];

// â”€â”€â”€ CREACIÃ“N DEL DOCUMENTO â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function createDocument() {
    var preset = new DocumentPreset();
    preset.title       = "MICSA_Safety_Identity_v1.0";
    preset.colorMode   = DocumentColorSpace.CMYK;
    preset.units       = RulerUnits.Millimeters;
    preset.width       = PT.a4w;
    preset.height      = PT.a4h;
    preset.numArtboards = 1;

    var doc = app.documents.addDocument(DocumentColorSpace.CMYK, preset);
    doc.defaultUnit = RulerUnits.Millimeters;

    // Renombrar el primer artboard
    doc.artboards[0].name = ARTBOARDS[0].name;
    var r0 = ARTBOARDS[0];
    doc.artboards[0].artboardRect = [0, 0, r0.w, -r0.h];

    // Agregar los artboards restantes
    for (var i = 1; i < ARTBOARDS.length; i++) {
        var ab = ARTBOARDS[i];
        // Posicionar horizontalmente con separaciÃ³n de 20mm
        var offsetX = 0;
        for (var j = 0; j < i; j++) {
            offsetX += ARTBOARDS[j].w + 20 * MM;
        }
        var newAB = doc.artboards.add([offsetX, 0, offsetX + ab.w, -ab.h]);
        newAB.name = ab.name;
    }

    return doc;
}

// â”€â”€â”€ EXPORTAR PDF â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function exportPDF(doc, abIndex, filename) {
    var opts = new PDFSaveOptions();
    opts.compatibility       = PDFCompatibility.ACROBAT5;
    opts.colorBars           = false;
    opts.cropMarks           = true;
    opts.registrationMarks   = true;
    opts.preserveEditability = false;
    opts.generateThumbnails  = true;
    opts.artboardRange       = String(abIndex + 1);

    var file = new File(BASE_PATH + "PDFs\\" + filename + ".pdf");
    doc.saveAs(file, opts);
}

// â”€â”€â”€ EXPORTAR PNG â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function exportPNG(doc, abIndex, abName) {
    // Seleccionar el artboard activo
    doc.artboards.setActiveArtboardIndex(abIndex);

    var opts = new ExportOptionsPNG24();
    opts.resolution      = 150;  // 150 dpi para preview (300 para producciÃ³n)
    opts.antiAliasing    = true;
    opts.artBoardClipping = true;
    opts.transparency    = false;

    var file = new File(BASE_PATH + "PNGs\\" + abName + ".png");
    doc.exportFile(file, ExportType.PNG24, opts);
}

// â”€â”€â”€ MAIN â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function main() {
    app.coordinateSystem = CoordinateSystem.ARTBOARDCOORDINATESYSTEM;

    $.writeln("=== MICSA Safety Identity Generator ===");
    $.writeln("Iniciando creaciÃ³n del documento...");

    var doc = createDocument();

    // Construir cada artboard
    for (var i = 0; i < ARTBOARDS.length; i++) {
        $.writeln("Construyendo artboard " + (i+1) + ": " + ARTBOARDS[i].name);
        doc.artboards.setActiveArtboardIndex(i);
        try {
            BUILDERS[i](doc, doc.artboards[i]);
        } catch (e) {
            $.writeln("  ERROR en artboard " + (i+1) + ": " + e.message);
        }
    }

    // Guardar como .AI
    $.writeln("Guardando archivo .AI maestro...");
    var aiFile = new File(BASE_PATH + "MICSA_Safety_Identity_v1.0.ai");
    var saveOpts = new IllustratorSaveOptions();
    saveOpts.compatibility  = Compatibility.ILLUSTRATOR19; // AI 2025
    saveOpts.flattenOutput  = OutputFlattening.PRESERVEAPPEARANCE;
    saveOpts.embedICCProfile = true;
    saveOpts.saveMultipleArtboards = true;
    doc.saveAs(aiFile, saveOpts);

    // Exportar PDFs y PNGs por artboard
    $.writeln("Exportando PDFs...");
    for (var p = 0; p < ARTBOARDS.length; p++) {
        try {
            exportPDF(doc, p, ARTBOARDS[p].name);
            $.writeln("  PDF: " + ARTBOARDS[p].name + " âœ“");
        } catch (e) {
            $.writeln("  PDF ERROR (" + ARTBOARDS[p].name + "): " + e.message);
        }
    }

    $.writeln("Exportando PNGs...");
    for (var n = 0; n < ARTBOARDS.length; n++) {
        try {
            exportPNG(doc, n, ARTBOARDS[n].name);
            $.writeln("  PNG: " + ARTBOARDS[n].name + " âœ“");
        } catch (e) {
            $.writeln("  PNG ERROR (" + ARTBOARDS[n].name + "): " + e.message);
        }
    }

    $.writeln("");
    $.writeln("=== COMPLETADO ===");
    $.writeln("Archivo AI: " + BASE_PATH + "MICSA_Safety_Identity_v1.0.ai");
    $.writeln("PDFs: " + BASE_PATH + "PDFs\\");
    $.writeln("PNGs: " + BASE_PATH + "PNGs\\");

    alert(
        "MICSA Safety Identity â€” Generado exitosamente\n\n" +
        "Archivo AI:  MICSA_Safety_Identity_v1.0.ai\n" +
        "PDFs:        " + ARTBOARDS.length + " archivos\n" +
        "PNGs:        " + ARTBOARDS.length + " previews\n\n" +
        "UbicaciÃ³n: " + BASE_PATH
    );
}

main();

