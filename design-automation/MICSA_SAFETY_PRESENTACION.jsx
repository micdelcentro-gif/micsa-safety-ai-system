/**
 * MICSA SAFETY DIVISION â€” PresentaciÃ³n Corporativa Completa
 * ExtendScript para Adobe Illustrator
 * 7 Artboards â€” TamaÃ±o Carta â€” CMYK
 * USO: Archivo > Scripts > Otro script...
 */

// â”€â”€â”€ CONFIG â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
var BASE_PATH = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\MICSA_Safety\\";
var LOGO_FILE = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\MICSA_Safety\\logo.png";
var OUT_PATH  = BASE_PATH + "PDFs\\";

var MM  = 2.834645669;
var LTW = 216 * MM;   // Carta: 612.284 pt
var LTH = 279 * MM;   // Carta: 791.168 pt
var MRG = 20  * MM;   // 56.69 pt
var BLD = 3   * MM;   // 8.50 pt (sangrado)

// Paleta CMYK
function cmyk(c,m,y,k){ var col=new CMYKColor(); col.cyan=c; col.magenta=m; col.yellow=y; col.black=k; return col; }
var NAVY  = cmyk(85,65,0,40);   // Azul marino
var GOLD  = cmyk(0,20,80,15);   // Dorado metÃ¡lico
var WHITE = cmyk(0,0,0,0);      // Blanco
var DGRAY = cmyk(0,0,0,70);     // Gris oscuro
var LGRAY = cmyk(0,0,0,8);      // Gris muy claro (fondos)
var MGRAY = cmyk(0,0,0,25);     // Gris medio (separadores)
var GOLD2 = cmyk(0,10,40,5);    // Dorado claro

// Fuentes
var FB = "Montserrat-Bold";
var FS = "Montserrat-SemiBold";
var FR = "OpenSans-Regular";
var FL = "OpenSans-Light";

// â”€â”€â”€ HELPERS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function rect(doc, x, y, w, h, fill, stroke, sw){
    var r = doc.pathItems.rectangle(-y, x, w, h);
    r.filled = true; r.fillColor = fill;
    if(stroke){ r.stroked=true; r.strokeColor=stroke; r.strokeWidth=sw||0.5; }
    else r.stroked=false;
    return r;
}

function txt(doc, s, x, y, sz, col, font, w, align){
    var t = doc.textFrames.add();
    t.contents = s; t.left = x; t.top = -y;
    if(w) t.width = w;
    t.textRange.characterAttributes.size = sz;
    t.textRange.characterAttributes.fillColor = col;
    try{ t.textRange.characterAttributes.textFont = app.textFonts.getByName(font); }catch(e){}
    if(align === "C") t.textRange.paragraphAttributes.justification = Justification.CENTER;
    else if(align === "R") t.textRange.paragraphAttributes.justification = Justification.RIGHT;
    return t;
}

function line(doc, x1,y1,x2,y2, col, w){
    var l=doc.pathItems.add();
    l.setEntirePath([[x1,-y1],[x2,-y2]]);
    l.stroked=true; l.strokeColor=col; l.strokeWidth=w||0.5; l.filled=false;
    return l;
}

function logoPlace(doc, x, y, w){
    try{
        var f=new File(LOGO_FILE);
        if(!f.exists) return;
        var p=doc.placedItems.add();
        p.file=f;
        var sc = w/p.width;
        p.width=w; p.height=p.height*sc;
        p.left=x; p.top=-y;
    }catch(e){}
}

// Header corporativo reutilizable
function header(doc, O, titulo){
    rect(doc, O*LTW, 0, LTW, 18*MM, NAVY);
    rect(doc, O*LTW, 18*MM, LTW, 0.8*MM, GOLD);
    txt(doc, "MICSA SAFETY DIVISION", O*LTW+MRG, 5*MM, 8, WHITE, FB, LTW*0.5);
    txt(doc, titulo, O*LTW+MRG, 12*MM, 7, GOLD2, FL, LTW*0.6);
    // NÃºmero de artboard
    txt(doc, String(O+1)+" / 7", O*LTW+LTW-MRG-18, 5*MM, 8, GOLD, FS, 20);
}

// Footer
function footer(doc, O){
    rect(doc, O*LTW, LTH-14*MM, LTW, 14*MM, NAVY);
    rect(doc, O*LTW, LTH-14*MM, LTW, 0.5*MM, GOLD);
    txt(doc, "Confidencial â€” Documento Propietario MICSA Safety Division  |  contacto@micsa.mx  |  800-MICSA-01  |  www.micsa.mx",
        O*LTW+MRG, LTH-9*MM, 6, MGRAY, FL, LTW-MRG*2);
}

// Tabla genÃ©rica
function tabla(doc, O, startY, hdrs, rows, cols){
    var x=O*LTW+MRG, y=startY, rh=10*MM, tw=LTW-MRG*2;
    // header
    var cx=x;
    for(var h=0;h<hdrs.length;h++){
        var cw=cols[h]*tw;
        rect(doc,cx,y,cw,rh,NAVY);
        line(doc,cx,y,cx,y+rh,GOLD,0.3);
        txt(doc,hdrs[h],cx+2*MM,y+3*MM,6.5,WHITE,FB,cw-4*MM);
        cx+=cw;
    }
    y+=rh;
    for(var r=0;r<rows.length;r++){
        cx=x;
        var bg=(r%2===0)?LGRAY:WHITE;
        for(var c=0;c<rows[r].length;c++){
            var cw=cols[c]*tw;
            rect(doc,cx,y,cw,rh,bg,MGRAY,0.3);
            txt(doc,rows[r][c],cx+2*MM,y+3*MM,6.5,DGRAY,FR,cw-4*MM);
            cx+=cw;
        }
        y+=rh;
    }
    return y;
}

// Bloque pilar (para cultura y valor agregado)
function pilar(doc, O, x, y, w, h, titulo, desc, num){
    rect(doc, x, y, w, h, WHITE, MGRAY, 0.5);
    // barra superior dorada
    rect(doc, x, y, w, 2*MM, GOLD);
    // nÃºmero
    rect(doc, x+3*MM, y+4*MM, 12*MM, 12*MM, NAVY);
    txt(doc, num, x+4*MM, y+5.5*MM, 10, WHITE, FB, 12*MM, "C");
    // tÃ­tulo
    txt(doc, titulo, x+18*MM, y+5*MM, 8, NAVY, FB, w-22*MM);
    // desc
    txt(doc, desc, x+3*MM, y+19*MM, 7, DGRAY, FR, w-6*MM);
}

// Timeline step
function timeStep(doc, O, x, y, fase, titulo, desc, isLast){
    // CÃ­rculo
    var cr=6*MM;
    var ell=doc.pathItems.ellipse(-(y), x, cr*2, cr*2);
    ell.filled=true; ell.fillColor=NAVY; ell.stroked=true; ell.strokeColor=GOLD; ell.strokeWidth=1.5;
    txt(doc, fase, x+1*MM, y+cr*0.6, 7, WHITE, FB, cr*2, "C");
    // LÃ­nea conectora
    if(!isLast) line(doc, x+cr*2, y+cr, x+cr*2+24*MM, y+cr, MGRAY, 1);
    // Texto
    txt(doc, titulo, x-2*MM, y+cr*2+3*MM, 7.5, NAVY, FB, cr*4+8*MM);
    txt(doc, desc,   x-2*MM, y+cr*2+12*MM, 6, DGRAY, FL, cr*4+8*MM);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// ARTBOARD 1 â€” PORTADA
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function portada(doc){
    var O=0;
    // Fondo blanco
    rect(doc, 0, 0, LTW, LTH, WHITE);
    // Bloque navy superior 55%
    rect(doc, 0, 0, LTW, LTH*0.55, NAVY);
    // Barra gold
    rect(doc, 0, LTH*0.55, LTW, 2*MM, GOLD);
    // Franja lateral izq decorativa
    rect(doc, 0, 0, 8*MM, LTH, GOLD);

    // Logo
    logoPlace(doc, LTW*0.5-35*MM, 15*MM, 70*MM);

    // Textos portada
    txt(doc, "MICSA SAFETY DIVISION", 12*MM, LTH*0.34, 22, WHITE, FB, LTW-14*MM);
    txt(doc, "Seguridad Patrimonial Industrial", 12*MM, LTH*0.34+18, 12, GOLD2, FS, LTW-14*MM);

    // LÃ­nea dorada corta
    rect(doc, 12*MM, LTH*0.34+34, 60*MM, 1.5*MM, GOLD);

    txt(doc, '"Disciplina, criterio y control operativo"', 12*MM, LTH*0.34+42, 10, LGRAY, FL, LTW-14*MM);

    // Bloque info inferior
    rect(doc, 12*MM, LTH*0.60, LTW-24*MM, LTH*0.28, LGRAY);
    rect(doc, 12*MM, LTH*0.60, 2*MM, LTH*0.28, NAVY);

    txt(doc, "PROPUESTA COMERCIAL 2026", 18*MM, LTH*0.62, 9, NAVY, FB, LTW*0.5);
    txt(doc, "STELLANTIS â€” FCA PLANTA SALTILLO", 18*MM, LTH*0.62+12, 8, DGRAY, FR, LTW*0.5);
    txt(doc, new Date().toLocaleDateString('es-MX',{}), 18*MM, LTH*0.62+24, 8, DGRAY, FL, LTW*0.5);

    // Footer portada
    rect(doc, 0, LTH-14*MM, LTW, 14*MM, NAVY);
    rect(doc, 0, LTH-14*MM, LTW, 0.5*MM, GOLD);
    txt(doc, "MICSA Safety Division  Â·  Seguridad Industrial  Â·  contacto@micsa.mx", 12*MM, LTH-9*MM, 6.5, MGRAY, FL, LTW-24*MM);
    txt(doc, "1 / 7", LTW-MRG-20, LTH-9*MM, 7, GOLD, FS, 20);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// ARTBOARD 2 â€” CULTURA OPERATIVA
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function cultura(doc){
    var O=1;
    rect(doc, O*LTW, 0, LTW, LTH, WHITE);
    header(doc, O, "CULTURA OPERATIVA");
    footer(doc, O);

    var y=22*MM;
    txt(doc, "LOS 4 PILARES DE MICSA SAFETY", O*LTW+MRG, y, 13, NAVY, FB, LTW-MRG*2); y+=8*MM;
    line(doc, O*LTW+MRG, y, O*LTW+LTW-MRG, y, GOLD, 1); y+=6*MM;
    txt(doc, "Nuestra cultura no es un documento. Es la forma en que cada guardia piensa, actÃºa y decide en campo.", O*LTW+MRG, y, 8, DGRAY, FL, LTW-MRG*2); y+=12*MM;

    var pw=(LTW-MRG*2)*0.48;
    var ph=52*MM;
    var pilares=[
        {n:"01", t:"RESPETO OPERATIVO", d:"Toda interacciÃ³n con personal del cliente, visitantes y compaÃ±eros refleja la imagen de MICSA Safety. El respeto es la base de la confianza operativa y el requisito mÃ­nimo para permanecer en el puesto."},
        {n:"02", t:"DISCIPLINA REAL",   d:"Los protocolos existen para cumplirse, no para interpretarse. Sin disciplina no hay control. La disciplina no es negociable: aplica en turno, uniforme, bitÃ¡cora y reporte sin excepciÃ³n."},
        {n:"03", t:"CRITERIO EN CAMPO", d:"Un guardia entrenado toma decisiones correctas bajo presiÃ³n. El criterio se desarrolla con capacitaciÃ³n continua, evaluaciÃ³n mensual y retroalimentaciÃ³n estructurada de supervisores."},
        {n:"04", t:"CONFIABILIDAD DESDE EL ORIGEN", d:"El cliente confÃ­a en MICSA Safety porque seleccionamos, evaluamos y supervisamos a cada elemento. La confiabilidad empieza en el reclutamiento y se sostiene con KPIs medibles."}
    ];
    for(var i=0;i<pilares.length;i++){
        var px=O*LTW+MRG + (i%2)*(pw+MRG*0.5);
        var py=y + Math.floor(i/2)*(ph+5*MM);
        pilar(doc, O, px, py, pw, ph, pilares[i].t, pilares[i].d, pilares[i].n);
    }
    y += 2*ph + 5*MM + 10*MM;

    // Frase institucional
    rect(doc, O*LTW+MRG, y, LTW-MRG*2, 18*MM, NAVY);
    rect(doc, O*LTW+MRG, y, 3*MM, 18*MM, GOLD);
    txt(doc, '"Puede faltar flujo, pero nunca posiciÃ³n."', O*LTW+MRG+8*MM, y+5*MM, 10, WHITE, FS, LTW-MRG*2-12*MM);
    txt(doc, "Principio operativo MICSA Safety Division", O*LTW+MRG+8*MM, y+12*MM, 7, GOLD2, FL, LTW-MRG*2);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// ARTBOARD 3 â€” PLAN DE CAPACITACIÃ“N
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function capacitacion(doc){
    var O=2;
    rect(doc, O*LTW, 0, LTW, LTH, WHITE);
    header(doc, O, "PLAN DE CAPACITACIÃ“N Y DESARROLLO");
    footer(doc, O);

    var y=22*MM;
    txt(doc, "PROGRAMA DE FORMACIÃ“N OPERATIVA", O*LTW+MRG, y, 13, NAVY, FB, LTW-MRG*2); y+=8*MM;
    line(doc, O*LTW+MRG, y, O*LTW+LTW-MRG, y, GOLD, 1); y+=6*MM;
    txt(doc, "5 fases de formaciÃ³n continua para garantizar que cada elemento opere con criterio, protocolo y estÃ¡ndar MICSA.", O*LTW+MRG, y, 8, DGRAY, FL, LTW-MRG*2); y+=16*MM;

    // Timeline horizontal
    var fases=[
        {f:"01", t:"INDUCCIÃ“N",      d:"Historia MICSA\nValores y cultura\nReglamento interno\nPresentaciÃ³n cliente\n(DÃ­as 1-3)"},
        {f:"02", t:"FORMACIÃ“N\nTÃ‰CNICA",   d:"Control de accesos\nRondines QR\nUso EPP correcto\nBitÃ¡cora digital\n(DÃ­as 4-10)"},
        {f:"03", t:"ESPECIALI-\nZACIÃ“N",   d:"Protocolos NOM-030\nManejo cuatrimoto\nMonitoreo CCTV\nReporte incidentes\n(DÃ­as 11-20)"},
        {f:"04", t:"EVALUACIÃ“N", d:"Examen tÃ©cnico\nSimulaciÃ³n campo\nPsicomÃ©trica\nCriterio â‰¥80%\n(DÃ­a 21-25)"},
        {f:"05", t:"CAPACIT.\nCONTINUA",  d:"SesiÃ³n mensual\nKPI revisiÃ³n\nActualizaciÃ³n NOM\nRetroalimentaciÃ³n\n(Mensual)"}
    ];
    var stepW=(LTW-MRG*2)/fases.length;
    for(var i=0;i<fases.length;i++){
        var sx=O*LTW+MRG+i*stepW+stepW/2-6*MM;
        timeStep(doc, O, sx, y, fases[i].f, fases[i].t, fases[i].d, i===fases.length-1);
    }
    y += 55*MM;

    // Matriz de evaluaciÃ³n
    txt(doc, "MATRIZ DE EVALUACIÃ“N DE INGRESO", O*LTW+MRG, y, 11, NAVY, FB, LTW-MRG*2); y+=8*MM;
    line(doc, O*LTW+MRG, y, O*LTW+LTW-MRG, y, GOLD, 0.8); y+=5*MM;
    var mRows=[
        ["Actitud y disciplina",    "25%","Entrevista estructurada + observaciÃ³n de campo"],
        ["EvaluaciÃ³n tÃ©cnica",      "25%","SimulaciÃ³n en campo: control acceso, reporte, EPP"],
        ["Experiencia comprobada",  "20%","VerificaciÃ³n directa con empleadores anteriores"],
        ["Prueba psicomÃ©trica",     "15%","Test honestidad, confiabilidad y reacciÃ³n a presiÃ³n"],
        ["Referencias laborales",   "15%","Llamada directa a empresa anterior â€” cargo y motivo baja"],
        ["MÃNIMO DE APROBACIÃ“N",    "80%","Elemento no aprobado = no ingresa. Sin excepciones."]
    ];
    tabla(doc, O, y, ["FACTOR","PESO","MÃ‰TODO DE MEDICIÃ“N"], mRows, [0.30,0.10,0.60]);
    y += (mRows.length+1)*10*MM + 10*MM;

    // DuraciÃ³n y compromisos
    txt(doc, "COMPROMISOS DE CAPACITACIÃ“N CONTINUA", O*LTW+MRG, y, 11, NAVY, FB, LTW-MRG*2); y+=8*MM;
    line(doc, O*LTW+MRG, y, O*LTW+LTW-MRG, y, GOLD, 0.8); y+=5*MM;
    var comprRows=[
        ["SesiÃ³n mensual obligatoria","60 min","Todos los elementos asignados","Supervisor + Jefe Seguridad"],
        ["EvaluaciÃ³n de desempeÃ±o","Mensual","Individual por guardia","Supervisor FCA"],
        ["ActualizaciÃ³n normativa","Semestral","NOM-030, NOM-017 y reglamento cliente","DirecciÃ³n MICSA"],
        ["CertificaciÃ³n EPP","Anual","Examen prÃ¡ctico uso correcto","MICSA + cliente"]
    ];
    tabla(doc, O, y, ["ACTIVIDAD","FRECUENCIA","ALCANCE","RESPONSABLE"], comprRows, [0.30,0.12,0.33,0.25]);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// ARTBOARD 4 â€” VALOR AGREGADO
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function valorAgregado(doc){
    var O=3;
    rect(doc, O*LTW, 0, LTW, LTH, WHITE);
    header(doc, O, "VALOR AGREGADO â€” POR QUÃ‰ MICSA SAFETY");
    footer(doc, O);

    var y=22*MM;
    txt(doc, "VENTAJAS COMPETITIVAS REALES", O*LTW+MRG, y, 13, NAVY, FB, LTW-MRG*2); y+=8*MM;
    line(doc, O*LTW+MRG, y, O*LTW+LTW-MRG, y, GOLD, 1); y+=6*MM;
    txt(doc, "No competimos por precio. Competimos por autoridad, estructura y resultados medibles.", O*LTW+MRG, y, 8, DGRAY, FL, LTW-MRG*2); y+=14*MM;

    var valores=[
        {n:"01",t:"PERSONAL CON CRITERIO",       d:"Cada elemento pasa por filtro psicomÃ©trico, evaluaciÃ³n tÃ©cnica y verificaciÃ³n de antecedentes. No contratamos cuerpos â€” contratamos criterio. El guardia que no aprueba no ingresa."},
        {n:"02",t:"CONTROL OPERATIVO REAL",       d:"BitÃ¡cora digital, rondines con QR, reportes de incidente en tiempo real. El cliente tiene visibilidad del 100% de la operaciÃ³n. Sin zonas grises, sin turnos sin evidencia."},
        {n:"03",t:"INTEGRACIÃ“N INDUSTRIAL",       d:"Conocemos los ritmos de producciÃ³n automotriz. Nuestros supervisores hablan el mismo lenguaje que los superintendentes de planta. Nos integramos al sistema del cliente, no al revÃ©s."},
        {n:"04",t:"SUPERVISIÃ“N ACTIVA",           d:"Jefe de Seguridad + Supervisor en sitio. No supervision remota, no modelos de \"autogestiÃ³n\". Alguien responsable presente en cada turno. Respuesta a incidente garantizada en minutos."},
        {n:"05",t:"TRAZABILIDAD DOCUMENTAL",      d:"Cada guardia tiene expediente completo: contrato, evaluaciones, capacitaciones, sanciones e historial. Disponible para auditorÃ­a del cliente en cualquier momento. Sin sorpresas."}
    ];
    var vw=LTW-MRG*2;
    var vh=33*MM;
    for(var i=0;i<valores.length;i++){
        var vx=O*LTW+MRG;
        var vy=y+i*(vh+3*MM);
        rect(doc, vx, vy, vw, vh, WHITE, MGRAY, 0.4);
        rect(doc, vx, vy, 2.5*MM, vh, GOLD);
        rect(doc, vx+2.5*MM, vy, 20*MM, vh, LGRAY);
        txt(doc, valores[i].n, vx+3*MM, vy+8*MM, 12, NAVY, FB, 18*MM, "C");
        txt(doc, valores[i].t, vx+26*MM, vy+6*MM, 8.5, NAVY, FB, vw-30*MM);
        txt(doc, valores[i].d, vx+26*MM, vy+16*MM, 6.8, DGRAY, FR, vw-30*MM);
    }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// ARTBOARD 5 â€” SISTEMA OPERATIVO / MANUALES
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function manuales(doc){
    var O=4;
    rect(doc, O*LTW, 0, LTW, LTH, WHITE);
    header(doc, O, "SISTEMA OPERATIVO â€” MANUALES Y PROTOCOLOS");
    footer(doc, O);

    var y=22*MM;
    txt(doc, "ARQUITECTURA DE DOCUMENTACIÃ“N OPERATIVA", O*LTW+MRG, y, 13, NAVY, FB, LTW-MRG*2); y+=8*MM;
    line(doc, O*LTW+MRG, y, O*LTW+LTW-MRG, y, GOLD, 1); y+=6*MM;
    txt(doc, "Cada documento estÃ¡ alineado a normas NOM-STPS y al reglamento interno del cliente. Disponible en formato digital e impreso.", O*LTW+MRG, y, 8, DGRAY, FL, LTW-MRG*2); y+=12*MM;

    var mans=[
        {cod:"M-01",t:"MANUAL GENERAL DE OPERACIONES",
         items:["Estructura organizacional y cadena de mando","DescripciÃ³n de puestos y funciones por nivel","PolÃ­ticas de contrataciÃ³n, evaluaciÃ³n y baja","ComunicaciÃ³n interna y reportes oficiales","Procedimiento de sustituciÃ³n de personal"]},
        {cod:"M-02",t:"MANUAL INTRAMUROS â€” CONTROL DE PLANTA",
         items:["Control de acceso vehicular y peatonal","Registro de visitas y contratistas","Protocolos de cierre y apertura de planta","Manejo de materiales y activos crÃ­ticos","CoordinaciÃ³n con seguridad del cliente"]},
        {cod:"M-03",t:"MANUAL DE MONITOREO CCTV",
         items:["OperaciÃ³n de sistema de videovigilancia","Protocolo de detecciÃ³n de anomalÃ­as","Reporte y evidencia de incidentes en cÃ¡mara","Mantenimiento bÃ¡sico del equipo","BitÃ¡cora de monitoreo turno a turno"]},
        {cod:"M-04",t:"PROTOCOLO DE USO DE FUERZA",
         items:["Principios de proporcionalidad y necesidad","Niveles de respuesta a amenaza fÃ­sica","CoordinaciÃ³n con autoridades (policÃ­a, FGR)","DocumentaciÃ³n post-incidente obligatoria","Marco legal aplicable (LGEEPA, LFT)"]},
        {cod:"M-05",t:"MANUAL DE RECLUTAMIENTO",
         items:["Perfil mÃ­nimo por puesto (escolaridad, edad, experiencia)","Proceso de 5 fases con criterio â‰¥80%","DocumentaciÃ³n obligatoria de ingreso","PerÃ­odo de prueba 30 dÃ­as con evaluaciÃ³n semanal","Causales de baja y procedimiento formal"]},
        {cod:"M-06",t:"REGLAMENTO INTERNO",
         items:["CÃ³digo de vestimenta y uniforme completo","Uso de dispositivos mÃ³viles (prohibido en turno)","BitÃ¡cora obligatoria â€” sin firma = turno incompleto","Cero tolerancia: alcohol, sustancias, propinas","Sanciones progresivas y causal de baja inmediata"]}
    ];
    var bw=(LTW-MRG*2)*0.48;
    var bh=54*MM;
    for(var i=0;i<mans.length;i++){
        var bx=O*LTW+MRG+(i%2)*(bw+MRG*0.5);
        var by=y+Math.floor(i/2)*(bh+4*MM);
        rect(doc, bx, by, bw, bh, WHITE, MGRAY, 0.4);
        rect(doc, bx, by, bw, 10*MM, NAVY);
        rect(doc, bx, by, 2*MM, bh, GOLD);
        txt(doc, mans[i].cod, bx+4*MM, by+2.5*MM, 7, GOLD2, FS, 18*MM);
        txt(doc, mans[i].t, bx+23*MM, by+2.5*MM, 6.5, WHITE, FB, bw-26*MM);
        for(var j=0;j<mans[i].items.length;j++){
            txt(doc, "â€¢ "+mans[i].items[j], bx+4*MM, by+12*MM+j*7*MM, 6, DGRAY, FR, bw-8*MM);
        }
    }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// ARTBOARD 6 â€” PLAN DE IMPLEMENTACIÃ“N
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function implementacion(doc){
    var O=5;
    rect(doc, O*LTW, 0, LTW, LTH, WHITE);
    header(doc, O, "PLAN DE IMPLEMENTACIÃ“N â€” 5 FASES");
    footer(doc, O);

    var y=22*MM;
    txt(doc, "RUTA CRÃTICA DE ARRANQUE OPERATIVO", O*LTW+MRG, y, 13, NAVY, FB, LTW-MRG*2); y+=8*MM;
    line(doc, O*LTW+MRG, y, O*LTW+LTW-MRG, y, GOLD, 1); y+=6*MM;
    txt(doc, "Desde la firma del contrato hasta operaciÃ³n plena: 15 dÃ­as hÃ¡biles con hitos medibles.", O*LTW+MRG, y, 8, DGRAY, FL, LTW-MRG*2); y+=14*MM;

    var fases=[
        {f:"F1",t:"ESTRUCTURA",    dias:"DÃ­as 1-3",   col:NAVY,
         items:["Alta en sistema de proveedores cliente","Firma contrato + PO","IntegraciÃ³n REPSE y documentaciÃ³n fiscal","DesignaciÃ³n Jefe de Seguridad asignado","Entrega de carta de inicio formal"]},
        {f:"F2",t:"LEGAL",         dias:"DÃ­as 4-6",   col:NAVY,
         items:["RevisiÃ³n reglamento interno cliente","Acuerdo de NDA y confidencialidad","Alta IMSS del personal asignado","Registro en nÃ³mina con CURP y datos","Alta en STPS y obligaciones patronales"]},
        {f:"F3",t:"OPERACIÃ“N",     dias:"DÃ­as 7-10",  col:NAVY,
         items:["InducciÃ³n 3 dÃ­as del personal seleccionado","Entrega de uniformes y EPP completo","Reconocimiento fÃ­sico de instalaciones","Acuerdo de puntos de control y rondines","ConfiguraciÃ³n de bitÃ¡cora y reportes"]},
        {f:"F4",t:"IMPLEMENTACIÃ“N",dias:"DÃ­as 11-13", col:NAVY,
         items:["OperaciÃ³n supervisada (Jefe en campo)","ValidaciÃ³n de protocolos con cliente","Ajuste de turnos y puntos de control","ActivaciÃ³n de monitoreo CCTV","Primera bitÃ¡cora semanal entregada"]},
        {f:"F5",t:"CONTROL",       dias:"DÃ­as 14-15", col:NAVY,
         items:["Dashboard KPI del primer perÃ­odo","ReuniÃ³n de alineaciÃ³n cliente-supervisor","Plan de mejora continua acordado","Fecha de revisiÃ³n mensual agendada","OperaciÃ³n plena transferida"]}
    ];
    var fw=LTW-MRG*2;
    var fh=40*MM;
    for(var i=0;i<fases.length;i++){
        var fx=O*LTW+MRG;
        var fy=y+i*(fh+3*MM);
        // Contenedor
        rect(doc, fx, fy, fw, fh, WHITE, MGRAY, 0.4);
        // Columna izq (fase)
        rect(doc, fx, fy, 32*MM, fh, NAVY);
        txt(doc, fases[i].f, fx+2*MM, fy+5*MM, 18, GOLD, FB, 28*MM, "C");
        txt(doc, fases[i].t, fx+2*MM, fy+20*MM, 7, WHITE, FS, 28*MM, "C");
        txt(doc, fases[i].dias, fx+2*MM, fy+29*MM, 6.5, GOLD2, FL, 28*MM, "C");
        // LÃ­nea separadora
        line(doc, fx+32*MM, fy+5*MM, fx+32*MM, fy+fh-5*MM, GOLD, 0.8);
        // Items
        var iw=(fw-36*MM)/2;
        for(var j=0;j<fases[i].items.length;j++){
            var ix=fx+35*MM + Math.floor(j/3)*iw;
            var iy=fy+5*MM + (j%3)*11*MM;
            txt(doc, "â–¸ "+fases[i].items[j], ix, iy, 6.5, DGRAY, FR, iw-5*MM);
        }
    }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// ARTBOARD 7 â€” MOCKUP PAPELERÃA
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function mockupPapeleria(doc){
    var O=6;
    rect(doc, O*LTW, 0, LTW, LTH, LGRAY);
    header(doc, O, "PAPELERÃA CORPORATIVA â€” SISTEMA DE IDENTIDAD");
    footer(doc, O);

    var y=22*MM;
    txt(doc, "SISTEMA DE PAPELERÃA MICSA SAFETY DIVISION", O*LTW+MRG, y, 13, NAVY, FB, LTW-MRG*2); y+=8*MM;
    line(doc, O*LTW+MRG, y, O*LTW+LTW-MRG, y, GOLD, 1); y+=8*MM;

    // â”€â”€ Hoja membretada (mini) â”€â”€
    var hx=O*LTW+MRG, hy=y, hw=LTW*0.44, hh=60*MM;
    rect(doc, hx, hy, hw, hh, WHITE, MGRAY, 0.5);
    rect(doc, hx, hy, hw, 9*MM, NAVY);
    rect(doc, hx, hy+9*MM, hw, 1.2*MM, GOLD);
    logoPlace(doc, hx+2*MM, hy+1*MM, 20*MM);
    txt(doc, "MICSA SAFETY DIVISION", hx+24*MM, hy+2.5*MM, 5.5, WHITE, FB, hw-26*MM);
    txt(doc, "Seguridad Patrimonial Industrial", hx+24*MM, hy+7*MM, 4.5, GOLD2, FL, hw-26*MM);
    txt(doc, "Hoja Membretada", hx+3*MM, hy+12*MM, 5, DGRAY, FL, hw-6*MM);
    txt(doc, "â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”", hx+3*MM, hy+17*MM, 4.5, MGRAY, FL, hw-6*MM);
    txt(doc, "â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”", hx+3*MM, hy+22*MM, 4.5, MGRAY, FL, hw-6*MM);
    txt(doc, "â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”  â€”", hx+3*MM, hy+27*MM, 4.5, MGRAY, FL, hw-6*MM);
    txt(doc, "â€”  â€”  â€”  â€”  â€”  â€”  â€”", hx+3*MM, hy+32*MM, 4.5, MGRAY, FL, hw-6*MM);
    txt(doc, "Hoja A4 Â· CMYK Â· Montserrat + Open Sans", hx+3*MM, hy+hh-5*MM, 4, DGRAY, FL, hw-6*MM);

    // â”€â”€ Tarjeta de presentaciÃ³n â”€â”€
    var tx=O*LTW+LTW-MRG-LTW*0.44, ty=y, tw=LTW*0.44, th=25*MM;
    rect(doc, tx, ty, tw, th, NAVY, MGRAY, 0.5);
    rect(doc, tx, ty+th-2.5*MM, tw, 2.5*MM, GOLD);
    logoPlace(doc, tx+2*MM, ty+1.5*MM, 12*MM);
    txt(doc, "NOMBRE COMPLETO", tx+16*MM, ty+4*MM, 5.5, WHITE, FB, tw-18*MM);
    txt(doc, "Jefe de Seguridad Industrial", tx+16*MM, ty+9*MM, 4.5, GOLD2, FL, tw-18*MM);
    line(doc, tx+16*MM, ty+13*MM, tx+tw-3*MM, ty+13*MM, MGRAY, 0.4);
    txt(doc, "contacto@micsa.mx  |  800-MICSA-01", tx+2*MM, ty+15*MM, 4, LGRAY, FL, tw-4*MM);
    txt(doc, "Tarjeta CR85 Â· Frente", tx+3*MM, ty+th-1*MM, 3.5, GOLD2, FL, tw-6*MM);

    // â”€â”€ Tarjeta reverso â”€â”€
    var rx=O*LTW+LTW-MRG-LTW*0.44, ry=ty+th+3*MM, rw=LTW*0.44, rh=25*MM;
    rect(doc, rx, ry, rw, rh, WHITE, MGRAY, 0.5);
    rect(doc, rx, ry, rw, 2.5*MM, GOLD);
    rect(doc, rx, ry+rh-2.5*MM, rw, 2.5*MM, NAVY);
    txt(doc, "MICSA SAFETY DIVISION", rx+3*MM, ry+5*MM, 5.5, NAVY, FB, rw-6*MM, "C");
    txt(doc, "Seguridad Patrimonial Industrial", rx+3*MM, ry+10*MM, 4.5, DGRAY, FL, rw-6*MM, "C");
    txt(doc, '"Disciplina, criterio y control operativo"', rx+3*MM, ry+15.5*MM, 4, DGRAY, FL, rw-6*MM, "C");
    txt(doc, "Tarjeta CR85 Â· Reverso", rx+3*MM, ry+rh-1*MM, 3.5, MGRAY, FL, rw-6*MM);

    y += 65*MM;

    // â”€â”€ Credencial de empleado â”€â”€
    var cx2=O*LTW+MRG, cy2=y, cw2=LTW*0.44, ch2=32*MM;
    rect(doc, cx2, cy2, cw2, ch2, WHITE, MGRAY, 0.5);
    rect(doc, cx2, cy2, cw2, 7*MM, NAVY);
    rect(doc, cx2, cy2+7*MM, 1.5*MM, ch2-7*MM, GOLD);
    txt(doc, "CREDENCIAL DE EMPLEADO", cx2+3*MM, cy2+2*MM, 5, WHITE, FB, cw2-6*MM);
    txt(doc, "MICSA Safety Division", cx2+3*MM, cy2+5.5*MM, 4, GOLD2, FL, cw2-6*MM);
    // Foto placeholder
    rect(doc, cx2+3*MM, cy2+9*MM, 16*MM, 20*MM, LGRAY, MGRAY, 0.4);
    txt(doc, "FOTO", cx2+3*MM, cy2+16*MM, 5, MGRAY, FL, 16*MM, "C");
    txt(doc, "NOMBRE COMPLETO", cx2+21*MM, cy2+10*MM, 5.5, NAVY, FB, cw2-24*MM);
    txt(doc, "No. Empleado: MSF-0001", cx2+21*MM, cy2+16*MM, 4.5, DGRAY, FR, cw2-24*MM);
    txt(doc, "Ãrea: Seguridad Industrial", cx2+21*MM, cy2+21*MM, 4.5, DGRAY, FR, cw2-24*MM);
    txt(doc, "Vigencia: 31/12/2026", cx2+21*MM, cy2+26*MM, 4.5, DGRAY, FR, cw2-24*MM);
    txt(doc, "Credencial CR80 Â· 3 tinta", cx2+3*MM, cy2+ch2-1*MM, 3.5, MGRAY, FL, cw2-6*MM);

    // â”€â”€ Sobre DL â”€â”€
    var sx=O*LTW+LTW-MRG-LTW*0.44, sy=y, sw=LTW*0.44, sh=32*MM;
    rect(doc, sx, sy, sw, sh, WHITE, MGRAY, 0.5);
    rect(doc, sx, sy, sw, 1.5*MM, NAVY);
    rect(doc, sx, sy+sh-1.5*MM, sw, 1.5*MM, GOLD);
    logoPlace(doc, sx+3*MM, sy+4*MM, 14*MM);
    txt(doc, "MICSA SAFETY DIVISION", sx+19*MM, sy+5*MM, 5, NAVY, FB, sw-22*MM);
    txt(doc, "Seguridad Patrimonial Industrial", sx+19*MM, sy+10*MM, 4, DGRAY, FL, sw-22*MM);
    txt(doc, "contacto@micsa.mx  Â·  800-MICSA-01", sx+19*MM, sy+15*MM, 4, DGRAY, FL, sw-22*MM);
    // Ventana sobre
    rect(doc, sx+sw-25*MM, sy+5*MM, 22*MM, 14*MM, LGRAY, MGRAY, 0.4);
    txt(doc, "DESTINATARIO", sx+sw-24*MM, sy+9*MM, 4, MGRAY, FL, 20*MM);
    txt(doc, "Sobre DL Â· 110Ã—220mm", sx+3*MM, sy+sh-1*MM, 3.5, MGRAY, FL, sw-6*MM);

    y += 38*MM;

    // Nota de sistema
    rect(doc, O*LTW+MRG, y, LTW-MRG*2, 16*MM, NAVY);
    rect(doc, O*LTW+MRG, y, LTW-MRG*2, 1*MM, GOLD);
    txt(doc, "SISTEMA DE IDENTIDAD VISUAL â€” Para archivos editables completos (.AI) y exportaciones PDF/PNG consultar:", O*LTW+MRG+4*MM, y+4*MM, 6.5, WHITE, FL, LTW-MRG*2-8*MM);
    txt(doc, "MICSA_Safety/MICSA_Safety_Identity.jsx", O*LTW+MRG+4*MM, y+10*MM, 7, GOLD2, FS, LTW-MRG*2-8*MM);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// MAIN
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function main(){
    // Crear documento
    var doc=app.documents.add(DocumentColorSpace.CMYK, LTW*7, LTH);

    // Artboards (ya existe 1, agregar 6 mÃ¡s)
    for(var i=1;i<7;i++) doc.artboards.add([i*LTW, 0, (i+1)*LTW, -LTH]);
    var names=["01_PORTADA","02_CULTURA","03_CAPACITACION","04_VALOR_AGREGADO","05_MANUALES","06_IMPLEMENTACION","07_MOCKUP_PAPELERIA"];
    for(var i=0;i<7;i++) doc.artboards[i].name=names[i];

    // Dibujar
    portada(doc);
    cultura(doc);
    capacitacion(doc);
    valorAgregado(doc);
    manuales(doc);
    implementacion(doc);
    mockupPapeleria(doc);

    // Guardar .ai
    var saveFile = new File(BASE_PATH + "MICSA_SAFETY_PRESENTACION.ai");
    var opts = new IllustratorSaveOptions();
    opts.compatibility=Compatibility.ILLUSTRATOR24;
    opts.saveMultipleArtboards = true;
    doc.saveAs(saveFile, opts);

    // PDF
    try{
        var pdfFile = new File(OUT_PATH + "MICSA_SAFETY_PRESENTACION.pdf");
        var pdfOpts = new PDFSaveOptions();
        pdfOpts.compatibility=PDFCompatibility.ACROBAT8;
        pdfOpts.generateThumbnails = true;
        pdfOpts.preserveEditability = false;
        pdfOpts.saveMultipleArtboards = true;
        pdfOpts.artboardRange = "1-7";
        doc.saveAs(pdfFile, pdfOpts);
    }catch(e){}

    alert("MICSA_SAFETY_PRESENTACION.ai generada â€” 7 artboards.\n\nPDF exportado en PDFs/\n\n7 pÃ¡ginas: Portada Â· Cultura Â· CapacitaciÃ³n Â· Valor Agregado Â· Manuales Â· ImplementaciÃ³n Â· Mockup");
}

main();



