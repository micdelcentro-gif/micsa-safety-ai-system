// Polyfill Array.map para ExtendScript ES3
if(!Array.prototype.map){Array.prototype.map=function(fn){var r=[];for(var i=0;i<this.length;i++)r.push(fn(this[i],i,this));return r;};}

/**
 * MICSA SAFETY DIVISION — Cotización FCA FreightCar v2
 * Folio: MSD-COT-033 — Datos financieros reales validados
 * ExtendScript Adobe Illustrator — 6 artboards A4 CMYK
 *
 * DATOS FUENTE: Cotizacion Guardias Seguridad Micsa FCA.xlsx
 * RFC: MIC2301268S5  |  REPSE: 282364
 * Personal total: 40 (Planta 27 + Estancias 7 + Norte 6)
 */

var BASE_PATH = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\MICSA_Safety\\";
var LOGO_FILE = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\adobe\\ID-5608-20260329T034218Z-1-001\\ID-5608\\LOGO 7\\editable\\editable.ai";
var OUT_PATH  = BASE_PATH + "PDFs\\";

var MM=2.834645669, A4W=210*MM, A4H=297*MM, MRG=20*MM;

// ─── COLORES CMYK ───────────────────────────────────────────────────────────
function cmyk(c,m,y,k){var col=new CMYKColor();col.cyan=c;col.magenta=m;col.yellow=y;col.black=k;return col;}
var C_NEGRO  = cmyk(60,40,40,100);
var C_PLATA  = cmyk(0,0,0,30);
var C_GRIS   = cmyk(0,0,0,55);
var C_GRISC  = cmyk(0,0,0,8);
var C_BLANCO = cmyk(0,0,0,0);
var C_ROJO   = cmyk(0,100,85,5);    // rojo corporativo MICSA
var C_ROJOSC = cmyk(0,100,90,20);   // rojo oscuro

// ─── TIPOGRAFÍAS ─────────────────────────────────────────────────────────────
var FB="Montserrat-Bold", FS="Montserrat-SemiBold", FR="OpenSans-Regular", FL="OpenSans-Light";

// ─── FECHA Y FOLIO ────────────────────────────────────────────────────────────
var HOY   = new Date();
var FECHA = HOY.getDate()+"/"+(HOY.getMonth()+1)+"/"+HOY.getFullYear();
var FOLIO = "MSD-COT-033";

// ═══════════════════════════════════════════════════════════════════════════════
// DATOS FINANCIEROS REALES (fuente: Excel validado 2026)
// ═══════════════════════════════════════════════════════════════════════════════

// ── PERSONAL POR UBICACIÓN ──────────────────────────────────────────────────
var PERSONAL_PLANTA = [
    {puesto:"Jefe de Seguridad",         cant:1,  ubicacion:"Planta FCA"},
    {puesto:"Supervisor Turno A",         cant:1,  ubicacion:"Planta FCA"},
    {puesto:"Supervisor Turno B",         cant:1,  ubicacion:"Planta FCA"},
    {puesto:"Operador Monitoreo CCTV",    cant:3,  ubicacion:"Planta FCA"},
    {puesto:"Oficial Seguridad Turno A",  cant:10, ubicacion:"Planta FCA"},
    {puesto:"Oficial Seguridad Turno B",  cant:9,  ubicacion:"Planta FCA"},
    {puesto:"Control Accesos Vehicular",  cant:2,  ubicacion:"Planta FCA"}
];  // total: 27

var PERSONAL_ESTANCIAS = [
    {puesto:"Supervisor Operativo",       cant:1,  ubicacion:"Bodega Estancias"},
    {puesto:"Oficial Seguridad Turno A",  cant:3,  ubicacion:"Bodega Estancias"},
    {puesto:"Oficial Seguridad Turno B",  cant:3,  ubicacion:"Bodega Estancias"}
];  // total: 7

var PERSONAL_NORTE = [
    {puesto:"Supervisor Operativo",       cant:1,  ubicacion:"Bodega Tierra Norte"},
    {puesto:"Oficial Seguridad Turno A",  cant:3,  ubicacion:"Bodega Tierra Norte"},
    {puesto:"Oficial Seguridad con Cuatrimoto", cant:2, ubicacion:"Bodega Tierra Norte"}
];  // total: 6

// ── COSTOS REALES MENSUALES (datos Excel) ───────────────────────────────────
var NOMINA_MENSUAL  = 764318.45;   // nómina total 40 personas (IMSS incluido)
var TRASLADOS_MES   = 39143.50;    // traslados operativos
var EPP_MES         = 17000.00;    // equipo de protección personal recurrente
var CUATRIMOTOS_MES = 15280.00;    // renta/cuota cuatrimotos operativas

// ── COSTOS ÚNICOS MES 1 ──────────────────────────────────────────────────────
var DC3_UNICO        = 118800.00;  // constancias DC-3 y capacitación
var UNIFORMES_UNICO  = 115200.00;  // uniformes completos 40 personas
var HERRAMIENTAS_UNI = 73860.00;   // herramientas y equipo inicial

// ── TOTALES CALCULADOS ───────────────────────────────────────────────────────
var COSTO_REC_MES  = NOMINA_MENSUAL + TRASLADOS_MES + EPP_MES + CUATRIMOTOS_MES;
// = $835,741.95 / mes (meses 2–12)

var COSTO_MES1     = COSTO_REC_MES + DC3_UNICO + UNIFORMES_UNICO + HERRAMIENTAS_UNI;
// = $1,143,601.95 (mes 1, incluye inversión inicial)

var COSTO_12_MESES = 10504219.52;  // total costo directo 12 meses (validado Excel)
var ADMIN_5PCT     = 525210.98;    // administración 5%
var UTIL_20PCT     = 2100843.90;   // utilidad 20%

var PRECIO_SIN_IVA = 13130274.40;  // precio venta sin IVA 12 meses
var IVA_16PCT      = 2100843.90;   // IVA 16%
var PRECIO_CON_IVA = 15231118.30;  // PRECIO TOTAL CON IVA 12 MESES

// Precios mensuales
var PRECIO_MES1_SINIVA  = COSTO_MES1 * (PRECIO_SIN_IVA / COSTO_12_MESES);
// ≈ $1,429,502
var PRECIO_MES1_CONIVA  = PRECIO_MES1_SINIVA * 1.16;
// ≈ $1,658,223
var PRECIO_REC_SINIVA   = COSTO_REC_MES * (PRECIO_SIN_IVA / COSTO_12_MESES);
// ≈ $1,044,677
var PRECIO_REC_CONIVA   = PRECIO_REC_SINIVA * 1.16;
// ≈ $1,211,826

// Desglose por ubicación (proporcional headcount)
var PCT_PLANTA    = 27/40;  // 67.5%
var PCT_ESTANCIAS = 7/40;   // 17.5%
var PCT_NORTE     = 6/40;   // 15.0%

// ── DATOS RFC / REPSE ────────────────────────────────────────────────────────
var RFC_MICSA   = "MIC2301268S5";
var REPSE_NUM   = "282364";

// ═══════════════════════════════════════════════════════════════════════════════
// UTILIDADES
// ═══════════════════════════════════════════════════════════════════════════════
function fmt(n){
    var fixed=n.toFixed(2);
    return "$"+fixed.replace(/\B(?=(\d{3})+(?!\d))/g,",");
}
function fmtK(n){
    return "$"+(n/1000).toFixed(1)+"K";
}

// ─── PRIMITIVAS DE DIBUJO ────────────────────────────────────────────────────
function R(doc,x,y,w,h,f,s,sw){
    var r=doc.pathItems.rectangle(-y,x,w,h);
    r.filled=true; r.fillColor=f;
    if(s){r.stroked=true;r.strokeColor=s;r.strokeWidth=sw||0.5;}else r.stroked=false;
    return r;
}
function T(doc,s,x,y,sz,col,font,w,al){
    var t=doc.textFrames.add();
    t.contents=s; t.left=x; t.top=-y;
    if(w)t.width=w;
    t.textRange.characterAttributes.size=sz;
    t.textRange.characterAttributes.fillColor=col;
    try{t.textRange.characterAttributes.textFont=app.textFonts.getByName(font);}catch(e){}
    if(al==="C")t.textRange.paragraphAttributes.justification=Justification.CENTER;
    if(al==="R")t.textRange.paragraphAttributes.justification=Justification.RIGHT;
    return t;
}
function L(doc,x1,y1,x2,y2,col,w){
    var l=doc.pathItems.add();
    l.setEntirePath([[x1,-y1],[x2,-y2]]);
    l.stroked=true; l.strokeColor=col; l.strokeWidth=w||0.5; l.filled=false;
    return l;
}
function logo(doc,x,y,w){
    try{
        var f=new File(LOGO_FILE);
        if(!f.exists)return;
        var p=doc.placedItems.add();
        p.file=f; var sc=w/p.width;
        p.width=w; p.height=p.height*sc;
        p.left=x; p.top=-y;
    }catch(e){}
}

// ─── COMPONENTES DE LAYOUT ───────────────────────────────────────────────────
var TW = A4W - MRG*2;

function mkHeader(doc,O,titulo){
    R(doc,O*A4W,0,A4W,15*MM,C_NEGRO);
    R(doc,O*A4W,15*MM,A4W,1*MM,C_ROJO);
    logo(doc,O*A4W+MRG,1.5*MM,16*MM);
    T(doc,"MICSA SAFETY DIVISION",O*A4W+MRG+18*MM,4*MM,8,C_BLANCO,FB,A4W*0.5);
    T(doc,titulo,O*A4W+MRG+18*MM,10*MM,6.5,C_PLATA,FL,A4W*0.55);
    T(doc,FOLIO+" | "+FECHA,(O+1)*A4W-MRG-90,7*MM,6.5,C_PLATA,FL,88,"R");
}
function mkFooter(doc,O,pg){
    R(doc,O*A4W,A4H-12*MM,A4W,12*MM,C_NEGRO);
    R(doc,O*A4W,A4H-12*MM,A4W,0.8*MM,C_ROJO);
    T(doc,"MICSA Safety Division  |  RFC: "+RFC_MICSA+"  |  REPSE: "+REPSE_NUM+"  |  contacto@micsa.mx",
        O*A4W+MRG,A4H-7*MM,5.5,C_GRIS,FL,A4W-MRG*2-20);
    T(doc,pg+" / 6",(O+1)*A4W-MRG-18,A4H-7*MM,7,C_PLATA,FS,20);
}
function SEC(doc,O,y,t){
    R(doc,O*A4W+MRG-2,y-1,4,10*MM,C_ROJO);
    T(doc,t,O*A4W+MRG+5,y,11,C_NEGRO,FB,TW-5);
    y+=7*MM;
    L(doc,O*A4W+MRG,(O+1)*A4W-MRG,y,y,C_PLATA,0.6);
    return y+5*MM;
}
function tabla(doc,O,y,hdrs,rows,colW){
    var x=O*A4W+MRG, rh=10*MM, cx, cw, r, c, bg;
    cx=x;
    for(var h=0;h<hdrs.length;h++){
        cw=colW[h];
        R(doc,cx,y,cw,rh,C_NEGRO,C_PLATA,0.3);
        T(doc,hdrs[h],cx+2*MM,y+3*MM,6.5,C_BLANCO,FB,cw-4*MM);
        cx+=cw;
    }
    y+=rh;
    for(r=0;r<rows.length;r++){
        cx=x; bg=r%2===0?C_GRISC:C_BLANCO;
        for(c=0;c<rows[r].length;c++){
            cw=colW[c];
            R(doc,cx,y,cw,rh,bg,C_PLATA,0.3);
            T(doc,rows[r][c],cx+2*MM,y+3*MM,6.5,C_GRIS,FR,cw-4*MM);
            cx+=cw;
        }
        y+=rh;
    }
    return y;
}
function filaTotal(doc,O,y,etiq,val){
    R(doc,O*A4W+MRG,y,TW,10*MM,C_NEGRO,C_PLATA,0.3);
    T(doc,etiq,O*A4W+MRG+2*MM,y+3*MM,7,C_BLANCO,FB,TW*0.65);
    T(doc,val,O*A4W+A4W-MRG-2*MM,y+3*MM,8,C_ROJO,FB,TW*0.30,"R");
    return y+10*MM;
}

// ═══════════════════════════════════════════════════════════════════════════════
// ARTBOARD 1 — PORTADA
// ═══════════════════════════════════════════════════════════════════════════════
function portada(doc){
    var O=0;
    R(doc,0,0,A4W,A4H,C_BLANCO);

    // Bloque superior negro
    R(doc,0,0,A4W,A4H*0.58,C_NEGRO);

    // Franja lateral roja
    R(doc,0,0,7*MM,A4H,C_ROJO);
    R(doc,0,A4H*0.58,A4W,1.5*MM,C_ROJO);

    // Logotipo centrado
    logo(doc,A4W/2-25*MM,A4H*0.06,50*MM);

    // Título principal
    T(doc,"PROPUESTA INTEGRAL DE",12*MM,A4H*0.29,17,C_BLANCO,FB,A4W-18*MM);
    T(doc,"SEGURIDAD OPERATIVA",12*MM,A4H*0.29+22,17,C_BLANCO,FB,A4W-18*MM);
    T(doc,"FCA — FREIGHTCAR MÉXICO",12*MM,A4H*0.29+42,11,C_ROJO,FS,A4W-18*MM);

    // Línea separadora roja
    R(doc,12*MM,A4H*0.29+58,80*MM,1.5*MM,C_ROJO);

    // Caja de datos
    R(doc,10*MM,A4H*0.62,A4W-20*MM,A4H*0.27,C_GRISC);
    R(doc,10*MM,A4H*0.62,2.5*MM,A4H*0.27,C_NEGRO);

    T(doc,"Folio:",   15*MM, A4H*0.64,    7,C_GRIS,FL,35);
    T(doc,FOLIO,      15*MM, A4H*0.64+11, 10,C_NEGRO,FB,160);
    T(doc,"Fecha:",   15*MM, A4H*0.64+26, 7,C_GRIS,FL,35);
    T(doc,FECHA,      15*MM, A4H*0.64+37, 9,C_NEGRO,FS,100);
    T(doc,"Vigencia:",15*MM, A4H*0.64+51, 7,C_GRIS,FL,50);
    T(doc,"15 días naturales",15*MM,A4H*0.64+61,9,C_NEGRO,FS,130);

    T(doc,"MICSA SAFETY DIVISION",A4W-MRG-165,A4H*0.64+11,9,C_NEGRO,FB,163);
    T(doc,"RFC: "+RFC_MICSA,A4W-MRG-165,A4H*0.64+25,7.5,C_GRIS,FR,163);
    T(doc,"REPSE: "+REPSE_NUM,A4W-MRG-165,A4H*0.64+37,7.5,C_GRIS,FR,163);
    T(doc,"contacto@micsa.mx  |  800-MICSA-01",A4W-MRG-165,A4H*0.64+49,7,C_GRIS,FL,163);

    // Footer
    R(doc,0,A4H*0.91,A4W,A4H*0.09,C_NEGRO);
    T(doc,'"No cuidamos instalaciones. Controlamos operaciones."',
        10*MM,A4H*0.935,8,C_PLATA,FL,A4W-14*MM,"C");
    T(doc,"CONFIDENCIAL — Propietario MICSA Safety Division 2026",
        10*MM,A4H*0.963,6.5,C_GRIS,FL,A4W-14*MM,"C");
}

// ═══════════════════════════════════════════════════════════════════════════════
// ARTBOARD 2 — ALCANCE Y RIESGO
// ═══════════════════════════════════════════════════════════════════════════════
function alcance(doc){
    var O=1;
    R(doc,O*A4W,0,A4W,A4H,C_BLANCO);
    mkHeader(doc,O,"ANÁLISIS DE RIESGO Y ALCANCE DEL SERVICIO");
    mkFooter(doc,O,"2");
    var y=19*MM;

    y=SEC(doc,O,y,"1. RIESGO OPERATIVO NO CUBIERTO");
    var cW1=[TW*0.58,TW*0.42];
    y=tabla(doc,O,y,["RIESGO IDENTIFICADO","IMPACTO ESTIMADO"],[
        ["Pérdida patrimonial por turno sin cobertura","$85,000–$320,000 MXN / evento"],
        ["Paro de línea por incidente de seguridad","$420,000–$1,200,000 MXN / hora"],
        ["Multa STPS incumplimiento NOM-030","$80,000–$240,000 MXN"],
        ["Responsabilidad civil patronal no cubierta","Litigio + fianza + daño imagen"],
        ["Robo componentes sin cobertura nocturna","$180,000–$600,000 MXN / turno"]
    ],cW1);

    y+=10*MM;
    y=SEC(doc,O,y,"2. ALCANCE DE LA SOLUCIÓN — 3 UBICACIONES");
    var cW2=[TW*0.22,TW*0.12,TW*0.36,TW*0.30];
    y=tabla(doc,O,y,["UBICACIÓN","PERS.","COBERTURA","SERVICIO"],[
        ["Planta FCA Saltillo","27","24/7 · Turnos A/B 12h","Control acceso vehicular/peatonal"],
        ["Bodega Estancias","7","24/7 · Turnos A/B 12h","Registro digital ingresos/salidas"],
        ["Bodega Tierra Norte","6","24/7 + Cuatrimoto","Rondines QR y bitácora digital"],
        ["Centro Monitoreo","incl.","3 operadores CCTV 24/7","Reporte incidentes tiempo real"],
        ["Gestión KPI Global","incl.","Jefe Seg. + Supervisor","Dashboard mensual indicadores"]
    ],cW2);

    y+=10*MM;
    y=SEC(doc,O,y,"3. MARCO NORMATIVO Y ACREDITACIÓN");
    var cW3=[TW*0.20,TW*0.52,TW*0.28];
    tabla(doc,O,y,["NORMA","DESCRIPCIÓN","ESTATUS"],[
        ["NOM-030-STPS","Servicios preventivos de seguridad y salud","✓ Cumplimiento total"],
        ["NOM-017-STPS","Equipo de Protección Personal","✓ Cumplimiento total"],
        ["REPSE "+REPSE_NUM,"Registro Prestadoras Servicios Especializados","✓ Vigente 2026"],
        ["RFC "+RFC_MICSA,"Constancia fiscal SAT — Situación Regular","✓ Al corriente"]
    ],cW3);
}

// ═══════════════════════════════════════════════════════════════════════════════
// ARTBOARD 3 — ESTRUCTURA DE PERSONAL Y NÓMINA
// ═══════════════════════════════════════════════════════════════════════════════
function nomina(doc){
    var O=2;
    R(doc,O*A4W,0,A4W,A4H,C_BLANCO);
    mkHeader(doc,O,"ESTRUCTURA DE PERSONAL — INVERSIÓN EN CAPITAL HUMANO");
    mkFooter(doc,O,"3");
    var y=19*MM;

    // Tabla por ubicación
    y=SEC(doc,O,y,"3. DISTRIBUCIÓN OPERATIVA POR UBICACIÓN");
    T(doc,"Total: 40 personas activas — nómina IMSS al corriente",
        O*A4W+MRG,y,7,C_GRIS,FL,TW); y+=9*MM;

    var cW=[TW*0.35,TW*0.10,TW*0.18,TW*0.18,TW*0.19];
    var rowsNom=[
        ["Planta FCA Saltillo","27","Supervisores + Oficiales + CCTV","Cobertura 24/7",
         fmt(NOMINA_MENSUAL*PCT_PLANTA)],
        ["Bodega Estancias","7","Supervisor + Oficiales 2 turnos","Cobertura 24/7",
         fmt(NOMINA_MENSUAL*PCT_ESTANCIAS)],
        ["Bodega Tierra Norte","6","Supervisor + Oficiales + Cuatrimoto","Cobertura 24/7",
         fmt(NOMINA_MENSUAL*PCT_NORTE)],
        ["SUBTOTAL NÓMINA MENSUAL","40","—","—",fmt(NOMINA_MENSUAL)]
    ];
    y=tabla(doc,O,y,["UBICACIÓN","PERS.","PERFIL","TURNO","COSTO MES"],rowsNom,cW);

    y+=10*MM;
    y=SEC(doc,O,y,"PERFIL MÍNIMO DE CONTRATACIÓN");
    var cW2=[TW*0.28,TW*0.72];
    y=tabla(doc,O,y,["CRITERIO","ESPECIFICACIÓN"],[
        ["Escolaridad","Secundaria concluida mínima — bachillerato preferente"],
        ["Experiencia","Mínimo 1 año en seguridad industrial o patrimonial certificada"],
        ["Evaluación","Psicométrica ≥80%: honestidad + estabilidad emocional + criterio"],
        ["Documentación","Sin antecedentes penales · Cartilla militar · Domicilio verificable"],
        ["Capacitación","DC-3 vigente — Protocolo MICSA Safety (16 horas inducción)"],
        ["Sustitución","Garantía reemplazo < 4 horas ante ausencia no programada"]
    ],cW2);

    y+=10*MM;
    y=SEC(doc,O,y,"KPI'S DE DESEMPEÑO GARANTIZADOS");
    var cW3=[TW*0.30,TW*0.18,TW*0.52];
    tabla(doc,O,y,["KPI","META","CRITERIO DE CONTROL"],[
        ["Puntualidad","≥ 97%","3 faltas en 30 días = baja automática + reemplazo inmediato"],
        ["Bitácoras completas","100%","Sin firma = turno incompleto + sanción documentada"],
        ["Rondines cumplidos","≥ 95%","Registro QR o firma por punto de control en ruta"],
        ["Incidentes reportados","100%","Notificación inmediata + informe ≤ 2 horas hábiles"],
        ["Rotación mensual","≤ 5%","Evaluación semanal + expediente individual activo"]
    ],cW3);
}

// ═══════════════════════════════════════════════════════════════════════════════
// ARTBOARD 4 — EQUIPAMIENTO Y LOGÍSTICA
// ═══════════════════════════════════════════════════════════════════════════════
function equipamiento(doc){
    var O=3;
    R(doc,O*A4W,0,A4W,A4H,C_BLANCO);
    mkHeader(doc,O,"EQUIPAMIENTO, TRASLADOS Y LOGÍSTICA OPERATIVA");
    mkFooter(doc,O,"4");
    var y=19*MM;

    y=SEC(doc,O,y,"4. DESGLOSE DE COSTOS OPERATIVOS");
    T(doc,"Inversión inicial Mes 1 vs. costo mensual recurrente Meses 2–12",
        O*A4W+MRG,y,7,C_GRIS,FL,TW); y+=9*MM;

    var cW=[TW*0.42,TW*0.16,TW*0.20,TW*0.22];
    y=tabla(doc,O,y,["CONCEPTO","FRECUENCIA","MONTO","TIPO"],[
        // Costos recurrentes
        ["Nómina operativa — 40 personas (IMSS inc.)","Mensual",fmt(NOMINA_MENSUAL),"Recurrente"],
        ["Traslados y logística operativa","Mensual",fmt(TRASLADOS_MES),"Recurrente"],
        ["EPP rotación mensual (protección personal)","Mensual",fmt(EPP_MES),"Recurrente"],
        ["Cuatrimotos operativas (renta/cuota)","Mensual",fmt(CUATRIMOTOS_MES),"Recurrente"],
        // Costos únicos
        ["Constancias DC-3 + capacitación inicial 40 pers.","Mes 1",fmt(DC3_UNICO),"Único"],
        ["Uniformes completos — 40 personas","Mes 1",fmt(UNIFORMES_UNICO),"Único"],
        ["Herramientas y equipo base inicial","Mes 1",fmt(HERRAMIENTAS_UNI),"Único"]
    ],cW);

    y+=8*MM;
    // Cuadros resumen mes
    var bw=(TW-6*MM)/2, bh=22*MM;
    R(doc,O*A4W+MRG,y,bw,bh,C_NEGRO);
    R(doc,O*A4W+MRG+bw+6*MM,y,bw,bh,C_GRIS);
    T(doc,"COSTO DIRECTO MES 1",O*A4W+MRG+3*MM,y+4*MM,7,C_PLATA,FL,bw-6*MM,"C");
    T(doc,fmt(COSTO_MES1),O*A4W+MRG+3*MM,y+10*MM,10,C_ROJO,FB,bw-6*MM,"C");
    T(doc,"(inversión inicial + recurrente)",O*A4W+MRG+3*MM,y+18*MM,6,C_PLATA,FL,bw-6*MM,"C");
    T(doc,"COSTO DIRECTO MES 2–12",O*A4W+MRG+bw+9*MM,y+4*MM,7,C_BLANCO,FL,bw-6*MM,"C");
    T(doc,fmt(COSTO_REC_MES),O*A4W+MRG+bw+9*MM,y+10*MM,10,C_BLANCO,FB,bw-6*MM,"C");
    T(doc,"(solo costos recurrentes)",O*A4W+MRG+bw+9*MM,y+18*MM,6,C_PLATA,FL,bw-6*MM,"C");
    y+=bh+12*MM;

    y=SEC(doc,O,y,"EPP INDIVIDUAL — DOTACIÓN POR GUARDIA");
    var cW2=[TW*0.48,TW*0.22,TW*0.30];
    tabla(doc,O,y,["EQUIPO DE PROTECCIÓN","DOTACIÓN","NORMA APLICABLE"],[
        ["Uniforme completo (pantalón táctico + camisa ML)","2 juegos","NOM identificación MICSA"],
        ["Casco de seguridad clase E con barbiquejo","1 pieza","NOM-115-STPS-2009"],
        ["Calzado de seguridad punta de acero","1 par","NOM-113-STPS-2009"],
        ["Chaleco reflectante clase II","1 pieza","ANSI/ISEA 107-2020"],
        ["Guantes de protección media caña","2 pares","NOM-138-STPS-2012"],
        ["Gafas de seguridad policarbonato","1 pieza","ANSI Z87.1"],
        ["Radio portátil digital + batería cargador","1 equipo","MICSA operativo"]
    ],cW2);
}

// ═══════════════════════════════════════════════════════════════════════════════
// ARTBOARD 5 — RESUMEN FINANCIERO
// ═══════════════════════════════════════════════════════════════════════════════
function financiero(doc){
    var O=4;
    R(doc,O*A4W,0,A4W,A4H,C_BLANCO);
    mkHeader(doc,O,"RESUMEN FINANCIERO — INVERSIÓN ANUAL TOTAL");
    mkFooter(doc,O,"5");
    var y=19*MM;

    y=SEC(doc,O,y,"5. ESTRUCTURA DE PRECIOS — PROYECTO ANUAL 12 MESES");
    T(doc,"Desglose completo: costo directo + administración + utilidad + IVA 16%",
        O*A4W+MRG,y,7,C_GRIS,FL,TW); y+=9*MM;

    var cW=[TW*0.44,TW*0.18,TW*0.14,TW*0.24];
    y=tabla(doc,O,y,["CONCEPTO","TIPO","MESES","MONTO"],[
        ["Nómina operativa 40 personas",  "Recurrente","12",fmt(NOMINA_MENSUAL*12)],
        ["Traslados y logística",          "Recurrente","12",fmt(TRASLADOS_MES*12)],
        ["EPP mensual",                    "Recurrente","12",fmt(EPP_MES*12)],
        ["Cuatrimotos",                    "Recurrente","12",fmt(CUATRIMOTOS_MES*12)],
        ["Constancias DC-3 / Capacitación","Único","1",  fmt(DC3_UNICO)],
        ["Uniformes completos 40 personas","Único","1",  fmt(UNIFORMES_UNICO)],
        ["Herramientas y equipo base",     "Único","1",  fmt(HERRAMIENTAS_UNI)]
    ],cW);
    y+=2*MM;
    // Totales escalados
    y=filaTotal(doc,O,y,"COSTO DIRECTO 12 MESES",fmt(COSTO_12_MESES));
    R(doc,O*A4W+MRG,y,TW,9*MM,C_GRISC,C_PLATA,0.3);
    T(doc,"  Administración 5%",O*A4W+MRG+2*MM,y+2.5*MM,7,C_GRIS,FR,TW*0.65);
    T(doc,fmt(ADMIN_5PCT),(O+1)*A4W-MRG-2*MM,y+2.5*MM,7,C_GRIS,FR,TW*0.30,"R");
    y+=9*MM;
    R(doc,O*A4W+MRG,y,TW,9*MM,C_GRISC,C_PLATA,0.3);
    T(doc,"  Utilidad 20%",O*A4W+MRG+2*MM,y+2.5*MM,7,C_GRIS,FR,TW*0.65);
    T(doc,fmt(UTIL_20PCT),(O+1)*A4W-MRG-2*MM,y+2.5*MM,7,C_GRIS,FR,TW*0.30,"R");
    y+=9*MM;
    y=filaTotal(doc,O,y,"PRECIO VENTA SIN IVA — 12 MESES",fmt(PRECIO_SIN_IVA));
    R(doc,O*A4W+MRG,y,TW,9*MM,C_GRISC,C_PLATA,0.3);
    T(doc,"  IVA 16%",O*A4W+MRG+2*MM,y+2.5*MM,7,C_GRIS,FR,TW*0.65);
    T(doc,fmt(IVA_16PCT),(O+1)*A4W-MRG-2*MM,y+2.5*MM,7,C_GRIS,FR,TW*0.30,"R");
    y+=9*MM;

    // GRAN TOTAL
    R(doc,O*A4W+MRG,y,TW,14*MM,C_ROJO);
    T(doc,"PRECIO TOTAL CON IVA — 12 MESES",O*A4W+MRG+3*MM,y+4*MM,8.5,C_BLANCO,FB,TW*0.60);
    T(doc,fmt(PRECIO_CON_IVA),(O+1)*A4W-MRG-3*MM,y+4*MM,10,C_BLANCO,FB,TW*0.36,"R");
    y+=18*MM;

    y=SEC(doc,O,y,"PRECIOS MENSUALES POR UBICACIÓN (referencia)");
    var cW2=[TW*0.30,TW*0.10,TW*0.20,TW*0.20,TW*0.20];
    tabla(doc,O,y,["UBICACIÓN","PERS.","MES 1 s/IVA","MES 2-12 s/IVA","TOTAL 12m c/IVA"],[
        ["Planta FCA Saltillo","27",
         fmt(PRECIO_MES1_SINIVA*PCT_PLANTA),
         fmt(PRECIO_REC_SINIVA*PCT_PLANTA),
         fmt(PRECIO_CON_IVA*PCT_PLANTA)],
        ["Bodega Estancias","7",
         fmt(PRECIO_MES1_SINIVA*PCT_ESTANCIAS),
         fmt(PRECIO_REC_SINIVA*PCT_ESTANCIAS),
         fmt(PRECIO_CON_IVA*PCT_ESTANCIAS)],
        ["Bodega Tierra Norte","6",
         fmt(PRECIO_MES1_SINIVA*PCT_NORTE),
         fmt(PRECIO_REC_SINIVA*PCT_NORTE),
         fmt(PRECIO_CON_IVA*PCT_NORTE)],
        ["GLOBAL 40 PERSONAS","40",
         fmt(PRECIO_MES1_SINIVA),
         fmt(PRECIO_REC_SINIVA),
         fmt(PRECIO_CON_IVA)]
    ],cW2);
}

// ═══════════════════════════════════════════════════════════════════════════════
// ARTBOARD 6 — CONDICIONES DE FORMALIZACIÓN
// ═══════════════════════════════════════════════════════════════════════════════
function condiciones(doc){
    var O=5;
    R(doc,O*A4W,0,A4W,A4H,C_BLANCO);
    mkHeader(doc,O,"CONDICIONES DE FORMALIZACIÓN Y CIERRE");
    mkFooter(doc,O,"6");
    var y=19*MM;

    y=SEC(doc,O,y,"6. CONDICIONES COMERCIALES");
    var cW=[TW*0.33,TW*0.67];
    y=tabla(doc,O,y,["CONDICIÓN","DETALLE"],[
        ["Vigencia de propuesta","15 días naturales desde fecha de emisión"],
        ["Requisito de inicio","Firma de contrato + liquidación primer mes"],
        ["Inicio de operaciones","5 días hábiles tras formalización y pago"],
        ["Ajuste de precios","Semestral conforme índices STPS / INPC vigentes"],
        ["Terminación anticipada","Aviso por escrito 30 días naturales previos"],
        ["Confidencialidad","NDA firmado — información protegida 5 años"]
    ],cW);

    y+=10*MM;
    y=SEC(doc,O,y,"DOCUMENTACIÓN REQUERIDA");
    var cW2=[TW*0.48,TW*0.32,TW*0.20];
    y=tabla(doc,O,y,["DOCUMENTO","RESPONSABLE","ESTATUS"],[
        ["Contrato Servicios Especializados (REPSE "+REPSE_NUM+")","MICSA + Cliente","Obligatorio"],
        ["Carta aceptación propuesta — firma representante legal","Rep. Legal Cliente","Obligatorio"],
        ["Alta sistema de proveedores cliente","Depto. Compras","Obligatorio"],
        ["Orden de compra (PO) autorizada","Finanzas / Compras","Obligatorio"],
        ["Constancia fiscal SAT ambas partes","Administración","Obligatorio"],
        ["Contacto operativo de planta asignado","Gerencia Seguridad FCA","Obligatorio"]
    ],cW2);

    y+=12*MM;
    // Bloque de cierre
    R(doc,O*A4W+MRG,y,TW,52*MM,C_NEGRO);
    R(doc,O*A4W+MRG,y,TW,1.5*MM,C_ROJO);
    R(doc,O*A4W+MRG,y+52*MM,TW,1.5*MM,C_ROJO);
    T(doc,"MICSA SAFETY DIVISION",O*A4W+MRG,y+8*MM,10,C_BLANCO,FB,TW,"C");
    T(doc,"Tu socio estratégico en seguridad industrial",O*A4W+MRG,y+18*MM,7.5,C_PLATA,FL,TW,"C");
    T(doc,'"No cuidamos instalaciones. Controlamos operaciones."',
        O*A4W+MRG,y+27*MM,8,C_ROJO,FS,TW,"C");
    T(doc,"contacto@micsa.mx  ·  800-MICSA-01  ·  www.micsa.mx  ·  RFC: "+RFC_MICSA,
        O*A4W+MRG,y+38*MM,6.5,C_PLATA,FL,TW,"C");

    y+=58*MM;
    // Líneas de firma
    var fw=TW*0.43;
    L(doc,O*A4W+MRG,O*A4W+MRG+fw,y+20,y+20,C_GRIS,0.8);
    L(doc,O*A4W+A4W-MRG-fw,O*A4W+A4W-MRG,y+20,y+20,C_GRIS,0.8);
    T(doc,"Director General — MICSA Safety Division",O*A4W+MRG,y+25,7,C_GRIS,FL,fw);
    T(doc,"Representante Legal — FCA / FreightCar México",O*A4W+A4W-MRG-fw,y+25,7,C_GRIS,FL,fw);
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN — CREAR DOCUMENTO Y ARTBOARDS
// ═══════════════════════════════════════════════════════════════════════════════
function main(){
    var doc = app.documents.add(
        DocumentColorSpace.CMYK,
        A4W * 6,  // 6 artboards en fila
        A4H,
        DocumentArtboardLayout.GridByRow,
        0,
        6
    );
    doc.artboards[0].artboardRect = [0,       0, A4W,       -A4H];
    doc.artboards[1].artboardRect = [A4W,     0, A4W*2,     -A4H];
    doc.artboards[2].artboardRect = [A4W*2,   0, A4W*3,     -A4H];
    doc.artboards[3].artboardRect = [A4W*3,   0, A4W*4,     -A4H];
    doc.artboards[4].artboardRect = [A4W*4,   0, A4W*5,     -A4H];
    doc.artboards[5].artboardRect = [A4W*5,   0, A4W*6,     -A4H];

    var nms=["01 Portada","02 Alcance","03 Personal","04 Equipamiento","05 Financiero","06 Condiciones"];
    for(var i=0;i<6;i++) doc.artboards[i].name=nms[i];

    app.coordinateSystem = CoordinateSystem.ARTBOARDCOORDINATESYSTEM;

    portada(doc);
    alcance(doc);
    nomina(doc);
    equipamiento(doc);
    financiero(doc);
    condiciones(doc);

    // Guardar como .ai
    var saveFile = new File(BASE_PATH + "COTIZACION_FCA_v2.ai");
    var opts = new IllustratorSaveOptions();
    opts.compatibility = Compatibility.ILLUSTRATOR24;
    opts.compressed = true;
    doc.saveAs(saveFile, opts);

    alert("COTIZACION_FCA_v2.ai generada correctamente.\n\nFolio: "+FOLIO+
          "\nPersonal: 40 personas (27+7+6)\nTotal 12m c/IVA: "+fmt(PRECIO_CON_IVA)+
          "\n\nArchivo: "+saveFile.fsName);
}

main();
