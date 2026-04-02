// Polyfill Array.map para ExtendScript ES3
if(!Array.prototype.map){Array.prototype.map=function(fn){var r=[];for(var i=0;i<this.length;i++)r.push(fn(this[i],i,this));return r;};}

/**
 * MICSA SAFETY DIVISION — Cotización Branded FCA
 * ExtendScript Adobe Illustrator — 6 artboards A4 — CMYK
 */
var BASE_PATH = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\MICSA_Safety\\";
var LOGO_FILE = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\adobe\\ID-5608-20260329T034218Z-1-001\\ID-5608\\LOGO 7\\editable\\editable.ai";
var OUT_PATH  = BASE_PATH + "PDFs\\";

var MM=2.834645669, A4W=210*MM, A4H=297*MM, MRG=20*MM;

function cmyk(c,m,y,k){var col=new CMYKColor();col.cyan=c;col.magenta=m;col.yellow=y;col.black=k;return col;}
var C_NEGRO  =cmyk(60,40,40,100);
var C_PLATA  =cmyk(0,0,0,30);
var C_GRIS   =cmyk(0,0,0,55);
var C_GRISC  =cmyk(0,0,0,10);
var C_BLANCO =cmyk(0,0,0,0);
var C_AMARILLO=cmyk(0,15,100,0);

var FB="Montserrat-Bold", FS="Montserrat-SemiBold", FR="OpenSans-Regular", FL="OpenSans-Light";
var HOY=new Date();
var FECHA=HOY.getDate()+"/"+(HOY.getMonth()+1)+"/"+HOY.getFullYear();
var FOLIO="PROP-MSF-FCA-"+HOY.getFullYear()+"-001";

function R(doc,x,y,w,h,f,s,sw){var r=doc.pathItems.rectangle(-y,x,w,h);r.filled=true;r.fillColor=f;if(s){r.stroked=true;r.strokeColor=s;r.strokeWidth=sw||0.5;}else r.stroked=false;return r;}
function T(doc,s,x,y,sz,col,font,w,al){var t=doc.textFrames.add();t.contents=s;t.left=x;t.top=-y;if(w)t.width=w;t.textRange.characterAttributes.size=sz;t.textRange.characterAttributes.fillColor=col;try{t.textRange.characterAttributes.textFont=app.textFonts.getByName(font);}catch(e){}if(al==="C")t.textRange.paragraphAttributes.justification=Justification.CENTER;if(al==="R")t.textRange.paragraphAttributes.justification=Justification.RIGHT;return t;}
function L(doc,x1,y1,x2,y2,col,w){var l=doc.pathItems.add();l.setEntirePath([[x1,-y1],[x2,-y2]]);l.stroked=true;l.strokeColor=col;l.strokeWidth=w||0.5;l.filled=false;return l;}
function logo(doc,x,y,w){try{var f=new File(LOGO_FILE);if(!f.exists)return;var p=doc.placedItems.add();p.file=f;var sc=w/p.width;p.width=w;p.height=p.height*sc;p.left=x;p.top=-y;}catch(e){}}

function mkHeader(doc,O,titulo){
    R(doc,O*A4W,0,A4W,15*MM,C_NEGRO);R(doc,O*A4W,15*MM,A4W,1*MM,C_PLATA);
    logo(doc,O*A4W+MRG,1.5*MM,16*MM);
    T(doc,"MICSA SAFETY DIVISION",O*A4W+MRG+18*MM,4*MM,8,C_BLANCO,FB,A4W*0.5);
    T(doc,titulo,O*A4W+MRG+18*MM,10*MM,6.5,C_PLATA,FL,A4W*0.55);
    T(doc,FOLIO,(O+1)*A4W-MRG-80,10*MM,6,C_PLATA,FL,78);
}
function mkFooter(doc,O,pg){
    R(doc,O*A4W,A4H-12*MM,A4W,12*MM,C_NEGRO);R(doc,O*A4W,A4H-12*MM,A4W,0.7*MM,C_PLATA);
    T(doc,"MICSA Safety Division  |  Seguridad Industrial  |  contacto@micsa.mx  |  800-MICSA-01",O*A4W+MRG,A4H-7*MM,6,C_GRIS,FL,A4W-MRG*2-20);
    T(doc,pg+" / 6",(O+1)*A4W-MRG-18,A4H-7*MM,7,C_PLATA,FS,20);
}

var TW=A4W-MRG*2;
function tabla(doc,O,y,hdrs,rows,colW){
    var x=O*A4W+MRG,rh=10*MM,cx,cw,r,c,bg;
    cx=x;
    for(var h=0;h<hdrs.length;h++){cw=colW[h];R(doc,cx,y,cw,rh,C_NEGRO,C_PLATA,0.3);T(doc,hdrs[h],cx+2*MM,y+3*MM,6.5,C_BLANCO,FB,cw-4*MM);cx+=cw;}
    y+=rh;
    for(r=0;r<rows.length;r++){cx=x;bg=r%2===0?C_GRISC:C_BLANCO;for(c=0;c<rows[r].length;c++){cw=colW[c];R(doc,cx,y,cw,rh,bg,C_PLATA,0.3);T(doc,rows[r][c],cx+2*MM,y+3*MM,6.5,C_GRIS,FR,cw-4*MM);cx+=cw;}y+=rh;}
    return y;
}
function SEC(doc,O,y,t){T(doc,t,O*A4W+MRG,y,11,C_NEGRO,FB,TW);y+=7*MM;L(doc,O*A4W+MRG,y,(O+1)*A4W-MRG,y,C_PLATA,0.8);return y+5*MM;}
function SPACE(){return 6*MM;}

// ─── DATOS NÓMINA ────────────────────────────────────────────────────────────
var PERSONAL=[
    {puesto:"Jefe de Seguridad",cant:1,sueldo:6000,cp:3933},
    {puesto:"Supervisor FCA",cant:1,sueldo:6000,cp:3933},
    {puesto:"Administrativos",cant:3,sueldo:4800,cp:3933},
    {puesto:"Monitoreo CCTV",cant:3,sueldo:5000,cp:3933},
    {puesto:"Oficiales Turno A",cant:11,sueldo:4000,cp:3933},
    {puesto:"Oficiales Turno B",cant:10,sueldo:4000,cp:3933},
    {puesto:"Chofer / Logística",cant:1,sueldo:4800,cp:3933}
];
var NOM_MEN=0;
for(var _i=0;_i<PERSONAL.length;_i++) NOM_MEN+=PERSONAL[_i].cant*(PERSONAL[_i].sueldo+PERSONAL[_i].cp)*4.33;
var ADMIN=NOM_MEN*0.18, EQUIPO=(345000*0.65/11)+10063+4800;
var SUB=NOM_MEN+ADMIN+EQUIPO, IVA=SUB*0.16, TOTAL=SUB+IVA, PRIMER=TOTAL+(345000*0.35)+110000;
function fmt(n){return "$"+n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g,",");}

// ═══ AB1 PORTADA ══════════════════════════════════════════════════════════════
function portada(doc){
    var O=0;
    R(doc,0,0,A4W,A4H,C_BLANCO);
    R(doc,0,0,A4W,A4H*0.60,C_NEGRO);
    R(doc,0,0,7*MM,A4H,C_PLATA);
    R(doc,0,A4H*0.60,A4W,1.5*MM,C_PLATA);
    logo(doc,A4W/2-25*MM,A4H*0.07,50*MM);
    T(doc,"PROPUESTA DE PRESERVACIÓN OPERATIVA",10*MM,A4H*0.32,18,C_BLANCO,FB,A4W-14*MM);
    T(doc,"Y CONTINUIDAD DE SEGURIDAD INDUSTRIAL",10*MM,A4H*0.32+20,12,C_PLATA,FS,A4W-14*MM);
    R(doc,10*MM,A4H*0.32+36,50*MM,1.5*MM,C_AMARILLO);
    T(doc,"CLIENTE: STELLANTIS — FCA PLANTA SALTILLO",10*MM,A4H*0.32+44,10,C_BLANCO,FS,A4W-14*MM);
    R(doc,10*MM,A4H*0.63,A4W-20*MM,A4H*0.25,C_GRISC);
    R(doc,10*MM,A4H*0.63,2*MM,A4H*0.25,C_NEGRO);
    T(doc,"Folio:",14*MM,A4H*0.65,7,C_GRIS,FL,40);T(doc,FOLIO,14*MM,A4H*0.65+10,9,C_NEGRO,FB,200);
    T(doc,"Fecha:",14*MM,A4H*0.65+24,7,C_GRIS,FL,40);T(doc,FECHA,14*MM,A4H*0.65+34,9,C_NEGRO,FB,100);
    T(doc,"Vigencia:",14*MM,A4H*0.65+48,7,C_GRIS,FL,50);T(doc,"15 días naturales",14*MM,A4H*0.65+58,9,C_NEGRO,FS,150);
    T(doc,"MICSA SAFETY DIVISION",A4W-MRG-160,A4H*0.65+10,8,C_NEGRO,FB,160);
    T(doc,"contacto@micsa.mx  |  800-MICSA-01",A4W-MRG-160,A4H*0.65+24,7,C_GRIS,FL,160);
    R(doc,0,A4H*0.92,A4W,A4H*0.08,C_NEGRO);
    T(doc,"CONFIDENCIAL — Propietario MICSA Safety Division 2026",10*MM,A4H*0.94,7,C_PLATA,FL,A4W-14*MM,"C");
}

// ═══ AB2 ALCANCE ══════════════════════════════════════════════════════════════
function alcance(doc){
    var O=1; R(doc,O*A4W,0,A4W,A4H,C_BLANCO);
    mkHeader(doc,O,"ANÁLISIS DE RIESGO Y ALCANCE DEL SERVICIO");mkFooter(doc,O,"2");
    var y=19*MM;
    y=SEC(doc,O,y,"1. ANÁLISIS DE RIESGO OPERATIVO");
    var cW1=[TW*0.60,TW*0.40];
    y=tabla(doc,O,y,["RIESGO NO ATENDIDO","IMPACTO ESTIMADO"],[
        ["Pérdida patrimonial por turno sin cobertura","$85,000 – $320,000 MXN / evento"],
        ["Paro de línea por incidente de seguridad","$420,000 – $1,200,000 MXN / hora"],
        ["Multa STPS incumplimiento NOM-030","$80,000 – $240,000 MXN"],
        ["Responsabilidad civil patronal no cubierta","Litigio + fianza + daño imagen"],
        ["Robo de componentes sin cobertura nocturna","$180,000 – $600,000 MXN / turno"]
    ],cW1);
    y+=10*MM;
    y=SEC(doc,O,y,"2. ALCANCE DE LA SOLUCIÓN OPERATIVA");
    var cW2=[TW*0.22,TW*0.40,TW*0.38];
    y=tabla(doc,O,y,["ÁREA / ZONA","COBERTURA","SERVICIO"],[
        ["Planta FCA Saltillo","24/7 turnos A y B de 12h","Control acceso vehicular y peatonal"],
        ["Nave Estancias","Supervisión operativa continua","Registro digital ingresos/salidas"],
        ["Bodega Norte (NASA)","Rondines con cuatrimoto","Bitácora QR en puntos de control"],
        ["Centro de Monitoreo","3 operadores CCTV 24/7","Reporte incidentes en tiempo real"],
        ["Gestión KPI","Jefe de Seguridad + Supervisor","Dashboard mensual con indicadores"]
    ],cW2);
    y+=10*MM;
    y=SEC(doc,O,y,"3. MARCO NORMATIVO");
    var cW3=[TW*0.22,TW*0.52,TW*0.26];
    tabla(doc,O,y,["NORMA","DESCRIPCIÓN","ESTATUS"],[
        ["NOM-030-STPS-2009","Servicios preventivos de seguridad y salud","Cumplimiento total"],
        ["NOM-017-STPS-2008","Equipo de Protección Personal","Cumplimiento total"],
        ["REPSE","Registro Prestadoras Servicios Especializados","Registrado vigente"],
        ["IMSS / INFONAVIT","Obligaciones patronales de seguridad social","Al corriente"]
    ],cW3);
}

// ═══ AB3 NÓMINA ═══════════════════════════════════════════════════════════════
function nomina(doc){
    var O=2; R(doc,O*A4W,0,A4W,A4H,C_BLANCO);
    mkHeader(doc,O,"ESTRUCTURA DE INVERSIÓN EN CAPITAL HUMANO");mkFooter(doc,O,"3");
    var y=19*MM;
    y=SEC(doc,O,y,"3. INVERSIÓN EN CAPITAL HUMANO");
    T(doc,"Costos semanales por puesto — proyección mensual ×4.33 (IMSS incluido)",O*A4W+MRG,y,7,C_GRIS,FL,TW);y+=10*MM;
    var rows=[];
    for(var i=0;i<PERSONAL.length;i++){
        var p=PERSONAL[i],sub=p.cant*(p.sueldo+p.cp),men=sub*4.33;
        rows.push([p.puesto,String(p.cant),fmt(p.sueldo),fmt(p.cp),fmt(sub),fmt(men)]);
    }
    rows.push(["TOTAL NÓMINA MENSUAL","","","","",fmt(NOM_MEN)]);
    var cW=[TW*0.28,TW*0.06,TW*0.14,TW*0.14,TW*0.17,TW*0.21];
    y=tabla(doc,O,y,["PUESTO","CANT","SUELDO/SEM","CUOTAS P.","SUBTOTAL/SEM","PROY. MES"],rows,cW);
    y+=12*MM;
    y=SEC(doc,O,y,"PERFIL DEL PERSONAL ASIGNADO");
    var cW2=[TW*0.25,TW*0.75];
    y=tabla(doc,O,y,["CRITERIO","ESPECIFICACIÓN"],[
        ["Escolaridad","Secundaria concluida mínima, bachillerato preferente"],
        ["Experiencia","Mínimo 1 año en seguridad industrial o patrimonial"],
        ["Evaluación","Psicométrica ≥80%: honestidad y reacción bajo presión"],
        ["Documentación","Sin antecedentes penales · Cartilla militar · Domicilio verificable"],
        ["Turno","Rotativo 12h: A 07:00-19:00 / B 19:00-07:00"],
        ["Sustitución","Garantía reemplazo < 4 horas ante ausencia"]
    ],cW2);
    y+=12*MM;
    y=SEC(doc,O,y,"KPI'S DE DESEMPEÑO GARANTIZADOS");
    var cW3=[TW*0.30,TW*0.18,TW*0.52];
    tabla(doc,O,y,["KPI","META","CRITERIO"],[
        ["Puntualidad","≥ 97%","3 faltas en 30 días = baja automática"],
        ["Bitácoras completas","100%","Sin firma = turno incompleto + sanción"],
        ["Rondines cumplidos","≥ 95%","Registro QR o firma por punto de control"],
        ["Incidentes no reportados","0","Notificación inmediata + informe ≤ 2 horas"],
        ["Rotación mensual","≤ 5%","Evaluación semanal + expediente individual"]
    ],cW3);
}

// ═══ AB4 EQUIPAMIENTO ══════════════════════════════════════════════════════════
function equipamiento(doc){
    var O=3; R(doc,O*A4W,0,A4W,A4H,C_BLANCO);
    mkHeader(doc,O,"EQUIPAMIENTO Y LOGÍSTICA OPERATIVA");mkFooter(doc,O,"4");
    var y=19*MM;
    y=SEC(doc,O,y,"4. EQUIPAMIENTO Y LOGÍSTICA");
    T(doc,"Inversión en activos operativos — desglose primer mes vs. costo mensual recurrente",O*A4W+MRG,y,7,C_GRIS,FL,TW);y+=10*MM;
    var cW=[TW*0.38,TW*0.06,TW*0.18,TW*0.18,TW*0.20];
    y=tabla(doc,O,y,["CONCEPTO","CANT.","VALOR UNIT.","TOTAL","FREC."],[
        ["Polaris Sportsman 570 (cuatrimoto)","2",fmt(172500),fmt(345000),"Mes 1"],
        ["  — Enganche inicial 35%","1",fmt(120750),fmt(120750),"Mes 1"],
        ["  — Saldo 11 mensualidades","—","—",fmt(345000*0.65),"M 2-12"],
        ["Póliza seguros anual ambas motos","1",fmt(120750),fmt(120750),"Anual"],
        ["  — Cuota mensual póliza","12",fmt(10062.5),fmt(10062.5),"Mensual"],
        ["Gasolina estimada (ambas motos)","—","—",fmt(4800),"Mensual"],
        ["Uniformes completos + EPP (30 personas)","30",fmt(3666.67),fmt(110000),"Mes 1"],
        ["Radios comunicación (pares)","6",fmt(2500),fmt(15000),"Mes 1"],
        ["Linternas tácticas recargables","15",fmt(450),fmt(6750),"Mes 1"]
    ],cW);
    y+=12*MM;
    y=SEC(doc,O,y,"EPP INDIVIDUAL POR GUARDIA");
    var cW2=[TW*0.45,TW*0.20,TW*0.35];
    tabla(doc,O,y,["EQUIPO DE PROTECCIÓN PERSONAL","DOTACIÓN","NORMA"],[
        ["Uniforme completo (pantalón + camisa manga larga)","2 juegos","NOM identificación MICSA"],
        ["Casco de seguridad clase E","1 pieza","NOM-115-STPS-2009"],
        ["Calzado de seguridad punta de acero","1 par","NOM-113-STPS-2009"],
        ["Chaleco reflectante clase II","1 pieza","ANSI/ISEA 107"],
        ["Guantes de protección media caña","2 pares","NOM-138-STPS-2012"],
        ["Gafas de seguridad policarbonato","1 pieza","ANSI Z87.1"],
        ["Silbato de emergencia + lanyard","1 pieza","Señalización acústica"],
        ["Mascarilla N95 (caja mensual)","Mensual","Protección respiratoria básica"]
    ],cW2);
}

// ═══ AB5 FINANCIERO ════════════════════════════════════════════════════════════
function financiero(doc){
    var O=4; R(doc,O*A4W,0,A4W,A4H,C_BLANCO);
    mkHeader(doc,O,"RESUMEN ECONÓMICO — EQUILIBRIO OPERATIVO");mkFooter(doc,O,"5");
    var y=19*MM;
    y=SEC(doc,O,y,"5. RESUMEN ECONÓMICO");
    var cW=[TW*0.48,TW*0.16,TW*0.14,TW*0.22];
    y=tabla(doc,O,y,["CONCEPTO","TIPO","FREC.",  "MONTO"],[
        ["Nómina operativa (30 elementos)","Variable","Mensual",fmt(NOM_MEN)],
        ["Fee administrativo MICSA Safety (18%)","Fijo","Mensual",fmt(ADMIN)],
        ["Equipamiento y logística cuota mensual","Variable","Mensual",fmt(EQUIPO)],
        ["SUBTOTAL SIN IVA","","",fmt(SUB)],
        ["IVA 16%","Fiscal","Mensual",fmt(IVA)],
        ["TOTAL MENSUAL (Meses 2-12)","","",fmt(TOTAL)],
        ["Uniformes + EPP 30 personas (único)","","Mes 1",fmt(110000)],
        ["Enganche cuatrimotos 35% (único)","","Mes 1",fmt(345000*0.35)],
        ["PRIMER MES — INVERSIÓN INICIAL TOTAL","","",fmt(PRIMER)]
    ],cW);
    y+=12*MM;
    y=SEC(doc,O,y,"ANÁLISIS COSTO vs. RIESGO");
    var cW2=[TW*0.22,TW*0.48,TW*0.30];
    y=tabla(doc,O,y,["MONTO","CONCEPTO","NATURALEZA"],[
        [fmt(TOTAL),"Costo mensual total con IVA (meses 2-12)","Conocido, presupuestable"],
        [fmt(420000),"1 hora de paro de línea FCA","Riesgo sin cobertura"],
        [fmt(1200000),"Paro extendido 3h + daño reputacional","Riesgo sin cobertura"],
        [fmt(240000),"Multa STPS máxima NOM-030","Riesgo sin cobertura"]
    ],cW2);
    y+=12*MM;
    y=SEC(doc,O,y,"CONDICIONES DE PAGO");
    var cW3=[TW*0.24,TW*0.18,TW*0.35,TW*0.23];
    tabla(doc,O,y,["CONCEPTO","MODALIDAD","FECHA LÍMITE","MONTO"],[
        ["Primer pago (Mes 1)","Contado","A la firma del contrato",fmt(PRIMER)],
        ["Mensualidades 2-12","30 días","Del 1-5 de cada mes",fmt(TOTAL)],
        ["Moneda","MXN","Pesos mexicanos","—"],
        ["Forma de pago","SPEI / Cheque","—","—"]
    ],cW3);
}

// ═══ AB6 CONDICIONES ══════════════════════════════════════════════════════════
function condiciones(doc){
    var O=5; R(doc,O*A4W,0,A4W,A4H,C_BLANCO);
    mkHeader(doc,O,"CONDICIONES DE FORMALIZACIÓN");mkFooter(doc,O,"6");
    var y=19*MM;
    y=SEC(doc,O,y,"6. CONDICIONES DE FORMALIZACIÓN");
    var cW=[TW*0.35,TW*0.65];
    y=tabla(doc,O,y,["CONDICIÓN","DETALLE"],[
        ["Vigencia de propuesta","15 días naturales desde emisión"],
        ["Requisito de inicio","Firma de contrato + liquidación primer pago"],
        ["Inicio de operaciones","5 días hábiles tras formalización"],
        ["Ajuste de precios","Semestral conforme índices STPS / INPC"],
        ["Terminación anticipada","Aviso 30 días naturales por escrito"],
        ["Confidencialidad","NDA firmado — información protegida 5 años"]
    ],cW);
    y+=10*MM;
    y=SEC(doc,O,y,"DOCUMENTACIÓN REQUERIDA");
    var cW2=[TW*0.47,TW*0.33,TW*0.20];
    y=tabla(doc,O,y,["DOCUMENTO","RESPONSABLE","ESTATUS"],[
        ["Contrato Servicios Especializados (REPSE)","MICSA Safety + Cliente","Obligatorio"],
        ["Carta aceptación de propuesta firmada","Representante legal cliente","Obligatorio"],
        ["Alta sistema de proveedores cliente","Depto. Compras","Obligatorio"],
        ["Orden de compra (PO)","Finanzas / Compras","Obligatorio"],
        ["Datos fiscales (constancia SAT)","Ambas partes","Obligatorio"],
        ["Asignación contacto operativo planta","Gerencia Seguridad FCA","Obligatorio"]
    ],cW2);
    y+=12*MM;
    R(doc,O*A4W+MRG,y,TW,50,C_NEGRO);
    T(doc,"MICSA SAFETY DIVISION — TU SOCIO ESTRATÉGICO EN SEGURIDAD INDUSTRIAL",O*A4W+MRG,y+8,9,C_BLANCO,FB,TW,"C");
    T(doc,'"La demora en la formalización transfiere al cliente la responsabilidad operativa"',O*A4W+MRG,y+20,7.5,C_PLATA,FL,TW,"C");
    T(doc,"contacto@micsa.mx  ·  800-MICSA-01  ·  www.micsa.mx",O*A4W+MRG,y+33,7,C_AMARILLO,FS,TW,"C");
    y+=60;
    var fw=TW*0.42;
    T(doc,"________________________________",O*A4W+MRG,y+22,8,C_GRIS,FL,fw);
    T(doc,"________________________________",O*A4W+A4W-MRG-fw,y+22,8,C_GRIS,FL,fw);
    T(doc,"Director General — MICSA Safety Division",O*A4W+MRG,y+28,7,C_GRIS,FL,fw);
    T(doc,"Representante Legal — Stellantis / FCA",O*A4W+A4W-MRG-fw,y+28,7,C_GRIS,FL,fw);
}

// ═══ MAIN ══════════════════════════════════════════════════════════════════════
function main(){
    var doc=app.documents.add(DocumentColorSpace.CMYK, A4W*6, A4H);
    for(var i=1;i<6;i++) doc.artboards.add([i*A4W,0,(i+1)*A4W,-A4H]);
    var N=["01_PORTADA","02_ALCANCE","03_NOMINA","04_EQUIPAMIENTO","05_FINANCIERO","06_CONDICIONES"];
    for(var i=0;i<6;i++) doc.artboards[i].name=N[i];

    portada(doc); alcance(doc); nomina(doc);
    equipamiento(doc); financiero(doc); condiciones(doc);

    var ai=new File(BASE_PATH+"COTIZACION_FCA_BRANDED.ai");
    var ao=new IllustratorSaveOptions();
    ao.compatibility=Compatibility.ILLUSTRATOR24;
    ao.saveMultipleArtboards=true;
    doc.saveAs(ai,ao);

    try{
        var pdf=new File(OUT_PATH+"COTIZACION_FCA_BRANDED.pdf");
        var po=new PDFSaveOptions();
        po.compatibility=PDFCompatibility.ACROBAT8;
        po.generateThumbnails=true;
        po.preserveEditability=false;
        po.saveMultipleArtboards=true;
        po.artboardRange="1-6";
        doc.saveAs(pdf,po);
    }catch(e){}

    alert("COTIZACION_FCA_BRANDED.ai generada — 6 páginas.\nPortada · Alcance · Nómina · Equipamiento · Financiero · Condiciones\nPDF exportado en PDFs/");
}
main();
