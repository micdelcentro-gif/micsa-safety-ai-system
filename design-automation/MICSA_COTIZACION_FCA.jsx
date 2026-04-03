/**
 * MICSA SAFETY DIVISION — Cotizacion Formal FCA / Stellantis Saltillo
 * ExtendScript para Adobe Illustrator
 * 7 Artboards — Letter 8.5x11in — CMYK Negro + Plateado
 * USO: Archivo > Scripts > Otro script...
 */

var BASE = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\micsa-safety-ai-system\\";
var LOGO = BASE + "assets\\logo.png";
var OUT  = BASE + "output\\";

var MM  = 2.834645669;
var LW  = 8.5 * 72;
var LH  = 11  * 72;
var MRG = 20  * MM;

function C(c,m,y,k){ var x=new CMYKColor(); x.cyan=c; x.magenta=m; x.yellow=y; x.black=k; return x; }
var BLK  = C(60,40,40,100);
var SLV  = C(0,0,0,25);
var SLV2 = C(0,0,0,12);
var DGRY = C(0,0,0,70);
var MGRY = C(0,0,0,30);
var LGRY = C(0,0,0,7);
var WHT  = C(0,0,0,0);
var RED  = C(0,90,85,0);

var FB = "Montserrat-Bold";
var FS = "Montserrat-SemiBold";
var FR = "OpenSans-Regular";
var FL = "OpenSans-Light";

// ── Primitivos ────────────────────────────────────────────────────────
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
    if(al==="C") t.textRange.paragraphAttributes.justification=Justification.CENTER;
    if(al==="R") t.textRange.paragraphAttributes.justification=Justification.RIGHT;
    return t;
}
function L(doc,x1,y1,x2,y2,col,w){
    var l=doc.pathItems.add();
    l.setEntirePath([[x1,-y1],[x2,-y2]]);
    l.stroked=true;l.strokeColor=col;l.strokeWidth=w||0.5;l.filled=false;return l;
}
function logo(doc,x,y,w){
    try{
        var f=new File(LOGO);if(!f.exists)return;
        var p=doc.placedItems.add();p.file=f;
        var sc=w/p.width;p.width=w;p.height=p.height*sc;
        p.left=x;p.top=-y;
    }catch(e){}
}

// ── Componentes reutilizables ─────────────────────────────────────────
function HDR(doc,O,titulo,pg){
    // Reset background
    R(doc,O*LW,0,LW,LH,WHT);
    
    // Tactical Sidebar
    R(doc,O*LW,0,4*MM,LH,BLK);
    R(doc,O*LW+4*MM,0,0.8*MM,LH,RED);

    // Header Banner
    R(doc,O*LW+4.8*MM,0,LW-4.8*MM,15*MM,BLK);
    L(doc,O*LW+4.8*MM,15*MM,O*LW+LW,15*MM,RED,0.5);
    
    logo(doc,O*LW+8*MM,2*MM,14*MM);
    T(doc,"MICSA SAFETY DIVISION",O*LW+25*MM,4.5*MM,8,WHT,FB,LW*0.5);
    T(doc,titulo.toUpperCase(),O*LW+25*MM,10.5*MM,6.5,LGRY,FL,LW*0.55);
    
    // Page Signal
    R(doc,(O+1)*LW-15*MM,2*MM,10*MM,11*MM,RED);
    T(doc,String(pg),(O+1)*LW-10*MM,5.5*MM,10,WHT,FB,8*MM,"C");
}
function FTR(doc,O){
    R(doc,O*LW+4.8*MM,LH-10*MM,LW-4.8*MM,10*MM,BLK);
    R(doc,O*LW+4.8*MM,LH-10*MM,LW-4.8*MM,0.5*MM,RED);
    T(doc,"CONFIDENCIAL  |  Propuesta Stellantis / FCA  |  micsasafety.com.mx",
      O*LW+10*MM,LH-6*MM,5.5,MGRY,FL,LW-MRG*2);
}
function SEC(doc,O,y,titulo){
    R(doc,O*LW+MRG,y,LW-MRG*2,8*MM,LGRY,RED,0.3);
    R(doc,O*LW+MRG,y,2.5*MM,8*MM,BLK);
    T(doc,titulo.toUpperCase(),O*LW+MRG+6*MM,y+1.5*MM,9,BLK,FB,LW-MRG*2-8*MM);
    y+=10*MM;
    L(doc,O*LW+MRG,y,O*LW+LW-MRG,y,MGRY,0.5);
    return y+6*MM;
}
function TBL(doc,O,y,hdrs,rows,cols){
    var x=O*LW+MRG,rh=9.5*MM,tw=LW-MRG*2,cx,cw,h,r,c;
    cx=x;
    for(h=0;h<hdrs.length;h++){
        cw=cols[h]*tw;
        R(doc,cx,y,cw,rh,BLK,SLV,0.3);
        L(doc,cx,y,cx,y+rh,SLV,0.3);
        T(doc,hdrs[h],cx+2*MM,y+2.5*MM,6.5,WHT,FB,cw-4*MM);
        cx+=cw;
    }
    y+=rh;
    for(r=0;r<rows.length;r++){
        cx=x;
        var bg=r%2===0?LGRY:WHT;
        for(c=0;c<rows[r].length;c++){
            cw=cols[c]*tw;
            R(doc,cx,y,cw,rh,bg,MGRY,0.3);
            T(doc,rows[r][c],cx+2*MM,y+2.5*MM,6.5,DGRY,FR,cw-4*MM);
            cx+=cw;
        }
        y+=rh;
    }
    return y;
}
function FILASUM(doc,O,y,label,valor,highlight){
    var tw=LW-MRG*2;
    var bg=highlight?BLK:LGRY;
    var tc=highlight?WHT:DGRY;
    var vc=highlight?RED:BLK;
    R(doc,O*LW+MRG,y,tw,10*MM,bg,MGRY,0.3);
    T(doc,label,O*LW+MRG+3*MM,y+2.5*MM,7,tc,highlight?FS:FR,tw*0.65);
    T(doc,valor,O*LW+MRG+tw*0.67,y+2.5*MM,7,vc,highlight?FB:FS,tw*0.30,"R");
    return y+10*MM;
}

// ═══════════════════════════════════════════════════════════════════
// AB1 — PORTADA
// ═══════════════════════════════════════════════════════════════════
function portada(doc){
    var O=0;
    // Fondo negro total
    R(doc,O*LW,0,LW,LH,BLK);
    // Sidebar Tactico
    R(doc,O*LW,0,5*MM,LH,RED);
    R(doc,O*LW+5*MM,0,0.8*MM,LH,WHT);
    
    // Zona blanca inferior
    R(doc,O*LW+5.8*MM,LH*0.52,(LW-5.8*MM),LH*0.38,WHT);
    R(doc,O*LW+5.8*MM,LH*0.52,LW-5.8*MM,1.5*MM,RED);
    // Logo centrado
    logo(doc,O*LW+LW/2-32*MM,LH*0.07,64*MM);
    // Etiqueta tipo documento
    R(doc,O*LW+10*MM,LH*0.27,52*MM,7*MM,RED);
    T(doc,"PROPUESTA COMERCIAL FORMAL",O*LW+11*MM,LH*0.27+1.5*MM,6.5,WHT,FB,50*MM,"C");
    // Titulo
    T(doc,"COTIZACION DE",O*LW+10*MM,LH*0.29+8*MM,24,WHT,FB,LW-14*MM);
    T(doc,"SERVICIOS DE SEGURIDAD",O*LW+10*MM,LH*0.29+36,24,WHT,FB,LW-14*MM);
    // Linea roja + subtitulo
    R(doc,O*LW+10*MM,LH*0.29+64,52*MM,1.5*MM,RED);
    T(doc,"Vigilancia Patrimonial  |  Control de Accesos  |  Seguridad Industrial",
      O*LW+10*MM,LH*0.29+71,9,LGRY,FL,LW-16*MM);
    // Datos cliente en zona blanca
    T(doc,"CLIENTE CORPORATIVO",O*LW+10*MM,LH*0.54,7,RED,FB,35*MM);
    T(doc,"FCA / Stellantis Saltillo â€” Planta Ensamble Norte",
      O*LW+10*MM,LH*0.54+10,9.5,BLK,FB,LW*0.55);
    T(doc,"Blvd. Industria Automotriz S/N  |  Saltillo, Coahuila  |  C.P. 25017",
      O*LW+10*MM,LH*0.54+22,7.5,DGRY,FR,LW*0.55);
    L(doc,O*LW+LW*0.60,LH*0.53,O*LW+LW*0.60,LH*0.90,RED,0.5);
    T(doc,"No. Cotizacion",O*LW+LW*0.62,LH*0.54,7,DGRY,FL,LW*0.33);
    T(doc,"COT-FCA-2026-04",O*LW+LW*0.62,LH*0.54+10,9,BLK,FB,LW*0.33);
    T(doc,"Fecha de Emision",O*LW+LW*0.62,LH*0.54+24,7,DGRY,FL,LW*0.33);
    T(doc,"Abril 2026",O*LW+LW*0.62,LH*0.54+34,8.5,BLK,FS,LW*0.33);
    T(doc,"Vigencia",O*LW+LW*0.62,LH*0.54+48,7,DGRY,FL,LW*0.33);
    T(doc,"30 dias naturales",O*LW+LW*0.62,LH*0.54+58,8.5,BLK,FS,LW*0.33);
    T(doc,"Director de Seguridad",O*LW+10*MM,LH*0.78,7,DGRY,FL,LW*0.50);
    T(doc,"Gerardo Guzman Alvarado",O*LW+10*MM,LH*0.78+10,9,BLK,FB,LW*0.50);
    T(doc,"Fuerza Civil  |  Proteccion Civil  |  Bomberos",
      O*LW+10*MM,LH*0.78+21,7,DGRY,FR,LW*0.50);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB2 — ANALISIS DE RIESGO
// ═══════════════════════════════════════════════════════════════════
function analisisRiesgo(doc){
    var O=1;
    HDR(doc,O,"Analisis de Riesgo del Sitio",2);
    var y=22*MM;
    // Ficha de sitio
    R(doc,O*LW+MRG,y,LW-MRG*2,22*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,3*MM,22*MM,RED);
    T(doc,"DESCRIPCION DEL SITIO",O*LW+MRG+5*MM,y+2.5*MM,7,DGRY,FB,80*MM);
    T(doc,"Planta Ensamble Norte — FCA / Stellantis Saltillo",
      O*LW+MRG+5*MM,y+10*MM,8.5,BLK,FB,LW*0.55);
    T(doc,"Superficie aprox. 320,000 m2   |   Turnos: 3 (24/7)   |   Trabajadores: ~4,500   |   Accesos: 6 puertas",
      O*LW+MRG+5*MM,y+18.5*MM,7,DGRY,FR,LW-MRG*2-8*MM);
    y+=26*MM;
    y=SEC(doc,O,y,"1. Matriz de Riesgo por Zona");
    y=TBL(doc,O,y,
        ["ZONA","NIVEL DE RIESGO","AMENAZA PRINCIPAL","CONTROL RECOMENDADO"],
        [
            ["Acceso vehicular principal","ALTO","Ingreso no autorizado","Caseta blindada + retractiles"],
            ["Perimetro industrial","ALTO","Robo de material / escalamiento","Rondines + CCTV perimetral"],
            ["Planta de ensamble lineas A-D","MEDIO-ALTO","Sabotaje / accidente industrial","Supervision continua por turno"],
            ["Almacen de materias primas","ALTO","Robo hormi\u0301n + merma","Control de inventario + C\u00E1mara 4K"],
            ["Estacionamiento empleados","MEDIO","Robo a persona","Luminaria + ronda motorizada"],
            ["Oficinas corporativas","MEDIO","Acceso no autorizado","Credencializacion + torniquetes"],
            ["Caseta de transformadores","ALTO","Vandalismo / sabotaje","Acceso restringido nivel 3"]
        ],
        [0.24,0.16,0.30,0.30]);
    y+=5*MM;
    y=SEC(doc,O,y,"2. Indicadores de Seguridad Actuales (auditoria inicial)");
    y=TBL(doc,O,y,
        ["INDICADOR","ESTADO ACTUAL","META MICSA"],
        [
            ["Cobertura CCTV","42% de zonas criticas","95% en 90 dias"],
            ["Respuesta a incidentes","8.5 minutos promedio","< 3 minutos"],
            ["Capacitacion del personal","Sin certificacion vigente","100% certificado NOM-030"],
            ["Control de accesos digitales","Parcial (3 de 6 accesos)","100% biometrico"],
            ["Reportes de incidentes","Manuales / irregulares","Sistema digital tiempo real"]
        ],
        [0.34,0.33,0.33]);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB3 — ALCANCE DEL SERVICIO
// ═══════════════════════════════════════════════════════════════════
function alcance(doc){
    var O=2;
    HDR(doc,O,"Alcance del Servicio",3);
    var y=22*MM;
    y=SEC(doc,O,y,"1. Descripcion General del Servicio");
    // Bloque descripcion
    R(doc,O*LW+MRG,y,LW-MRG*2,18*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,3*MM,18*MM,RED);
    T(doc,"MICSA Safety Division provee el servicio integral de seguridad patrimonial e industrial para la planta FCA / Stellantis Saltillo, incluyendo vigilancia presencial 24/7, supervisión operativa, control de accesos vehicular y peatonal, rondines internos y perímetro, atención a emergencias, y reportes ejecutivos mensuales conforme a protocolos NOM-030 y NOM-017.",
      O*LW+MRG+5*MM,y+3*MM,7.5,DGRY,FR,LW-MRG*2-8*MM);
    y+=22*MM;
    y=SEC(doc,O,y,"2. Cobertura por Puesto y Turno");
    y=TBL(doc,O,y,
        ["PUESTO","TURNO","HORARIO","CANTIDAD","ZONA ASIGNADA"],
        [
            ["Supervisor General","Lunes a Viernes","06:00–14:00","1","Toda la planta"],
            ["Supervisor Nocturno","Lunes a Domingo","22:00–06:00","1","Toda la planta"],
            ["Guardia Acceso Ppal.","24/7 x 365","Rotativo","4","Caseta principal"],
            ["Guardia Perimetro","24/7 x 365","Rotativo","6","Perimetro + rondas"],
            ["Guardia Almacen","Lunes a Sabado","07:00–21:00","3","Almacen general"],
            ["Guardia Of. Corporativas","Lunes a Viernes","08:00–18:00","2","Edificio admin."],
            ["Vigilante Estacionamiento","24/7 x 365","Rotativo","4","Parking A, B y C"]
        ],
        [0.20,0.18,0.16,0.12,0.34]);
    y+=5*MM;
    y=SEC(doc,O,y,"3. Servicios Complementarios Incluidos");
    y=TBL(doc,O,y,
        ["SERVICIO ADICIONAL","INCLUYE","FRECUENCIA"],
        [
            ["Reporte ejecutivo de incidentes","PDF + dashboard digital","Mensual"],
            ["Supervision tecnica MICSA","Visita de auditor senior","Quincenal"],
            ["Capacitacion continua NOM","Cursos certificados","Semestral"],
            ["Comunicacion radial 24/7","Radios Motorola + base","Permanente"],
            ["Protocolo de emergencias","Plan de evacuacion actualizado","Anual + simulacros"],
            ["Gestion de credenciales","Sistema biometrico + base datos","Permanente"]
        ],
        [0.34,0.40,0.26]);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB4 — ESTRUCTURA DE NOMINA
// ═══════════════════════════════════════════════════════════════════
function nomina(doc){
    var O=3;
    HDR(doc,O,"Estructura de Nomina y Costos de Personal",4);
    var y=22*MM;
    y=SEC(doc,O,y,"1. Desglose de Costos por Plaza");
    y=TBL(doc,O,y,
        ["PUESTO","CANT.","SUELDO NETO/MES","IMSS + SAR","TOTAL MENSUAL"],
        [
            ["Supervisor General","1","$18,500","$7,215","$25,715"],
            ["Supervisor Nocturno","1","$16,800","$6,552","$23,352"],
            ["Guardia Acceso Principal","4","$9,500 c/u","$3,705 c/u","$52,820"],
            ["Guardia Perimetro","6","$9,200 c/u","$3,588 c/u","$76,728"],
            ["Guardia Almacen","3","$9,200 c/u","$3,588 c/u","$38,364"],
            ["Guardia Of. Corporativas","2","$9,500 c/u","$3,705 c/u","$26,410"],
            ["Vigilante Estacionamiento","4","$9,000 c/u","$3,510 c/u","$50,040"]
        ],
        [0.30,0.08,0.20,0.18,0.24]);
    y+=5*MM;
    // Subtotal nomina
    R(doc,O*LW+MRG,y,LW-MRG*2,1*MM,SLV);
    y+=3*MM;
    y=FILASUM(doc,O,y,"Subtotal Nomina (21 elementos)","$293,429","");
    y=FILASUM(doc,O,y,"Factor prestaciones de ley (vacaciones, aguinaldo, prima vacacional)","+ $44,014","");
    y+=4*MM;
    y=SEC(doc,O,y,"2. Costo de Supervision y Administracion");
    y=TBL(doc,O,y,
        ["CONCEPTO","DETALLE","MONTO MENSUAL"],
        [
            ["Coordinador operativo MICSA","Seguimiento y auditoria","$12,000"],
            ["Administrativo / RRHH","Nomina, IMSS, contratos","$8,500"],
            ["Uniformes y equipamiento","Amortizacion mensual","$6,200"],
            ["Comunicacion y radios","Renta + mantenimiento Motorola","$4,800"],
            ["Seguros y fianzas","Responsabilidad civil operativa","$5,500"],
            ["Transporte supervisores","Viaticos + uso de vehiculo","$3,200"]
        ],
        [0.32,0.40,0.28]);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB5 — EQUIPAMIENTO Y TECNOLOGIA
// ═══════════════════════════════════════════════════════════════════
function equipamiento(doc){
    var O=4;
    HDR(doc,O,"Equipamiento y Tecnologia de Seguridad",5);
    var y=22*MM;
    y=SEC(doc,O,y,"1. Equipo Personal por Elemento");
    y=TBL(doc,O,y,
        ["ARTICULO","ESPECIFICACION","INCLUYE","CANT."],
        [
            ["Uniforme operativo","Pantalon tactico + camisa identificada + chaleco","Si","21 sets"],
            ["Calzado de seguridad","Punta de acero, suela antiderrapante","Si","21 pares"],
            ["Radio Motorola DP4400","Digital UHF, encriptado, bateria 8h","Si","21 radios"],
            ["Linterna tactica","LED 1000 lumens, resistente a golpes","Si","21 pzas"],
            ["Tolete de seguridad","Polimero reforzado, funda al cinto","Si","21 pzas"],
            ["Credencial con QR","Foto + datos IMSS, renovacion anual","Si","21 pzas"],
            ["Kit primeros auxilios","Botiquin estandar NOM-010-STPS","Si","6 kits"]
        ],
        [0.28,0.38,0.14,0.20]);
    y+=5*MM;
    y=SEC(doc,O,y,"2. Infraestructura y Tecnologia (incluida en cotizacion)");
    y=TBL(doc,O,y,
        ["SISTEMA","DESCRIPCION","CANTIDAD","ESTADO"],
        [
            ["Sistema biometrico accesos","Huella + tarjeta RFID, base cloud","6 lectores","Instalacion incluida"],
            ["CCTV Dahua 4K","DVR 32 canales, 30 dias grabacion","12 camaras","Instalacion incluida"],
            ["Base de comunicaciones","Repetidor Motorola MOTOTRBO","1 base","Instalacion incluida"],
            ["Bitacora digital","App MICSA, reporte tiempo real","Licencia total","Acceso cliente incluido"],
            ["Barrera retractil motorizada","Acero reforzado, control remoto","2 barreras","Instalacion incluida"],
            ["Tablero de control accesos","Monitor 24\", software integrado","1 tablero","Configuracion incluida"]
        ],
        [0.24,0.36,0.18,0.22]);
    y+=5*MM;
    R(doc,O*LW+MRG,y,LW-MRG*2,14*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,3*MM,14*MM,SLV);
    T(doc,"NOTA: Todo el equipo listado en esta seccion queda bajo resguardo de MICSA Safety Division y es parte del contrato de servicio. El cliente no realiza compra ni mantenimiento directo de equipo; MICSA gestiona el ciclo de vida completo.",
      O*LW+MRG+5*MM,y+3*MM,7,DGRY,FR,LW-MRG*2-8*MM);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB6 — RESUMEN FINANCIERO
// ═══════════════════════════════════════════════════════════════════
function resumenFinanciero(doc){
    var O=5;
    HDR(doc,O,"Resumen Financiero",6);
    var y=22*MM;
    y=SEC(doc,O,y,"Desglose de Inversion Mensual");
    // Tabla resumen
    y=TBL(doc,O,y,
        ["CONCEPTO","MONTO MENSUAL","NOTAS"],
        [
            ["Nomina base 21 elementos","$293,429","Sueldo neto + IMSS + SAR"],
            ["Prestaciones de ley","$44,014","Vacaciones, aguinaldo, prima"],
            ["Coordinacion y administracion","$20,500","Coord. operativo + RRHH MICSA"],
            ["Uniformes y equipamiento","$6,200","Amortizacion 24 meses"],
            ["Comunicacion radial","$4,800","Renta + mantenimiento Motorola"],
            ["Seguros y fianzas","$5,500","Responsabilidad civil"],
            ["Transporte supervisores","$3,200","Viaticos + vehiculo"],
            ["Infraestructura tecnologica","$12,000","Mantenimiento CCTV + biometrico"],
            ["Fee de administracion MICSA","$18,500","Gestion integral del contrato"]
        ],
        [0.44,0.28,0.28]);
    // Separador
    y+=3*MM;
    R(doc,O*LW+MRG,y,LW-MRG*2,1.2*MM,SLV);
    y+=4*MM;
    // Totales
    y=FILASUM(doc,O,y,"SUBTOTAL (sin IVA)","$408,143","");
    y=FILASUM(doc,O,y,"IVA (16%)","$65,303","");
    y+=2*MM;
    R(doc,O*LW+MRG,y,LW-MRG*2,1.2*MM,SLV);
    y+=2*MM;
    y=FILASUM(doc,O,y,"TOTAL MENSUAL (con IVA)","$473,446 MXN","highlight");
    y+=6*MM;
    // Primer mes
    R(doc,O*LW+MRG,y,LW-MRG*2,22*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,3*MM,22*MM,RED);
    T(doc,"PRIMER MES — PAGO ANTICIPADO",O*LW+MRG+5*MM,y+2.5*MM,7,BLK,FB,LW*0.5);
    T(doc,"El primer mes incluye ademas: instalacion de biometrico y CCTV, uniformes y equipamiento inicial, alta IMSS de personal, capacitacion NOM-030 de ingreso y configuracion de bitacora digital.",
      O*LW+MRG+5*MM,y+10.5*MM,7,DGRY,FR,LW-MRG*2-8*MM);
    y+=25*MM;
    y=FILASUM(doc,O,y,"Cargo por instalacion y puesta en marcha (unico)","$38,500 MXN","");
    y=FILASUM(doc,O,y,"TOTAL PRIMER MES (con instalacion e IVA)","$517,946 MXN","highlight");
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB7 — CONDICIONES Y FIRMA
// ═══════════════════════════════════════════════════════════════════
function condicionesFirma(doc){
    var O=6;
    HDR(doc,O,"Condiciones Comerciales y Firma",7);
    var y=22*MM;
    y=SEC(doc,O,y,"Condiciones Generales del Contrato");
    y=TBL(doc,O,y,
        ["CONDICION","DETALLE"],
        [
            ["Vigencia del contrato","12 meses con renovacion automatica anual"],
            ["Facturacion","Mensual al inicio del periodo, 15 dias de credito"],
            ["Forma de pago","Transferencia electronica / cheque certificado"],
            ["Ajuste por inflacion","INPC anual, notificacion con 30 dias de anticipacion"],
            ["Periodo de prueba","60 dias con opcion de ajuste de plantel sin penalizacion"],
            ["Penalizacion por cancelacion anticipada","2 mensualidades del contrato vigente"],
            ["Responsabilidad civil","Poliza $5,000,000 MXN incluida en tarifa"],
            ["Tiempo de reemplazo por baja","Maximo 72 horas para cubrir vacantes"],
            ["Reportes ejecutivos","Entrega mensual primeros 5 dias habiles"],
            ["Protocolo de emergencias","Activacion < 15 min, notificacion inmediata al cliente"]
        ],
        [0.38,0.62]);
    y+=5*MM;
    y=SEC(doc,O,y,"Aceptacion y Autorizacion");
    // Bloque de firmas
    var tw=LW-MRG*2;
    var col1=O*LW+MRG;
    var col2=O*LW+MRG+tw*0.52;
    var cw1=tw*0.45;
    var cw2=tw*0.45;
    // Firma cliente
    R(doc,col1,y,cw1,38*MM,WHT,MGRY,0.4);
    R(doc,col1,y,cw1,6*MM,BLK);
    T(doc,"CLIENTE — FCA / STELLANTIS",col1+3*MM,y+1.5*MM,7,WHT,FB,cw1-6*MM);
    L(doc,col1+5*MM,y+27*MM,col1+cw1-5*MM,y+27*MM,MGRY,0.5);
    T(doc,"Firma y sello",col1+5*MM,y+29*MM,6.5,MGRY,FL,cw1-10*MM,"C");
    T(doc,"Nombre:",col1+4*MM,y+33*MM,6.5,DGRY,FL,cw1-8*MM);
    T(doc,"Cargo:",col1+4*MM,y+37.5*MM,6.5,DGRY,FL,cw1-8*MM);
    // Firma MICSA
    R(doc,col2,y,cw2,38*MM,WHT,MGRY,0.4);
    R(doc,col2,y,cw2,6*MM,BLK);
    T(doc,"MICSA SAFETY DIVISION",col2+3*MM,y+1.5*MM,7,WHT,FB,cw2-6*MM);
    L(doc,col2+5*MM,y+27*MM,col2+cw2-5*MM,y+27*MM,MGRY,0.5);
    T(doc,"Firma y sello",col2+5*MM,y+29*MM,6.5,MGRY,FL,cw2-10*MM,"C");
    T(doc,"Gerardo Guzman Alvarado",col2+4*MM,y+33*MM,7.5,BLK,FB,cw2-8*MM);
    T(doc,"Director de Seguridad",col2+4*MM,y+37.5*MM,6.5,DGRY,FR,cw2-8*MM);
    y+=44*MM;
    // Leyenda final
    R(doc,O*LW+MRG,y,tw,14*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,3*MM,14*MM,RED);
    T(doc,"Al firmar este documento, ambas partes aceptan las condiciones descritas en la presente cotizacion y se obligan a formalizarlas mediante contrato de servicios. La vigencia de esta propuesta es de 30 dias naturales a partir de la fecha de emision.",
      O*LW+MRG+5*MM,y+3*MM,7,DGRY,FR,LW-MRG*2-8*MM);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// MAIN — Crear documento y exportar
// ═══════════════════════════════════════════════════════════════════
(function(){
    var docPreset = new DocumentPreset();
    docPreset.colorMode = DocumentColorSpace.CMYK;
    docPreset.units = RulerUnits.Points;
    docPreset.width = LW * 7;
    docPreset.height = LH;
    var doc = app.documents.addDocument(DocumentColorSpace.CMYK, docPreset);
    doc.artboards[0].artboardRect = [0, 0, LW, -LH];
    doc.artboards[0].name = "01_PORTADA";
    var abNames = ["02_ANALISIS_RIESGO","03_ALCANCE","04_NOMINA","05_EQUIPAMIENTO","06_RESUMEN_FINANCIERO","07_CONDICIONES_FIRMA"];
    for(var i=0;i<abNames.length;i++){
        doc.artboards.add([LW*(i+1),0,LW*(i+2),-LH]);
        doc.artboards[i+1].name = abNames[i];
    }
    portada(doc);
    analisisRiesgo(doc);
    alcance(doc);
    nomina(doc);
    equipamiento(doc);
    resumenFinanciero(doc);
    condicionesFirma(doc);
    // Guardar .ai
    var aiFile = new File(BASE + "MICSA_COTIZACION_FCA.ai");
    var saveOpts = new IllustratorSaveOptions();
    saveOpts.compatibility = Compatibility.ILLUSTRATOR17;
    saveOpts.compressed = false;
    doc.saveAs(aiFile, saveOpts);
    // Exportar PDF
    try{
        var pdfFile = new File(OUT + "MICSA_COTIZACION_FCA.pdf");
        var pdfOpts = new PDFSaveOptions();
        pdfOpts.compatibility = PDFCompatibility.ACROBAT7;
        pdfOpts.generateThumbnails = true;
        pdfOpts.preserveEditability = false;
        doc.saveAs(pdfFile, pdfOpts);
    }catch(e){}
    alert("MICSA_COTIZACION_FCA generado.\nGuardado en: " + BASE);
})();
