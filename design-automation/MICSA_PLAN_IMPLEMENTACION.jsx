/**
 * MICSA SAFETY DIVISION — Plan de Implementacion y Ruta Critica
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
    T(doc,"CONFIDENCIAL  |  Plan de Implementacion MICSA  |  micsasafety.com.mx",
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

// Barra de Gantt simplificada
function GANTT_HDR(doc,O,y){
    var tw=LW-MRG*2, lw=tw*0.28, bw=(tw*0.72)/8, rh=7*MM;
    R(doc,O*LW+MRG,y,lw,rh,BLK);
    T(doc,"FASE / ACTIVIDAD",O*LW+MRG+2*MM,y+2*MM,6,WHT,FB,lw-4*MM);
    var meses=["SEM 1","SEM 2","SEM 3","SEM 4","SEM 5","SEM 6","SEM 7","SEM 8"];
    for(var i=0;i<8;i++){
        R(doc,O*LW+MRG+lw+i*bw,y,bw,rh,BLK,RED,0.2);
        T(doc,meses[i],O*LW+MRG+lw+i*bw+1,y+2*MM,5.5,RED,FS,bw-2,"C");
    }
    return y+rh;
}
function GANTT_ROW(doc,O,y,fase,sem1,sem2,sem3,sem4,sem5,sem6,sem7,sem8){
    var tw=LW-MRG*2;
    var lw=tw*0.28, bw=(tw*0.72)/8, rh=8.5*MM;
    var semanas=[sem1,sem2,sem3,sem4,sem5,sem6,sem7,sem8];
    R(doc,O*LW+MRG,y,lw,rh,LGRY,MGRY,0.2);
    T(doc,fase,O*LW+MRG+2*MM,y+2*MM,6,DGRY,FR,lw-4*MM);
    for(var i=0;i<8;i++){
        var bx=O*LW+MRG+lw+i*bw;
        var fill=semanas[i]?RED:WHT;
        var stroke=semanas[i]?RED:MGRY;
        R(doc,bx,y,bw,rh,fill,stroke,0.2);
        if(semanas[i]) R(doc,bx+1,y+1.5*MM,bw-2,rh-3*MM,LGRY);
    }
    return y+rh;
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
    R(doc,O*LW+10*MM,LH*0.27,58*MM,7*MM,RED);
    T(doc,"GESTION OPERATIVA â€” DOCUMENTO INTERNO",O*LW+10*MM,LH*0.27+1.5*MM,6,WHT,FB,58*MM,"C");
    // Titulo
    T(doc,"PLAN DE",O*LW+10*MM,LH*0.29+8*MM,30,WHT,FB,LW-14*MM);
    T(doc,"IMPLEMENTACION",O*LW+10*MM,LH*0.29+42,22,WHT,FB,LW-14*MM);
    // Linea roja + subtitulo
    R(doc,O*LW+10*MM,LH*0.29+66,60*MM,1.5*MM,RED);
    T(doc,"Ruta Critica  |  Activacion de Contrato  |  Seguimiento Operativo",
      O*LW+10*MM,LH*0.29+74,9,LGRY,FL,LW-14*MM);
    // Datos y Alcance
    T(doc,"ALCANCE CORPORATIVO",O*LW+10*MM,LH*0.55,7,RED,FB,35*MM);
    T(doc,"Modelo de Arranque para Contratos de Seguridad Industrial de Alto Volumen",
      O*LW+10*MM,LH*0.55+10,9.5,BLK,FB,LW*0.55);
    T(doc,"v2.0  |  Monclova, Coahuila  |  2026",O*LW+10*MM,LH*0.55+24,7,DGRY,FL,LW*0.4);
    L(doc,O*LW+LW*0.60,LH*0.54,O*LW+LW*0.60,LH*0.90,RED,0.5);
    T(doc,"Director de Seguridad",O*LW+LW*0.62,LH*0.56,7,DGRY,FL,LW*0.33);
    T(doc,"Gerardo Guzman Alvarado",O*LW+LW*0.62,LH*0.56+10,8.5,BLK,FB,LW*0.33);
    T(doc,"Fuerza Civil  |  Proteccion Civil  |  Bomberos",O*LW+LW*0.62,LH*0.56+21,7,DGRY,FR,LW*0.33);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB2 — RESUMEN EJECUTIVO DEL PLAN
// ═══════════════════════════════════════════════════════════════════
function resumenEjecutivo(doc){
    var O=1;
    HDR(doc,O,"Resumen Ejecutivo del Plan",2);
    var y=22*MM;
    // KPIs de arranque
    var kpis=[["72 HRS","Alta IMSS de personal"],["5 DIAS","Instalacion de equipo tecnologico"],["2 SEM","Primer turno en operacion plena"],["30 DIAS","Auditoria de arranque con cliente"],["60 DIAS","Evaluacion y ajuste de plantilla"],["90 DIAS","Entrega primer reporte ejecutivo"]];
    var tw=LW-MRG*2;
    var cw=(tw-5*MM)/3;
    for(var i=0;i<kpis.length;i++){
        var col=i%3;
        var row=Math.floor(i/3);
        var kx=O*LW+MRG+col*(cw+2.5*MM);
        var ky=y+row*26*MM;
        R(doc,kx,ky,cw,22*MM,LGRY,MGRY,0.4);
        R(doc,kx,ky,cw,8*MM,BLK);
        R(doc,kx,ky,2*MM,22*MM,SLV);
        T(doc,kpis[i][0],kx+4*MM,ky+1.5*MM,10,SLV,FB,cw-6*MM);
        T(doc,kpis[i][1],kx+4*MM,ky+10.5*MM,6.5,DGRY,FR,cw-6*MM);
    }
    y+=2*26*MM+8*MM;
    y=SEC(doc,O,y,"Fases del Plan de Implementacion");
    var fases=[
        ["FASE 1","Semana 1","Pre-arranque y alta de personal","Alta IMSS, uniformes, credenciales, registro biometrico"],
        ["FASE 2","Semana 1-2","Instalacion tecnologica","CCTV, biometrico, radios, bitacora digital"],
        ["FASE 3","Semana 2","Capacitacion de ingreso","NOM-030 basico, induccion al cliente, protocolo de sitio"],
        ["FASE 4","Semana 2-3","Arranque operativo supervisado","Primer turno con supervisor MICSA en campo"],
        ["FASE 5","Semana 4","Auditoria de arranque","Revision de KPIs, ajuste de puestos, validacion cliente"]
    ];
    y=TBL(doc,O,y,
        ["#","FASE","PERIODO","ACTIVIDAD CLAVE","ENTREGABLE"],
        fases.map(function(f){ return [f[0],f[1],f[2],f[3]]; }),
        [0.08,0.16,0.16,0.34,0.26]);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB3 — RUTA CRITICA (DIAGRAMA DE GANTT)
// ═══════════════════════════════════════════════════════════════════
function rutaCritica(doc){
    var O=2;
    HDR(doc,O,"Ruta Critica — Diagrama de Gantt 8 Semanas",3);
    var y=22*MM;
    // Leyenda
    R(doc,O*LW+MRG,y,10*MM,6*MM,BLK,SLV,0.3);
    T(doc,"Activo",O*LW+MRG+12*MM,y+1*MM,6.5,DGRY,FR,25*MM);
    R(doc,O*LW+MRG+38*MM,y,10*MM,6*MM,WHT,MGRY,0.3);
    T(doc,"Inactivo / En espera",O*LW+MRG+50*MM,y+1*MM,6.5,DGRY,FR,40*MM);
    y+=10*MM;
    y=GANTT_HDR(doc,O,y);
    // Filas del Gantt
    // sem1..sem8: 1=activo, 0=inactivo
    y=GANTT_ROW(doc,O,y,"Alta IMSS del personal",       1,0,0,0,0,0,0,0);
    y=GANTT_ROW(doc,O,y,"Uniformes y equipamiento",     1,1,0,0,0,0,0,0);
    y=GANTT_ROW(doc,O,y,"Credenciales y biometrico",    1,1,0,0,0,0,0,0);
    y=GANTT_ROW(doc,O,y,"Instalacion CCTV",             1,1,0,0,0,0,0,0);
    y=GANTT_ROW(doc,O,y,"Config. radios y base",        0,1,0,0,0,0,0,0);
    y=GANTT_ROW(doc,O,y,"Capacitacion NOM-030 basico",  0,1,1,0,0,0,0,0);
    y=GANTT_ROW(doc,O,y,"Induccion al sitio cliente",   0,1,1,0,0,0,0,0);
    y=GANTT_ROW(doc,O,y,"1er. turno supervisado",       0,0,1,1,0,0,0,0);
    y=GANTT_ROW(doc,O,y,"Ajuste de posiciones",         0,0,0,1,0,0,0,0);
    y=GANTT_ROW(doc,O,y,"Auditoria de arranque",        0,0,0,1,0,0,0,0);
    y=GANTT_ROW(doc,O,y,"Operacion normal 24/7",        0,0,0,0,1,1,1,1);
    y=GANTT_ROW(doc,O,y,"Supervision tecnica MICSA",    0,0,0,0,1,0,1,0);
    y=GANTT_ROW(doc,O,y,"Reporte ejecutivo mes 1",      0,0,0,0,0,1,0,0);
    y=GANTT_ROW(doc,O,y,"Capacitacion NOM actualizacion",0,0,0,0,0,0,1,0);
    y+=3*MM;
    // Hitos
    R(doc,O*LW+MRG,y,LW-MRG*2,1*MM,SLV);
    y+=3*MM;
    T(doc,"HITOS CRITICOS",O*LW+MRG,y,7.5,BLK,FB,80*MM);
    y+=6*MM;
    var hitos=["DIA 3: Alta IMSS confirmada y personal con uniforme.",
               "DIA 10: Instalacion tecnologica completada y operativa.",
               "DIA 14: Primer turno en operacion sin supervision directa.",
               "DIA 30: Auditoria de arranque con representante del cliente firmada.",
               "DIA 60: Evaluacion de desempeno del 100% del personal activo."];
    for(var i=0;i<hitos.length;i++){
        T(doc,String.fromCharCode(9654)+"  "+hitos[i],O*LW+MRG,y,6.5,DGRY,FR,LW-MRG*2);
        y+=5.5*MM;
    }
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB4 — CHECKLIST DE ARRANQUE
// ═══════════════════════════════════════════════════════════════════
function checklistArranque(doc){
    var O=3;
    HDR(doc,O,"Checklist de Arranque Operativo",4);
    var y=22*MM;
    y=SEC(doc,O,y,"Semana 1 — Pre-Arranque (Dias 1 al 7)");
    y=TBL(doc,O,y,
        ["#","ACTIVIDAD","RESPONSABLE","VERIFICADO"],
        [
            ["1.1","Alta IMSS del 100% del personal contratado","RRHH MICSA","_________"],
            ["1.2","Entrega de uniformes completos por elemento","Coordinador","_________"],
            ["1.3","Emision de credenciales con foto y QR","Administracion","_________"],
            ["1.4","Registro biometrico (huella + foto) en sistema","TI / Coordinador","_________"],
            ["1.5","Entrega de radios Motorola programados","Coordinador","_________"],
            ["1.6","Instalacion de CCTV y DVR en sitio","TI MICSA","_________"],
            ["1.7","Configuracion de bitacora digital (acceso cliente)","TI MICSA","_________"],
            ["1.8","Firma de contrato laboral individual","RRHH / Director","_________"]
        ],
        [0.06,0.50,0.24,0.20]);
    y+=5*MM;
    y=SEC(doc,O,y,"Semana 2 — Capacitacion e Inicio de Operacion");
    y=TBL(doc,O,y,
        ["#","ACTIVIDAD","RESPONSABLE","VERIFICADO"],
        [
            ["2.1","Modulo NOM-030 basico (40 hrs)","Instructor certificado","_________"],
            ["2.2","Recorrido de induccion al sitio cliente","Supervisor + cliente","_________"],
            ["2.3","Entrega de planos y rutas de evacuacion","Coordinador","_________"],
            ["2.4","Simulacro de emergencia inicial","Supervisor","_________"],
            ["2.5","Primer turno completo operando (supervisado)","Supervisor MICSA","_________"],
            ["2.6","Verificacion de puntos de relevo y comunicacion","Supervisor","_________"]
        ],
        [0.06,0.50,0.24,0.20]);
    y+=5*MM;
    y=SEC(doc,O,y,"Mes 1 — Consolidacion");
    y=TBL(doc,O,y,
        ["#","ACTIVIDAD","RESPONSABLE","VERIFICADO"],
        [
            ["3.1","Auditoria de arranque con representante del cliente","Director / Coord.","_________"],
            ["3.2","Ajuste de posiciones segun evaluacion de campo","Coordinador","_________"],
            ["3.3","Entrega del primer reporte ejecutivo mensual","Coordinador","_________"],
            ["3.4","Evaluacion de desempeno de los primeros 30 dias","Supervisor","_________"]
        ],
        [0.06,0.50,0.24,0.20]);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB5 — PROTOCOLO DE EMERGENCIAS
// ═══════════════════════════════════════════════════════════════════
function protocoloEmergencias(doc){
    var O=4;
    HDR(doc,O,"Protocolo de Emergencias y Contingencias",5);
    var y=22*MM;
    y=SEC(doc,O,y,"Niveles de Alerta y Respuesta");
    y=TBL(doc,O,y,
        ["NIVEL","DENOMINACION","DESCRIPCION","TIEMPO RESPUESTA","ACCION"],
        [
            ["1","VERDE","Operacion normal, sin anomalias","N/A","Reporte de rutina"],
            ["2","AMARILLO","Situacion anormal en evaluacion","< 5 min","Supervisor activa protocolo"],
            ["3","NARANJA","Incidente confirmado en curso","< 3 min","Refuerzo + notificacion cliente"],
            ["4","ROJO","Emergencia activa (robo, agresion, incendio)","< 2 min","Alerta total + llamada 911"]
        ],
        [0.08,0.16,0.28,0.20,0.28]);
    y+=5*MM;
    y=SEC(doc,O,y,"Flujo de Activacion — Nivel Rojo");
    var pasos=[
        ["01","DETECCION","Elemento detecta emergencia — activa CODIGO 4 en radio"],
        ["02","CONFIRMACION","Supervisor confirma por radio en < 30 seg"],
        ["03","REFUERZO","Jefe de reaccion se desplaza al punto en < 2 min"],
        ["04","NOTIFICACION","Coordinador llama al contacto del cliente en < 3 min"],
        ["05","DOCUMENTACION","Supervisor inicia reporte de incidente en bitacora digital"],
        ["06","CIERRE","Director recibe resumen en < 15 min post-incidente"]
    ];
    var tw=LW-MRG*2;
    var pw=(tw-5*MM)/3;
    for(var i=0;i<pasos.length;i++){
        var col=i%3;
        var row=Math.floor(i/3);
        var px=O*LW+MRG+col*(pw+2.5*MM);
        var py=y+row*30*MM;
        R(doc,px,py,pw,26*MM,LGRY,MGRY,0.4);
        R(doc,px,py,pw,7*MM,BLK);
        R(doc,px,py,2*MM,26*MM,SLV);
        T(doc,pasos[i][0],px+4*MM,py+1.5*MM,8,SLV,FB,12*MM);
        T(doc,pasos[i][1],px+16*MM,py+1.5*MM,7,WHT,FB,pw-18*MM);
        T(doc,pasos[i][2],px+4*MM,py+9.5*MM,6.5,DGRY,FR,pw-7*MM);
    }
    y+=2*30*MM+8*MM;
    y=SEC(doc,O,y,"Contactos Clave de Emergencia (a definir con cliente)");
    y=TBL(doc,O,y,
        ["ROL","NOMBRE","TELEFONO","DISPONIBILIDAD"],
        [
            ["Contacto principal cliente","Por confirmar","Por confirmar","24/7"],
            ["Coordinador operativo MICSA","Por confirmar","Por confirmar","24/7"],
            ["Director de seguridad MICSA","Gerardo Guzman","Por confirmar","24/7"],
            ["Emergencias / 911","Policia, Bomberos, Proteccion Civil","911","24/7"]
        ],
        [0.30,0.24,0.24,0.22]);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB6 — INDICADORES DE DESEMPENO (KPIs)
// ═══════════════════════════════════════════════════════════════════
function kpis(doc){
    var O=5;
    HDR(doc,O,"Indicadores de Desempeno Operativo",6);
    var y=22*MM;
    y=SEC(doc,O,y,"Dashboard de KPIs Mensuales");
    y=TBL(doc,O,y,
        ["KPI","DEFINICION","META","FRECUENCIA","FUENTE"],
        [
            ["Tiempo respuesta incidentes","Minutos desde deteccion hasta atencion","< 3 min","Diaria","Bitacora digital"],
            ["Cobertura de puestos","% de puestos cubiertos sobre total asignado","100%","Diaria","Registro de turno"],
            ["Incidentes con protocolo correcto","% de incidentes atendidos segun protocolo","100%","Mensual","Reporte supervisor"],
            ["Puntualidad del personal","% de elementos puntuales en brief","> 98%","Diaria","Control de asistencia"],
            ["Reportes entregados","% de reportes entregados en tiempo","100%","Mensual","Sistema MICSA"],
            ["Satisfaccion del cliente","Calificacion de auditoria mensual del cliente","> 9.0 / 10","Mensual","Encuesta cliente"],
            ["Rotacion de personal","% de bajas sobre total del plantel","< 5%","Mensual","RRHH MICSA"],
            ["Incidentes graves (Nivel 3-4)","Numero de activaciones de alerta naranja/rojo","0 idealmente","Mensual","Bitacora digital"]
        ],
        [0.26,0.28,0.10,0.16,0.20]);
    y+=5*MM;
    y=SEC(doc,O,y,"Ciclo de Mejora Continua");
    R(doc,O*LW+MRG,y,LW-MRG*2,20*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,3*MM,20*MM,SLV);
    var ciclo=["MEDIR: Captura diaria de datos en bitacora digital y sistema MICSA.",
               "ANALIZAR: Revision semanal de KPIs por el coordinador operativo.",
               "ACTUAR: Ajuste de puestos, capacitacion o proceso segun desviacion detectada.",
               "REPORTAR: Entrega mensual de dashboard al cliente con plan de accion."];
    for(var i=0;i<ciclo.length;i++){
        T(doc,String(i+1)+".  "+ciclo[i],O*LW+MRG+5*MM,y+3*MM+i*4.5*MM,7,DGRY,FR,LW-MRG*2-8*MM);
    }
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB7 — CIERRE Y VALIDACION
// ═══════════════════════════════════════════════════════════════════
function cierreValidacion(doc){
    var O=6;
    HDR(doc,O,"Validacion del Plan e Inicio Formal",7);
    var y=22*MM;
    y=SEC(doc,O,y,"Condiciones de Inicio Formal");
    y=TBL(doc,O,y,
        ["CONDICION","ESTADO REQUERIDO","RESPONSABLE"],
        [
            ["Contrato de servicios firmado","Firmado por ambas partes","Director MICSA + cliente"],
            ["Primer pago recibido","Transferencia confirmada","Administracion MICSA"],
            ["Alta IMSS del 100% del personal","Confirmada en IMSS","RRHH MICSA"],
            ["Equipamiento entregado y firmado","Acta de entrega firmada","Coordinador"],
            ["Accesos al sitio habilitados","Credenciales activas","Cliente + Coordinador"],
            ["Brief inicial con cliente realizado","Minuta firmada","Director / Coordinador"]
        ],
        [0.40,0.36,0.24]);
    y+=8*MM;
    y=SEC(doc,O,y,"Firmas de Validacion del Plan");
    var tw=LW-MRG*2;
    var col1=O*LW+MRG;
    var col2=O*LW+MRG+tw*0.52;
    var cw1=tw*0.45;
    var cw2=tw*0.45;
    // Firma cliente
    R(doc,col1,y,cw1,38*MM,WHT,MGRY,0.4);
    R(doc,col1,y,cw1,6*MM,BLK);
    T(doc,"AUTORIZA — CLIENTE",col1+3*MM,y+1.5*MM,7,WHT,FB,cw1-6*MM);
    L(doc,col1+5*MM,y+27*MM,col1+cw1-5*MM,y+27*MM,MGRY,0.5);
    T(doc,"Firma y sello",col1+5*MM,y+29*MM,6.5,MGRY,FL,cw1-10*MM,"C");
    T(doc,"Nombre:",col1+4*MM,y+33*MM,6.5,DGRY,FL,cw1-8*MM);
    T(doc,"Fecha de inicio:",col1+4*MM,y+37.5*MM,6.5,DGRY,FL,cw1-8*MM);
    // Firma MICSA
    R(doc,col2,y,cw2,38*MM,WHT,MGRY,0.4);
    R(doc,col2,y,cw2,6*MM,BLK);
    T(doc,"ELABORO — MICSA",col2+3*MM,y+1.5*MM,7,WHT,FB,cw2-6*MM);
    L(doc,col2+5*MM,y+27*MM,col2+cw2-5*MM,y+27*MM,MGRY,0.5);
    T(doc,"Firma",col2+5*MM,y+29*MM,6.5,MGRY,FL,cw2-10*MM,"C");
    T(doc,"Gerardo Guzman Alvarado",col2+4*MM,y+33*MM,7.5,BLK,FB,cw2-8*MM);
    T(doc,"Director de Seguridad MICSA",col2+4*MM,y+37.5*MM,6.5,DGRY,FR,cw2-8*MM);
    y+=44*MM;
    // Nota de cierre
    R(doc,O*LW+MRG,y,tw,16*MM,BLK,SLV,0.4);
    R(doc,O*LW+MRG,y,3*MM,16*MM,SLV);
    T(doc,'"Puede faltar flujo, pero nunca posicion."',
      O*LW+MRG+5*MM,y+4*MM,11,WHT,FB,tw-8*MM);
    T(doc,"Gerardo Guzman Alvarado  —  MICSA Safety Division",
      O*LW+MRG+5*MM,y+13*MM,6.5,SLV,FL,tw-8*MM);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// MAIN
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
    var abNames = ["02_RESUMEN_EJECUTIVO","03_RUTA_CRITICA_GANTT","04_CHECKLIST_ARRANQUE","05_PROTOCOLO_EMERGENCIAS","06_KPIS_OPERATIVOS","07_CIERRE_VALIDACION"];
    for(var i=0;i<abNames.length;i++){
        doc.artboards.add([LW*(i+1),0,LW*(i+2),-LH]);
        doc.artboards[i+1].name = abNames[i];
    }
    portada(doc);
    resumenEjecutivo(doc);
    rutaCritica(doc);
    checklistArranque(doc);
    protocoloEmergencias(doc);
    kpis(doc);
    cierreValidacion(doc);
    var aiFile = new File(BASE + "MICSA_PLAN_IMPLEMENTACION.ai");
    var saveOpts = new IllustratorSaveOptions();
    saveOpts.compatibility = Compatibility.ILLUSTRATOR17;
    saveOpts.compressed = false;
    doc.saveAs(aiFile, saveOpts);
    try{
        var pdfFile = new File(OUT + "MICSA_PLAN_IMPLEMENTACION.pdf");
        var pdfOpts = new PDFSaveOptions();
        pdfOpts.compatibility = PDFCompatibility.ACROBAT7;
        pdfOpts.generateThumbnails = true;
        pdfOpts.preserveEditability = false;
        doc.saveAs(pdfFile, pdfOpts);
    }catch(e){}
    alert("MICSA_PLAN_IMPLEMENTACION generado.\nGuardado en: " + BASE);
})();
