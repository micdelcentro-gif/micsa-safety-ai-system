/**
 * MICSA SAFETY DIVISION — Manual de Seguridad Industrial
 * ExtendScript para Adobe Illustrator
 * 8 Artboards — Letter 8.5x11in — CMYK Negro + Plateado
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
function HDR(doc,O,titulo,pg,total){
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
    T(doc,"CONFIDENCIAL  |  Manual de Seguridad Industrial  |  micsasafety.com.mx",
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
    var x=O*LW+MRG,rh=9.5*MM,tw=LW-MRG*2,cx,cw,r,c,h;
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
function CARD(doc,O,x,y,w,h,cod,titulo,items){
    R(doc,x,y,w,h,WHT,MGRY,0.4);
    R(doc,x,y,w,9.5*MM,BLK);
    R(doc,x,y,2*MM,h,SLV);
    T(doc,cod,x+4*MM,y+2.5*MM,7,SLV,FS,20*MM);
    T(doc,titulo,x+25*MM,y+2.5*MM,6.8,WHT,FB,w-28*MM);
    for(var j=0;j<items.length;j++){
        T(doc,String.fromCharCode(9658)+"  "+items[j],x+4*MM,y+11.5*MM+j*7.5*MM,6.5,DGRY,FR,w-8*MM);
    }
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
    T(doc,"PAQUETE CORPORATIVO — NIVEL 5",O*LW+10*MM,LH*0.27+1.5*MM,6,WHT,FB,58*MM,"C");
    // Titulo
    T(doc,"MANUAL DE",O*LW+10*MM,LH*0.29+8*MM,30,WHT,FB,LW-14*MM);
    T(doc,"SEGURIDAD INDUSTRIAL",O*LW+10*MM,LH*0.29+42,22,WHT,FB,LW-14*MM);
    // Linea roja + subtitulo
    R(doc,O*LW+10*MM,LH*0.29+66,60*MM,1.5*MM,RED);
    T(doc,"Protocolos de Operacion  |  Control de Riesgos  |  Normatividad",
      O*LW+10*MM,LH*0.29+74,9,LGRY,FL,LW-14*MM);
    // Datos
    T(doc,"ESTANDARES MICSA",O*LW+10*MM,LH*0.55,7,RED,FB,35*MM);
    T(doc,"Manual Operativo para Plantas de Ensamble y Centros Logisticos",
      O*LW+10*MM,LH*0.55+10,9.5,BLK,FB,LW*0.55);
    T(doc,"Version 2026  |  Documento Maestro de Operaciones",O*LW+10*MM,LH*0.55+24,7,DGRY,FL,LW*0.4);
    L(doc,O*LW+LW*0.60,LH*0.54,O*LW+LW*0.60,LH*0.90,RED,0.5);
    T(doc,"Director de Seguridad",O*LW+LW*0.62,LH*0.56,7,DGRY,FL,LW*0.33);
    T(doc,"Gerardo Guzman Alvarado",O*LW+LW*0.62,LH*0.56+10,8.5,BLK,FB,LW*0.33);
    T(doc,"Certificacion NFPA  |  STPS DC-3",O*LW+LW*0.62,LH*0.56+21,7,DGRY,FR,LW*0.33);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB2 — OBJETO + MARCO LEGAL
// ═══════════════════════════════════════════════════════════════════
function marcoLegal(doc){
    var O=1;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"OBJETO Y MARCO LEGAL",2);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"01. OBJETO Y ALCANCE");
    T(doc,"El presente Manual establece lineamientos, procedimientos y estandares de conducta para el personal de seguridad industrial de MICSA Safety Division. De observancia obligatoria desde el primer dia de induccion. Aplica a: Oficiales de seguridad, Supervisores, Jefes de turno, Monitoreo CCTV y Chofer-Logistica.",
      O*LW+MRG,y,8,DGRY,FL,LW-MRG*2); y+=18*MM;
    y=SEC(doc,O,y,"02. MARCO LEGAL Y NORMATIVO");
    var normas=[
        ["NOM-030-STPS-2009","Servicios preventivos de seguridad y salud","Toda la operacion"],
        ["NOM-017-STPS-2008","EPP — seleccion, uso y mantenimiento","Planta y bodega"],
        ["NOM-026-STPS-2008","Colores y senales de seguridad e higiene","Senaletica interna"],
        ["NOM-002-STPS-2010","Prevencion y proteccion contra incendios","Protocolos emergencia"],
        ["Ley Federal del Trabajo","Relaciones laborales, jornada, derechos","Todo el personal"],
        ["Reglamento IMSS","Seguridad social y cuotas patronales","Nomina y altas IMSS"],
        ["Ley Seguridad Privada","Registro y operacion de personal","Alta en empresa"],
        ["Reglamento interno cliente","Normativa especifica del sitio","Sitio del cliente"],
    ];
    y=TBL(doc,O,y,["NORMA / LEY","DESCRIPCION","APLICA EN"],normas,[0.25,0.48,0.27]);
    y+=5*MM;
    R(doc,O*LW+MRG,y,LW-MRG*2,14*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,2*MM,14*MM,BLK);
    T(doc,"CRITERIO: El incumplimiento a cualquiera de estas normas genera responsabilidad directa para el personal infractor y para la empresa. La ignorancia de la norma no exime de la sancion.",
      O*LW+MRG+5*MM,y+3.5*MM,7.5,DGRY,FR,LW-MRG*2-8*MM);
}

// ═══════════════════════════════════════════════════════════════════
// AB3 — PERFIL DEL PERSONAL
// ═══════════════════════════════════════════════════════════════════
function perfil(doc){
    var O=2;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"PERFIL DEL PERSONAL",3);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"03. PERFIL Y REQUISITOS DEL PERSONAL");
    var req=[
        ["Escolaridad","Secundaria concluida minima. Bachillerato preferente."],
        ["Experiencia","Minimo 1 ano en seguridad industrial o patrimonial documentada."],
        ["Documentacion","INE vigente, CURP, comprobante domicilio, cartilla militar (H), NSS."],
        ["Antecedentes","Sin antecedentes penales. Certificado de no antecedentes obligatorio."],
        ["Condicion fisica","Examen medico aprobado. Indice de masa corporal en rango aceptable."],
        ["Psicometria",">=80% en honestidad, confiabilidad y toma de decisiones bajo presion."],
        ["Disponibilidad","Turnos rotativos 12h. A: 07:00-19:00 / B: 19:00-07:00. Fines de semana."],
        ["Licencia","Indispensable para Chofer/Logistica y uso de cuatrimoto certificado."],
    ];
    y=TBL(doc,O,y,["REQUISITO","ESPECIFICACION"],req,[0.28,0.72]);
    y+=6*MM;
    y=SEC(doc,O,y,"PERFIL POR PUESTO");
    var puestos=[
        ["Oficial de Seguridad","Secundaria","1 ano","Control accesos + bitacora"],
        ["Monitor CCTV","Preparatoria","1 ano + TI basica","Atencion sostenida + reporte"],
        ["Supervisor de Turno","Preparatoria","3 anos","Liderazgo + toma de decisiones"],
        ["Chofer / Logistica","Preparatoria + Lic.","2 anos","Manejo defensivo + logistica"],
        ["Jefe de Seguridad","Licenciatura","5 anos","Gestion operativa + cliente"],
    ];
    y=TBL(doc,O,y,["PUESTO","ESCOLARIDAD","EXP. MINIMA","HABILIDAD CLAVE"],puestos,[0.30,0.20,0.20,0.30]);
}

// ═══════════════════════════════════════════════════════════════════
// AB4 — PROCESO DE RECLUTAMIENTO
// ═══════════════════════════════════════════════════════════════════
function reclutamiento(doc){
    var O=3;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"PROCESO DE RECLUTAMIENTO",4);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"04. PROCESO DE RECLUTAMIENTO Y SELECCION");
    T(doc,"El proceso garantiza la contratacion de personal apto, integro y comprometido. No se aceleran fases por urgencia operativa. La seguridad del proceso de seleccion es la primera linea de seguridad.",
      O*LW+MRG,y,8,DGRY,FL,LW-MRG*2); y+=13*MM;

    var fases=[
        {n:"01",t:"PRE-FILTRO\nRH",d:"Dias 1",
         items:["Revision CV + verificacion datos","Llamada de validacion de actitud","Confirmacion disponibilidad real","Explicacion condiciones laborales"]},
        {n:"02",t:"ENTREVISTA\nESTRUCTURADA",d:"Dia 2",
         items:["Preguntas de reaccion y etica","Toma de decisiones bajo presion","Escenarios de seguridad reales","Verificacion de actitud y disciplina"]},
        {n:"03",t:"EVALUACION\nTECNICA",d:"Dia 3",
         items:["Simulacion control de acceso","Reporte de incidente en campo","Uso correcto de EPP NOM-017","Protocolo de rondines basico"]},
        {n:"04",t:"EVALUACION\nFISICA",d:"Dia 4",
         items:["Examen medico pre-empleo","Resistencia basica verificada","Indice de masa corporal","Aptitud psicometrica >=80%"]},
        {n:"05",t:"INVESTIGACION\nY CIERRE",d:"Dias 5-7",
         items:["Antecedentes penales verificados","Referencias laborales directas","Estudio socioeconomico domicilio","Decision: alta IMSS dia 1"]}
    ];
    var fw=LW-MRG*2,fh=42*MM;
    for(var i=0;i<fases.length;i++){
        var f=fases[i];
        var fx=O*LW+MRG,fy=y+i*(fh+3*MM);
        R(doc,fx,fy,fw,fh,WHT,MGRY,0.4);
        R(doc,fx,fy,25*MM,fh,BLK);
        T(doc,f.n,fx+1*MM,fy+5*MM,16,SLV,FB,23*MM,"C");
        T(doc,f.t,fx+1*MM,fy+19*MM,7,WHT,FS,23*MM,"C");
        T(doc,f.d,fx+1*MM,fy+30*MM,6.5,SLV2,FL,23*MM,"C");
        L(doc,fx+25*MM,fy+5*MM,fx+25*MM,fy+fh-5*MM,SLV,0.6);
        var iw=(fw-29*MM)/2-2*MM;
        for(var j=0;j<f.items.length;j++){
            var ix=fx+27*MM+Math.floor(j/2)*(iw+2*MM);
            var iy=fy+5*MM+(j%2)*16*MM;
            R(doc,ix,iy+2*MM,2*MM,7*MM,SLV);
            T(doc,f.items[j],ix+4*MM,iy+2.5*MM,7,DGRY,FR,iw-2*MM);
        }
        if(i<fases.length-1){
            L(doc,fx+fw/2,fy+fh,fx+fw/2,fy+fh+3*MM,MGRY,1);
        }
    }
}

// ═══════════════════════════════════════════════════════════════════
// AB5 — PROTOCOLOS OPERATIVOS
// ═══════════════════════════════════════════════════════════════════
function protocolos(doc){
    var O=4;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"PROTOCOLOS OPERATIVOS",5);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"06. PROTOCOLOS OPERATIVOS");
    T(doc,"Cada protocolo es de ejecucion obligatoria. No se interpretan — se ejecutan exactamente como estan definidos.",
      O*LW+MRG,y,8,DGRY,FL,LW-MRG*2); y+=11*MM;

    // Control de accesos
    y=SEC(doc,O,y,"6.1 CONTROL DE ACCESOS");
    var accesos=[
        ["01","Solicitar identificacion oficial vigente","INE, credencial empresa o pase del cliente — sin excepcion"],
        ["02","Validar en lista autorizada o con monitoreo","Comunicacion por radio si hay duda — no improvisar"],
        ["03","Registrar en bitacora con hora y destino","Sin excepcion, incluyendo directivos y visitas VIP"],
        ["04","Asignar gafete de visitante si aplica","Recuperar al salir. Documentar en bitacora de salida"],
        ["05","Notificar al receptor interno","Nunca dejar visitante sin acompanamiento autorizado"],
    ];
    y=TBL(doc,O,y,["PASO","ACCION","OBSERVACION"],accesos,[0.06,0.40,0.54]);
    y+=5*MM;

    // Rondines
    y=SEC(doc,O,y,"6.2 RONDINES Y REPORTE DE INCIDENTES");
    var rondines=[
        ["Frecuencia rondines","Cada 30 min perimetro. Cada 60 min bodega. Sin omisiones."],
        ["Registro obligatorio","Firma fisica o QR en punto de control. Supervisor valida."],
        ["Anomalia detectada","Reportar inmediato a supervisor. No resolver sin autorizacion."],
        ["Reporte de incidente","Notificar 0-5 min. Reporte escrito en <=2 horas. Evidencia."],
        ["Uso de cuatrimoto","Solo personal certificado. Vel. max 20 km/h en instalaciones."],
        ["EPP obligatorio","Chaleco, casco, calzado de seguridad en todo momento de turno."],
    ];
    y=TBL(doc,O,y,["PROTOCOLO","DESCRIPCION OPERATIVA"],rondines,[0.28,0.72]);
    y+=5*MM;

    // Emergencias
    y=SEC(doc,O,y,"6.3 PROTOCOLO DE EMERGENCIA");
    var emerg=[
        ["0-2 min","Detectar y contener el area","Guardia en turno"],
        ["2-5 min","Notificar a supervisor por radio","Guardia en turno"],
        ["5-10 min","Activar protocolo, notificar al cliente","Supervisor"],
        ["<=2 horas","Reporte escrito completo con evidencia","Supervisor"],
        ["24 horas","Seguimiento, cierre y recomendacion","Jefe Seguridad"],
    ];
    TBL(doc,O,y,["TIEMPO","ACCION","RESPONSABLE"],emerg,[0.15,0.55,0.30]);
}

// ═══════════════════════════════════════════════════════════════════
// AB6 — PLAN DE CAPACITACION
// ═══════════════════════════════════════════════════════════════════
function capacitacion(doc){
    var O=5;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"PLAN DE CAPACITACION",6);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"07. PLAN DE CAPACITACION CONTINUA — 12 MESES");
    T(doc,"Minimo 1 sesion mensual por guardia. La capacitacion no es opcional — es la promesa de calidad de MICSA al cliente.",
      O*LW+MRG,y,8,DGRY,FL,LW-MRG*2); y+=11*MM;
    var plan=[
        ["01","Induccion y cultura","Valores, protocolos, codigo de conducta MICSA","8h","Jefe Seguridad"],
        ["02","Control de accesos","Tecnicas, registros, escenarios practicos en campo","4h","Supervisor"],
        ["03","Primeros auxilios","RCP, hemorragias, traslado de heridos certificado","4h","Externo cert."],
        ["04","Contra incendio","Tipos de fuego, extintor, evacuacion NOM-002","4h","Externo cert."],
        ["05","Comunicacion operativa","Radio, reportes, lenguaje neutro y cadena de mando","2h","Supervisor"],
        ["06","Evaluacion semestral","Examen teorico + practica certificada en campo","4h","Jefe Seguridad"],
        ["07","Deteccion de amenazas","Perfil de riesgo, observacion situacional avanzada","3h","Jefe Seguridad"],
        ["08","Etica e integridad","Casos reales, dilemas, valores MICSA — director","2h","Director"],
        ["09","Manejo cuatrimoto","Tecnica, mantenimiento, seguridad vial certificada","4h","Inst. externo"],
        ["10","Reporte de incidentes","Formato oficial, redaccion, evidencia fotografica","2h","Supervisor"],
        ["11","NOM compliance","NOM-030, NOM-017, actualizacion legal 2026","3h","RH / Legal"],
        ["12","Evaluacion anual","Examen integral + renovacion de acreditacion operativa","8h","Jefe Seguridad"],
    ];
    y=TBL(doc,O,y,["MES","MODULO","CONTENIDO","HRS","INSTRUCTOR"],plan,[0.06,0.22,0.48,0.08,0.16]);
    y+=5*MM;
    R(doc,O*LW+MRG,y,LW-MRG*2,14*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,2*MM,14*MM,BLK);
    T(doc,"CRITERIO DE APROBACION: Puntaje >=80% en cada evaluacion. Elemento reprobado = baja sin recontatacion en 90 dias. La capacitacion genera expediente auditavle por STPS.",
      O*LW+MRG+5*MM,y+3.5*MM,7.5,DGRY,FR,LW-MRG*2-8*MM);
}

// ═══════════════════════════════════════════════════════════════════
// AB7 — KPI'S Y DASHBOARD
// ═══════════════════════════════════════════════════════════════════
function kpis(doc){
    var O=6;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"KPI'S Y DASHBOARD",7);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"08. KPI'S DE DESEMPENO — METRICAS MENSUALES");
    T(doc,"Dashboard entregado al cliente el primer dia de cada mes. Los KPIs son publicos, objetivos y auditables. Sin excusas — solo numeros.",
      O*LW+MRG,y,8,DGRY,FL,LW-MRG*2); y+=11*MM;

    // Metricas visuales 3x2
    var M=[
        {n:"97%",t:"PUNTUALIDAD",d:"Meta mensual minima.\n3 faltas = baja automatica."},
        {n:"100%",t:"BITACORAS",d:"Todos los turnos registrados.\nSin excepcion posible."},
        {n:"95%",t:"RONDINES",d:"Cumplimiento minimo mensual.\nRegistro QR verificado."},
        {n:"0",t:"INCIDENTES NO REPORTADOS",d:"Cero tolerancia absoluta.\nOcultar = baja inmediata."},
        {n:"<=5%",t:"ROTACION PERSONAL",d:"Objetivo mensual de\npermanencia del equipo."},
        {n:"1",t:"CAPACITACION / MES",d:"Sesion minima mensual\npor cada guardia asignado."}
    ];
    var mw=(LW-MRG*2-4*MM)/3,mh=32*MM;
    for(var i=0;i<M.length;i++){
        var mx=O*LW+MRG+(i%3)*(mw+2*MM);
        var my=y+Math.floor(i/3)*(mh+3*MM);
        R(doc,mx,my,mw,mh,WHT,MGRY,0.4);
        R(doc,mx,my,mw,2*MM,BLK);
        T(doc,M[i].n,mx+3*MM,my+4*MM,18,BLK,FB,mw-6*MM);
        T(doc,M[i].t,mx+3*MM,my+18*MM,6.5,MGRY,FS,mw-6*MM);
        L(doc,mx+3*MM,my+22*MM,mx+mw-3*MM,my+22*MM,LGRY,0.5);
        T(doc,M[i].d,mx+3*MM,my+24*MM,6.5,DGRY,FR,mw-6*MM);
    }
    y+=2*mh+3*MM+8*MM;

    y=SEC(doc,O,y,"DASHBOARD CONTROL MENSUAL");
    var krow=[
        ["Puntualidad (%)",">=97%","___","___","___","___","___"],
        ["Bitacoras completas (%)","100%","___","___","___","___","___"],
        ["Rondines cumplidos (%)",">=95%","___","___","___","___","___"],
        ["Incidentes no reportados","0","___","___","___","___","___"],
        ["Rotacion personal (%)","<=5%","—","—","—","___","___"],
        ["EPP en buen estado (%)","100%","___","___","___","___","___"],
    ];
    y=TBL(doc,O,y,["KPI","META","SEM 1","SEM 2","SEM 3","SEM 4","ACCION"],krow,[0.28,0.10,0.10,0.10,0.10,0.10,0.22]);
    y+=5*MM;

    // Firmas
    var sw=(LW-MRG*2)*0.45;
    R(doc,O*LW+MRG,y,sw,20*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,sw,1.5*MM,BLK);
    T(doc,"___________________________",O*LW+MRG+5*MM,y+11*MM,8,DGRY,FR,sw-10*MM);
    T(doc,"Supervisor responsable del mes",O*LW+MRG+5*MM,y+16*MM,6.5,DGRY,FL,sw-10*MM);
    R(doc,O*LW+LW-MRG-sw,y,sw,20*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+LW-MRG-sw,y,sw,1.5*MM,BLK);
    T(doc,"___________________________",O*LW+LW-MRG-sw+5*MM,y+11*MM,8,DGRY,FR,sw-10*MM);
    T(doc,"Vo.Bo. Jefe de Seguridad — Gerardo Guzman A.",O*LW+LW-MRG-sw+5*MM,y+16*MM,6.5,DGRY,FL,sw-10*MM);
}

// ═══════════════════════════════════════════════════════════════════
// AB8 — REPORTE DE INCIDENTE
// ═══════════════════════════════════════════════════════════════════
function reporteIncidente(doc){
    var O=7;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"FORMATO REPORTE DE INCIDENTE",8);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"10. FORMATO OFICIAL — REPORTE DE INCIDENTE");
    // Encabezado del reporte
    var meta=[
        ["Folio:","INC-___-2026","Fecha:","___ / ___ / 2026"],
        ["Hora incidente:","___:___","Hora reporte:","___:___"],
        ["Instalacion / Area:","_______________","Turno:","A  /  B"],
        ["Guardia reportante:","_______________","Supervisor:","_______________"],
    ];
    y=TBL(doc,O,y,["CAMPO","DATOS","CAMPO","DATOS"],meta,[0.18,0.32,0.18,0.32]);
    y+=5*MM;

    // Descripcion
    y=SEC(doc,O,y,"DESCRIPCION DEL INCIDENTE");
    R(doc,O*LW+MRG,y,LW-MRG*2,22*MM,LGRY,MGRY,0.4);
    T(doc,"Describa detalladamente lo ocurrido (que, quien, donde, como, cuando):",
      O*LW+MRG+3*MM,y+3*MM,7,MGRY,FL,LW-MRG*2-6*MM);
    y+=27*MM;

    // Clasificacion
    var pw=(LW-MRG*2)*0.48,ph=50*MM;
    // Col izq — Clasificacion
    var cx=O*LW+MRG,cy=y;
    R(doc,cx,cy,pw,ph,WHT,MGRY,0.4);
    R(doc,cx,cy,pw,9*MM,BLK);
    T(doc,"CLASIFICACION DEL INCIDENTE",cx+3*MM,cy+2.5*MM,7,WHT,FB,pw-6*MM);
    var nivsY=cy+11*MM;
    var nivs=[
        {col:LGRY,t:"LEVE — Advertencia",d:"Retardo, uniforme incompleto, omision menor"},
        {col:SLV2,t:"MODERADA — Suspension",d:"Reincidencia, celular, ronda incompleta. 1-3 dias sin goce"},
        {col:BLK,t:"GRAVE — Baja inmediata",d:"Abandono, sustancias, agresion, robo. Rescision contrato"}
    ];
    for(var i=0;i<nivs.length;i++){
        R(doc,cx+2*MM,nivsY+i*12*MM,pw-4*MM,11*MM,nivs[i].col,MGRY,0.3);
        T(doc,nivs[i].t,cx+4*MM,nivsY+i*12*MM+2.5*MM,7.5,i===2?WHT:BLK,FB,pw-8*MM);
        T(doc,nivs[i].d,cx+4*MM,nivsY+i*12*MM+8*MM,6.5,i===2?SLV2:DGRY,FR,pw-8*MM);
    }
    // Col der — Evidencia
    var dx=O*LW+LW-MRG-pw,dy=y;
    R(doc,dx,dy,pw,ph,WHT,MGRY,0.4);
    R(doc,dx,dy,pw,9*MM,BLK);
    T(doc,"EVIDENCIA ADJUNTA",dx+3*MM,dy+2.5*MM,7,WHT,FB,pw-6*MM);
    var evidY=dy+11*MM;
    var evid=["Fotografias (cantidad: _____)","Grabacion CCTV (tiempo: _____)","Declaracion escrita testigos","Otro: _____________________"];
    for(var j=0;j<evid.length;j++){
        R(doc,dx+2*MM,evidY+j*10*MM,pw-4*MM,9*MM,LGRY,MGRY,0.3);
        R(doc,dx+2*MM,evidY+j*10*MM,2*MM,9*MM,SLV);
        T(doc,"[ ]  "+evid[j],dx+6*MM,evidY+j*10*MM+2.5*MM,7,DGRY,FR,pw-10*MM);
    }
    y+=ph+8*MM;

    // Firmas
    var sw=(LW-MRG*2)*0.45;
    R(doc,O*LW+MRG,y,sw,22*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,sw,1.5*MM,BLK);
    T(doc,"___________________________",O*LW+MRG+5*MM,y+12*MM,8,DGRY,FR,sw-10*MM);
    T(doc,"Guardia reportante",O*LW+MRG+5*MM,y+17.5*MM,6.5,DGRY,FL,sw-10*MM);
    R(doc,O*LW+LW-MRG-sw,y,sw,22*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+LW-MRG-sw,y,sw,1.5*MM,BLK);
    T(doc,"___________________________",O*LW+LW-MRG-sw+5*MM,y+12*MM,8,DGRY,FR,sw-10*MM);
    T(doc,"Supervisor / Jefe de Seguridad",O*LW+LW-MRG-sw+5*MM,y+17.5*MM,6.5,DGRY,FL,sw-10*MM);
    y+=28*MM;
    // Cierre
    R(doc,O*LW+MRG,y,LW-MRG*2,16*MM,BLK);
    R(doc,O*LW+MRG,y,LW-MRG*2,1.2*MM,SLV);
    T(doc,"MICSA SAFETY DIVISION — Manual de Seguridad Industrial v2.0 — Uso Interno Confidencial",
      O*LW+MRG+5*MM,y+4.5*MM,9,WHT,FB,LW-MRG*2-10*MM,"C");
    R(doc,O*LW+LW/2-24*MM,y+13*MM,48*MM,0.7*MM,SLV);
    T(doc,'"Disciplina, criterio y control operativo"',O*LW+MRG,y+15.5*MM,9,SLV,FL,LW-MRG*2,"C");
}

// ═══════════════════════════════════════════════════════════════════
// MAIN
// ═══════════════════════════════════════════════════════════════════
function main(){
    var doc=app.documents.add(DocumentColorSpace.CMYK,LW*8,LH);
    for(var i=1;i<8;i++) doc.artboards.add([i*LW,0,(i+1)*LW,-LH]);
    var N=["01_PORTADA","02_MARCO_LEGAL","03_PERFIL","04_RECLUTAMIENTO",
           "05_PROTOCOLOS","06_CAPACITACION","07_KPIS","08_REPORTE_INCIDENTE"];
    for(var i=0;i<8;i++) doc.artboards[i].name=N[i];

    portada(doc);
    marcoLegal(doc);
    perfil(doc);
    reclutamiento(doc);
    protocolos(doc);
    capacitacion(doc);
    kpis(doc);
    reporteIncidente(doc);

    var ai=new File(BASE+"MICSA_MANUAL_SEGURIDAD.ai");
    var ao=new IllustratorSaveOptions();
    ao.compatibility=Compatibility.ILLUSTRATOR24;
    ao.saveMultipleArtboards=true;
    doc.saveAs(ai,ao);

    try{
        var pdf=new File(OUT+"MICSA_MANUAL_SEGURIDAD.pdf");
        var po=new PDFSaveOptions();
        po.compatibility=PDFCompatibility.ACROBAT8;
        po.generateThumbnails=true;
        po.preserveEditability=false;
        po.saveMultipleArtboards=true;
        po.artboardRange="1-8";
        doc.saveAs(pdf,po);
    }catch(e){}

    alert("MICSA_MANUAL_SEGURIDAD.ai generado.\n8 artboards: Portada / Marco Legal / Perfil / Reclutamiento / Protocolos / Capacitacion / KPIs / Reporte Incidente\n\nPaleta: Negro Rico + Plateado K25\nPDF exportado en PDFs/");
}
main();
