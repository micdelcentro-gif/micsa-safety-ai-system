/**
 * MICSA SAFETY DIVISION — Manual de Cultura Organizacional
 * ExtendScript para Adobe Illustrator
 * 8 Artboards — Letter 8.5x11in — CMYK Negro + Plateado
 * USO: Archivo > Scripts > Otro script...
 */

var BASE = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\MICSA_Safety\\";
var LOGO = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\adobe\\ID-5608-20260329T034218Z-1-001\\ID-5608\\LOGO 7\\editable\\editable.ai";
var OUT  = BASE + "PDFs\\";

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
    R(doc,O*LW,0,LW,15*MM,BLK);
    R(doc,O*LW,15*MM,LW,1*MM,SLV);
    logo(doc,O*LW+MRG,1.5*MM,16*MM);
    T(doc,"MICSA SAFETY DIVISION",O*LW+MRG+18*MM,4*MM,8,WHT,FB,LW*0.5);
    T(doc,titulo,O*LW+MRG+18*MM,10*MM,6.5,SLV,FL,LW*0.55);
    T(doc,String(pg)+" / 8",(O+1)*LW-MRG-18,5.5*MM,8,SLV,FS,20);
}
function FTR(doc,O){
    R(doc,O*LW,LH-12*MM,LW,12*MM,BLK);
    R(doc,O*LW,LH-12*MM,LW,0.7*MM,SLV);
    T(doc,"DOCUMENTO CONFIDENCIAL  |  MICSA Safety Division  |  Monclova, Coahuila  |  micsasafety.com.mx",
      O*LW+MRG,LH-7*MM,5.5,MGRY,FL,LW-MRG*2);
}
function SEC(doc,O,y,titulo){
    T(doc,titulo,O*LW+MRG,y,11,BLK,FB,LW-MRG*2);
    y+=7*MM;
    L(doc,O*LW+MRG,y,O*LW+LW-MRG,y,SLV,0.8);
    return y+5*MM;
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
function PILAR(doc,x,y,w,num,titulo,desc){
    R(doc,x,y,w,40*MM,WHT,MGRY,0.4);
    R(doc,x,y,w,9*MM,BLK);
    R(doc,x,y,2*MM,40*MM,SLV);
    T(doc,"0"+num,x+4*MM,y+2*MM,7.5,SLV,FS,12*MM);
    T(doc,titulo,x+17*MM,y+2*MM,7,WHT,FB,w-20*MM);
    T(doc,desc,x+4*MM,y+11*MM,6.5,DGRY,FR,w-8*MM);
}
function VALOR(doc,x,y,w,ico,nombre,desc){
    R(doc,x,y,w,32*MM,WHT,MGRY,0.4);
    R(doc,x,y,2*MM,32*MM,SLV);
    T(doc,ico,x+4*MM,y+3*MM,16,SLV,FB,10*MM);
    T(doc,nombre,x+16*MM,y+3*MM,7.5,BLK,FB,w-20*MM);
    T(doc,desc,x+4*MM,y+15*MM,6.5,DGRY,FR,w-8*MM);
}

// ═══════════════════════════════════════════════════════════════════
// AB1 — PORTADA
// ═══════════════════════════════════════════════════════════════════
function portada(doc){
    var O=0;
    R(doc,O*LW,0,LW,LH,BLK);
    R(doc,O*LW,0,6*MM,LH,SLV);
    R(doc,(O+1)*LW-3*MM,0,3*MM,LH,SLV2);
    R(doc,O*LW+6*MM,LH*0.52,LW-6*MM,LH*0.38,WHT);
    R(doc,O*LW+6*MM,LH*0.52,LW-6*MM,1.5*MM,SLV);
    logo(doc,O*LW+LW/2-32*MM,LH*0.07,64*MM);
    R(doc,O*LW+10*MM,LH*0.27,55*MM,7*MM,SLV);
    T(doc,"IDENTIDAD Y CULTURA CORPORATIVA",O*LW+11*MM,LH*0.27+1.5*MM,6,BLK,FB,53*MM);
    T(doc,"CULTURA",O*LW+10*MM,LH*0.29+8*MM,30,WHT,FB,LW-14*MM);
    T(doc,"ORGANIZACIONAL",O*LW+10*MM,LH*0.29+42,22,WHT,FB,LW-14*MM);
    R(doc,O*LW+10*MM,LH*0.29+66,60*MM,1.2*MM,SLV);
    T(doc,"El caracter de la operacion define la calidad del resultado.",
      O*LW+10*MM,LH*0.29+72,9,SLV2,FL,LW-14*MM);
    T(doc,"VISION INTERNA",O*LW+10*MM,LH*0.55,7,DGRY,FB,40*MM);
    T(doc,"Modelo de Gestion Humana para Operaciones de Seguridad de Alto Rendimiento",
      O*LW+10*MM,LH*0.55+10,9,BLK,FB,LW*0.55);
    T(doc,"v2.0  |  Monclova, Coahuila  |  2026",O*LW+10*MM,LH*0.55+22,7,DGRY,FL,LW*0.4);
    L(doc,O*LW+LW*0.60,LH*0.54,O*LW+LW*0.60,LH*0.90,MGRY,0.5);
    T(doc,"Director de Seguridad",O*LW+LW*0.62,LH*0.56,7,DGRY,FL,LW*0.33);
    T(doc,"Gerardo Guzman Alvarado",O*LW+LW*0.62,LH*0.56+10,8.5,BLK,FB,LW*0.33);
    T(doc,"Fuerza Civil  |  Proteccion Civil  |  Bomberos",O*LW+LW*0.62,LH*0.56+21,7,DGRY,FR,LW*0.33);
    R(doc,O*LW,LH-12*MM,LW,12*MM,BLK);
    R(doc,O*LW,LH-12*MM,LW,0.7*MM,SLV);
    T(doc,"DOCUMENTO CONFIDENCIAL  |  MICSA Safety Division  |  micsasafety.com.mx",
      O*LW+MRG,LH-7*MM,5.5,MGRY,FL,LW-MRG*2);
}

// ═══════════════════════════════════════════════════════════════════
// AB2 — FILOSOFIA CORPORATIVA
// ═══════════════════════════════════════════════════════════════════
function filosofia(doc){
    var O=1;
    HDR(doc,O,"Filosofia Corporativa",2);
    var y=22*MM;
    // Mision
    R(doc,O*LW+MRG,y,LW-MRG*2,24*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,3*MM,24*MM,BLK);
    R(doc,O*LW+MRG+3*MM,y,LW-MRG*2-3*MM,6*MM,BLK);
    T(doc,"MISION",O*LW+MRG+5*MM,y+1.5*MM,7,SLV,FB,30*MM);
    T(doc,"Proveer servicios de seguridad patrimonial e industrial con estandares de excelencia operativa, formando guardias altamente capacitados que protejan personas, activos e instalaciones mediante disciplina, tecnologia y liderazgo responsable.",
      O*LW+MRG+5*MM,y+8*MM,7.5,DGRY,FR,LW-MRG*2-8*MM);
    y+=28*MM;
    // Vision
    R(doc,O*LW+MRG,y,LW-MRG*2,24*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,3*MM,24*MM,BLK);
    R(doc,O*LW+MRG+3*MM,y,LW-MRG*2-3*MM,6*MM,BLK);
    T(doc,"VISION",O*LW+MRG+5*MM,y+1.5*MM,7,SLV,FB,30*MM);
    T(doc,"Ser la empresa de seguridad privada mas confiable del norte de Mexico, reconocida por la formacion integral de su personal, la transparencia de sus operaciones y la capacidad de respuesta ante cualquier nivel de riesgo.",
      O*LW+MRG+5*MM,y+8*MM,7.5,DGRY,FR,LW-MRG*2-8*MM);
    y+=28*MM;
    y=SEC(doc,O,y,"Proposito Fundamental");
    R(doc,O*LW+MRG,y,LW-MRG*2,20*MM,BLK);
    R(doc,O*LW+MRG,y,3*MM,20*MM,SLV);
    T(doc,'"No cuidamos instalaciones. Controlamos operaciones."',
      O*LW+MRG+5*MM,y+4*MM,13,WHT,FB,LW-MRG*2-8*MM);
    T(doc,"Gerardo Guzman Alvarado, Director de Seguridad MICSA",
      O*LW+MRG+5*MM,y+16*MM,7,SLV,FL,LW-MRG*2-8*MM);
    y+=26*MM;
    y=SEC(doc,O,y,"Los 3 Compromisos Irrenunciables");
    var tw=LW-MRG*2;
    var cw=(tw-4*MM)/3;
    var xs=[O*LW+MRG, O*LW+MRG+cw+2*MM, O*LW+MRG+(cw+2*MM)*2];
    var labels=["DISCIPLINA\nOPERATIVA","FORMACION\nPERMANENTE","LEALTAD\nINSTITUCIONAL"];
    var descs=[
        "Cada elemento sigue protocolos sin excepcion. La disciplina no es opcional.",
        "La capacitacion continua es parte del contrato, no un beneficio adicional.",
        "La confidencialidad y fidelidad al cliente son la base de cada contrato."
    ];
    for(var i=0;i<3;i++){
        R(doc,xs[i],y,cw,32*MM,LGRY,MGRY,0.4);
        R(doc,xs[i],y,cw,8*MM,BLK);
        R(doc,xs[i],y,2*MM,32*MM,SLV);
        T(doc,String(i+1),xs[i]+4*MM,y+1.5*MM,9,SLV,FB,8*MM);
        T(doc,labels[i],xs[i]+14*MM,y+1.5*MM,7,WHT,FB,cw-16*MM);
        T(doc,descs[i],xs[i]+4*MM,y+10*MM,6.5,DGRY,FR,cw-8*MM);
    }
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB3 — VALORES Y CODIGO DE CONDUCTA
// ═══════════════════════════════════════════════════════════════════
function valoresCodigo(doc){
    var O=2;
    HDR(doc,O,"Valores y Codigo de Conducta",3);
    var y=22*MM;
    y=SEC(doc,O,y,"Valores Fundamentales MICSA");
    var tw=LW-MRG*2;
    var cw=(tw-3*MM)/2;
    var valores=[
        ["#","INTEGRIDAD","Actuamos con honestidad en cada decision, reporte e interaccion con el cliente."],
        ["*","RESPONSABILIDAD","Asumimos consecuencias de nuestros actos y protegemos lo que se nos confia."],
        ["!","PROFESIONALISMO","Uniforme impecable, lenguaje correcto, postura firme: somos la imagen del cliente."],
        ["@","LEALTAD","No divulgamos informacion operativa. El secreto profesional es ley."],
        ["$","VALOR","Actuamos ante el riesgo con calma y decision, nunca con imprudencia."],
        ["+","TRABAJO EN EQUIPO","El turno no termina hasta que el relevo esta en posicion y briefed."]
    ];
    for(var i=0;i<valores.length;i++){
        var col=i%2===0?O*LW+MRG:O*LW+MRG+cw+3*MM;
        var row=Math.floor(i/2);
        VALOR(doc,col,y+row*36*MM,cw,valores[i][0],valores[i][1],valores[i][2]);
    }
    y+=3*36*MM+8*MM;
    y=SEC(doc,O,y,"Codigo de Conducta — Conductas Prohibidas");
    y=TBL(doc,O,y,
        ["CONDUCTA PROHIBIDA","CONSECUENCIA INMEDIATA"],
        [
            ["Abandono de puesto sin relevo autorizado","Baja inmediata + reporte IMSS"],
            ["Uso de telefono celular durante turno activo","Llamada de atencion + acta administrativa"],
            ["Divulgar informacion del cliente a terceros","Rescision de contrato + accion legal"],
            ["Presentarse en estado no apto para el servicio","Suspension 3 dias + evaluacion medica"],
            ["Aceptar dinero o beneficios del cliente directo","Baja inmediata + posible denuncia penal"],
            ["Uso excesivo de la fuerza sin protocolo","Suspension + investigacion interna"]
        ],
        [0.58,0.42]);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB4 — PERFIL DEL GUARDIAN MICSA
// ═══════════════════════════════════════════════════════════════════
function perfilGuardian(doc){
    var O=3;
    HDR(doc,O,"Perfil del Guardian MICSA",4);
    var y=22*MM;
    y=SEC(doc,O,y,"Atributos del Elemento de Alto Rendimiento");
    var tw=LW-MRG*2;
    var cw=(tw-3*MM)/2;
    var pilares=[
        ["01","PRESENCIA Y PORTE","Uniforme completo en todo momento, cabello recortado, calzado boleado. La imagen proyecta autoridad antes de hablar."],
        ["02","COMUNICACION EFECTIVA","Reporte claro, conciso y sin ambiguedades. El radio es herramienta de trabajo, no de entretenimiento."],
        ["03","OBSERVACION SISTEMATICA","Detectar anomalias antes de que se conviertan en incidentes. Ver todo, registrar lo importante, reportar lo critico."],
        ["04","CONTROL EMOCIONAL","El guardian no eleva la voz. La calma ante la presion es senial de entrenamiento profesional."],
        ["05","CONOCIMIENTO DEL SITIO","Dominar planos, rutas de evacuacion, ubicacion de extintores, accesos de emergencia y contactos clave del cliente."],
        ["06","ACTITUD DE SERVICIO","El guardian es el primer contacto del cliente. Cortesia firme, no sumision ni arrogancia."]
    ];
    for(var i=0;i<pilares.length;i++){
        var col=i%2===0?O*LW+MRG:O*LW+MRG+cw+3*MM;
        var row=Math.floor(i/2);
        PILAR(doc,col,y+row*44*MM,cw,pilares[i][0],pilares[i][1],pilares[i][2]);
    }
    y+=3*44*MM+5*MM;
    y=SEC(doc,O,y,"Requisitos de Ingreso y Permanencia");
    y=TBL(doc,O,y,
        ["REQUISITO","INGRESO","PERMANENCIA ANUAL"],
        [
            ["Certificado medico","Examen completo pre-empleo","Actualizacion cada 12 meses"],
            ["Antecedentes penales","Carta de no antecedentes vigente","Verificacion cada 2 anos"],
            ["Capacitacion NOM-030","Modulo basico 40 horas","Actualizacion semestral"],
            ["Evaluacion psicometrica","Test de confianza + aptitud","Evaluacion anual"]
        ],
        [0.34,0.33,0.33]);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB5 — ESTRUCTURA ORGANIZACIONAL
// ═══════════════════════════════════════════════════════════════════
function estructura(doc){
    var O=4;
    HDR(doc,O,"Estructura Organizacional",5);
    var y=22*MM;
    y=SEC(doc,O,y,"Organigrama Operativo MICSA Safety Division");
    // Nivel 1 — Director
    var bw=60*MM, bh=14*MM;
    var cx=O*LW+LW/2-bw/2;
    R(doc,cx,y,bw,bh,BLK,SLV,0.5);
    R(doc,cx,y,3*MM,bh,SLV);
    T(doc,"DIRECTOR DE SEGURIDAD",cx+5*MM,y+2*MM,6.5,WHT,FB,bw-7*MM);
    T(doc,"Gerardo Guzman Alvarado",cx+5*MM,y+9*MM,7,SLV,FS,bw-7*MM);
    L(doc,O*LW+LW/2,y+bh,O*LW+LW/2,y+bh+6*MM,MGRY,0.5);
    y+=bh+6*MM;
    // Nivel 2 — Coordinador Operativo + Admin
    var n2w=58*MM, n2h=12*MM;
    var n2xs=[O*LW+MRG+10*MM, O*LW+LW-MRG-n2w-10*MM];
    var n2labels=["COORDINADOR\nOPERATIVO","ADMINISTRACION\nY RRHH"];
    L(doc,O*LW+MRG+10*MM+n2w/2,y,O*LW+LW-MRG-n2w/2-10*MM,y,MGRY,0.5);
    for(var i=0;i<2;i++){
        L(doc,n2xs[i]+n2w/2,y,n2xs[i]+n2w/2,y+n2h,MGRY,0.5);
        R(doc,n2xs[i],y,n2w,n2h,LGRY,MGRY,0.4);
        R(doc,n2xs[i],y,2*MM,n2h,SLV);
        T(doc,n2labels[i],n2xs[i]+4*MM,y+2.5*MM,6.5,BLK,FB,n2w-6*MM);
    }
    L(doc,O*LW+LW/2,y-6*MM+1,O*LW+LW/2,y,MGRY,0.5);
    y+=n2h+6*MM;
    // Nivel 3 — 4 supervisores
    var n3w=36*MM, n3h=10*MM, gap=(LW-MRG*2-n3w*4)/3;
    var n3labels=["SUPERVISOR\nTURNO A","SUPERVISOR\nTURNO B","SUPERVISOR\nTURNO C","JEFE DE\nREACCION"];
    var n3xs=[];
    for(var j=0;j<4;j++) n3xs.push(O*LW+MRG+j*(n3w+gap));
    for(var j=0;j<4;j++){
        L(doc,n3xs[j]+n3w/2,y,n3xs[j]+n3w/2,y+n3h,MGRY,0.5);
        R(doc,n3xs[j],y,n3w,n3h,LGRY,MGRY,0.4);
        R(doc,n3xs[j],y,2*MM,n3h,SLV);
        T(doc,n3labels[j],n3xs[j]+3*MM,y+1.5*MM,6,BLK,FS,n3w-5*MM);
    }
    L(doc,n3xs[0]+n3w/2,y,n3xs[3]+n3w/2,y,MGRY,0.5);
    y+=n3h+6*MM;
    // Nivel 4 — Guardias
    R(doc,O*LW+MRG,y,LW-MRG*2,10*MM,BLK);
    R(doc,O*LW+MRG,y,3*MM,10*MM,SLV);
    T(doc,"ELEMENTOS OPERATIVOS  —  Guardias, Vigilantes y Personal de Apoyo (21 elementos)",
      O*LW+MRG+5*MM,y+2.5*MM,7.5,WHT,FB,LW-MRG*2-8*MM);
    y+=20*MM;
    y=SEC(doc,O,y,"Descripcion de Roles Clave");
    y=TBL(doc,O,y,
        ["ROL","RESPONSABILIDAD PRINCIPAL","REPORTA A"],
        [
            ["Director de Seguridad","Estrategia, clientes clave, auditoria operativa","Direccion MICSA"],
            ["Coordinador Operativo","Supervision diaria de turnos y logistica","Director"],
            ["Supervisor de Turno","Briefing, rondas, incidentes en tiempo real","Coordinador"],
            ["Jefe de Reaccion","Protocolo de emergencia y refuerzo inmediato","Supervisor"],
            ["Elemento Operativo","Vigilancia, control de accesos, reportes","Supervisor de turno"]
        ],
        [0.24,0.46,0.30]);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB6 — COMUNICACION Y PROTOCOLO DE TURNO
// ═══════════════════════════════════════════════════════════════════
function comunicacion(doc){
    var O=5;
    HDR(doc,O,"Comunicacion y Protocolo de Turno",6);
    var y=22*MM;
    y=SEC(doc,O,y,"Protocolo de Inicio de Turno (BRIEF)");
    y=TBL(doc,O,y,
        ["PASO","ACCION","RESPONSABLE","DURACION"],
        [
            ["1","Formacion en punto de reunion asignado","Supervisor","5 min"],
            ["2","Revision de uniformes e identificaciones","Supervisor","3 min"],
            ["3","Briefing de novedades del turno anterior","Supervisor saliente","5 min"],
            ["4","Asignacion de puestos y equipos","Supervisor entrante","3 min"],
            ["5","Verificacion de comunicacion radial","Elemento / Supervisor","2 min"],
            ["6","Confirmacion de relevo en posicion","Cada elemento","1 min"]
        ],
        [0.06,0.42,0.28,0.24]);
    y+=5*MM;
    y=SEC(doc,O,y,"Codigo de Comunicacion Radial MICSA");
    y=TBL(doc,O,y,
        ["CODIGO","SIGNIFICADO","USO"],
        [
            ["CODIGO 1","Todo en orden, posicion normal","Reporte cada 30 min"],
            ["CODIGO 2","Situacion anormal, en evaluacion","Notificacion inmediata"],
            ["CODIGO 3","Requiero apoyo en posicion","Urgente — respuesta < 3 min"],
            ["CODIGO 4","Emergencia activa — protocolo rojo","Critico — alerta general"],
            ["CODIGO 5","Solicitud de informacion / consulta","Uso cuando se requiera"],
            ["FUERA DE SERVICIO","Abandono temporal autorizado de radio","Solo con permiso supervisor"]
        ],
        [0.20,0.44,0.36]);
    y+=5*MM;
    y=SEC(doc,O,y,"Protocolo de Cierre de Turno (RELEVO)");
    R(doc,O*LW+MRG,y,LW-MRG*2,28*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,3*MM,28*MM,SLV);
    var steps=["El elemento saliente permanece en posicion hasta que el entrante confirma relevo fisico.",
               "Se entrega bitacora de novedades firmada por ambos elementos.",
               "Cualquier incidente no cerrado se reporta verbalmente y queda registrado en sistema.",
               "El supervisor firma la hoja de relevo. Sin firma de supervisor, el turno no cierra."];
    for(var i=0;i<steps.length;i++){
        T(doc,String(i+1)+".  "+steps[i],O*LW+MRG+5*MM,y+3*MM+i*6.5*MM,7,DGRY,FR,LW-MRG*2-8*MM);
    }
    T(doc,'"Puede faltar flujo, pero nunca posicion."',
      O*LW+MRG+5*MM,y+30*MM,8.5,BLK,FB,LW-MRG*2-8*MM);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB7 — EVALUACION Y CRECIMIENTO
// ═══════════════════════════════════════════════════════════════════
function evaluacion(doc){
    var O=6;
    HDR(doc,O,"Evaluacion de Desempeno y Crecimiento",7);
    var y=22*MM;
    y=SEC(doc,O,y,"Matriz de Evaluacion Trimestral");
    y=TBL(doc,O,y,
        ["CRITERIO","PESO","INDICADOR MEDIBLE","META"],
        [
            ["Puntualidad y asistencia","20%","Registros de entrada/salida","0 tardanzas / mes"],
            ["Presentacion e imagen","15%","Revision de supervisor en brief","10 de 10 en auditorias"],
            ["Conocimiento del puesto","20%","Evaluacion escrita semestral","Minimo 80 / 100"],
            ["Reportes y bitacora","15%","Calidad y puntualidad del reporte","100% entregados en tiempo"],
            ["Reaccion ante incidentes","20%","Tiempo de respuesta + protocolo","< 3 min, protocolo correcto"],
            ["Actitud y trabajo en equipo","10%","Evaluacion 360 del supervisor","Sin quejas formales"]
        ],
        [0.28,0.10,0.38,0.24]);
    y+=5*MM;
    y=SEC(doc,O,y,"Plan de Carrera MICSA");
    y=TBL(doc,O,y,
        ["NIVEL","DENOMINACION","REQUISITOS","INCREMENTO SALARIAL"],
        [
            ["N1","Elemento en formacion","Ingreso + capacitacion basica","Base contractual"],
            ["N2","Elemento certificado","NOM-030 + 6 meses, evaluacion > 80","+ 8%"],
            ["N3","Elemento senior","18 meses + liderazgo de relevo","+ 15%"],
            ["N4","Jefe de equipo","3 anos + curso supervision","+ 25%"],
            ["N5","Supervisor de turno","5 anos + examen de supervisor","+ 40%"]
        ],
        [0.08,0.26,0.40,0.26]);
    y+=5*MM;
    y=SEC(doc,O,y,"Reconocimientos y Incentivos");
    R(doc,O*LW+MRG,y,LW-MRG*2,22*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,3*MM,22*MM,SLV);
    var incentivos=["Elemento del mes: bono de $1,500 MXN + reconocimiento en tablero operativo.",
                    "Cero incidentes en turno durante trimestre: bono colectivo de equipo.",
                    "Certificacion adicional aprobada: apoyo del 50% del costo del curso.",
                    "Recomendacion de cliente documentada: carta de reconocimiento en expediente."];
    for(var i=0;i<incentivos.length;i++){
        T(doc,String.fromCharCode(9632)+"  "+incentivos[i],O*LW+MRG+5*MM,y+3*MM+i*5.5*MM,7,DGRY,FR,LW-MRG*2-8*MM);
    }
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// AB8 — CIERRE INSTITUCIONAL
// ═══════════════════════════════════════════════════════════════════
function cierreInstitucional(doc){
    var O=7;
    HDR(doc,O,"Compromiso Institucional",8);
    var y=22*MM;
    y=SEC(doc,O,y,"Declaracion de Principios Operativos");
    // Bloque central
    R(doc,O*LW+MRG,y,LW-MRG*2,28*MM,BLK,SLV,0.4);
    R(doc,O*LW+MRG,y,3*MM,28*MM,SLV);
    T(doc,"MICSA Safety Division declara que cada elemento que viste este uniforme representa una promesa: la de actuar con disciplina, servir con integridad y responder con valor. Nuestra cultura no se impone — se vive en cada relevo, en cada reporte, en cada decision tomada bajo presion.",
      O*LW+MRG+5*MM,y+4*MM,8,WHT,FR,LW-MRG*2-8*MM);
    y+=32*MM;
    y=SEC(doc,O,y,"Indicadores de Cultura Sana — Medicion Mensual");
    y=TBL(doc,O,y,
        ["INDICADOR DE CULTURA","META","MEDICION"],
        [
            ["Rotacion de personal","< 5% mensual","Control de bajas RRHH"],
            ["Quejas formales del cliente","0 por mes","Bitacora de comunicacion"],
            ["Incidentes por protocolo incorrecto","0 por trimestre","Reporte de supervision"],
            ["Participacion en capacitaciones","100%","Lista de asistencia"],
            ["Reportes entregados en tiempo","100%","Sistema digital MICSA"],
            ["Evaluaciones aprobadas (> 80)","100% del personal activo","Registro de evaluaciones"]
        ],
        [0.44,0.24,0.32]);
    y+=5*MM;
    // Firma del Director
    var fw=LW-MRG*2;
    R(doc,O*LW+MRG,y,fw,36*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,3*MM,36*MM,SLV);
    T(doc,"AUTORIZACION Y VALIDEZ DEL DOCUMENTO",O*LW+MRG+5*MM,y+2.5*MM,7,BLK,FB,fw*0.6);
    T(doc,"Este manual tiene vigencia anual y es de observancia obligatoria para todo el personal MICSA Safety Division. La ignorancia de su contenido no exime de su cumplimiento.",
      O*LW+MRG+5*MM,y+10*MM,7.5,DGRY,FR,fw-8*MM);
    L(doc,O*LW+MRG+5*MM,y+26*MM,O*LW+MRG+90*MM,y+26*MM,MGRY,0.5);
    T(doc,"Gerardo Guzman Alvarado  —  Director de Seguridad MICSA",O*LW+MRG+5*MM,y+28*MM,7,BLK,FS,fw-8*MM);
    T(doc,"Monclova, Coahuila  |  2026",O*LW+MRG+5*MM,y+34*MM,6.5,DGRY,FL,fw*0.4);
    FTR(doc,O);
}

// ═══════════════════════════════════════════════════════════════════
// MAIN
// ═══════════════════════════════════════════════════════════════════
(function(){
    var docPreset = new DocumentPreset();
    docPreset.colorMode = DocumentColorSpace.CMYK;
    docPreset.units = RulerUnits.Points;
    docPreset.width = LW * 8;
    docPreset.height = LH;
    var doc = app.documents.addDocument(DocumentColorSpace.CMYK, docPreset);
    doc.artboards[0].artboardRect = [0, 0, LW, -LH];
    doc.artboards[0].name = "01_PORTADA";
    var abNames = ["02_FILOSOFIA","03_VALORES_CONDUCTA","04_PERFIL_GUARDIAN","05_ESTRUCTURA_ORG","06_COMUNICACION_TURNO","07_EVALUACION_CARRERA","08_CIERRE_INSTITUCIONAL"];
    for(var i=0;i<abNames.length;i++){
        doc.artboards.add([LW*(i+1),0,LW*(i+2),-LH]);
        doc.artboards[i+1].name = abNames[i];
    }
    portada(doc);
    filosofia(doc);
    valoresCodigo(doc);
    perfilGuardian(doc);
    estructura(doc);
    comunicacion(doc);
    evaluacion(doc);
    cierreInstitucional(doc);
    var aiFile = new File(BASE + "MICSA_CULTURA_COMPLETO.ai");
    var saveOpts = new IllustratorSaveOptions();
    saveOpts.compatibility = Compatibility.ILLUSTRATOR17;
    saveOpts.compressed = false;
    doc.saveAs(aiFile, saveOpts);
    try{
        var pdfFile = new File(OUT + "MICSA_CULTURA_COMPLETO.pdf");
        var pdfOpts = new PDFSaveOptions();
        pdfOpts.compatibility = PDFCompatibility.ACROBAT7;
        pdfOpts.generateThumbnails = true;
        pdfOpts.preserveEditability = false;
        doc.saveAs(pdfFile, pdfOpts);
    }catch(e){}
    alert("MICSA_CULTURA_COMPLETO generado.\nGuardado en: " + BASE);
})();
