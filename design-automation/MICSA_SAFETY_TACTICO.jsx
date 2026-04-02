/**
 * MICSA SAFETY DIVISION â€” PresentaciÃ³n TÃ¡ctica
 * ExtendScript para Adobe Illustrator
 * 6 Artboards â€” Letter 8.5Ã—11in â€” CMYK Negro + Plateado
 * USO: Archivo > Scripts > Otro script...
 */

var BASE = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\MICSA_Safety\\";
var LOGO = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\MICSA_Safety\\logo.png";
var OUT  = BASE + "PDFs\\";

var MM  = 2.834645669;
var LW  = 8.5  * 72;   // 612 pt
var LH  = 11   * 72;   // 792 pt
var MRG = 20   * MM;   // 56.69 pt

function C(c,m,y,k){ var x=new CMYKColor(); x.cyan=c; x.magenta=m; x.yellow=y; x.black=k; return x; }
var BLK   = C(60,40,40,100);  // Negro Rico â€” autoridad
var SLV   = C(0,0,0,25);      // Plateado K25 â€” premium
var SLV2  = C(0,0,0,12);      // Plateado claro â€” fondos
var DGRY  = C(0,0,0,70);      // Gris oscuro â€” texto
var MGRY  = C(0,0,0,30);      // Gris medio â€” separadores
var LGRY  = C(0,0,0,7);       // Gris muy claro â€” fondos
var WHT   = C(0,0,0,0);       // Blanco

var FB = "Montserrat-Bold";
var FS = "Montserrat-SemiBold";
var FR = "OpenSans-Regular";
var FL = "OpenSans-Light";

// â”€â”€â”€ HELPERS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
    T(doc,String(pg)+" / 6",(O+1)*LW-MRG-14,5.5*MM,8,SLV,FS,16);
}
function FTR(doc,O){
    R(doc,O*LW,LH-12*MM,LW,12*MM,BLK);
    R(doc,O*LW,LH-12*MM,LW,0.7*MM,SLV);
    T(doc,"MICSA Safety Division  |  Seguridad Patrimonial Industrial  |  contacto@micsa.mx  |  800-MICSA-01",
      O*LW+MRG,LH-7*MM,6,MGRY,FL,LW-MRG*2);
}
function TBL(doc,O,y,hdrs,rows,cols){
    var x=O*LW+MRG,rh=10*MM,tw=LW-MRG*2,cx,cw,r,c,h;
    cx=x;
    for(h=0;h<hdrs.length;h++){
        cw=cols[h]*tw;
        R(doc,cx,y,cw,rh,BLK,SLV,0.3);
        L(doc,cx,y,cx,y+rh,SLV,0.3);
        T(doc,hdrs[h],cx+2*MM,y+3*MM,6.5,WHT,FB,cw-4*MM);
        cx+=cw;
    }
    y+=rh;
    for(r=0;r<rows.length;r++){
        cx=x;
        var bg=r%2===0?LGRY:WHT;
        for(c=0;c<rows[r].length;c++){
            cw=cols[c]*tw;
            R(doc,cx,y,cw,rh,bg,MGRY,0.3);
            T(doc,rows[r][c],cx+2*MM,y+3*MM,6.5,DGRY,FR,cw-4*MM);
            cx+=cw;
        }
        y+=rh;
    }
    return y;
}
function SEC(doc,O,y,titulo){
    T(doc,titulo,O*LW+MRG,y,11,BLK,FB,LW-MRG*2);
    y+=7*MM;
    L(doc,O*LW+MRG,y,O*LW+LW-MRG,y,SLV,0.8);
    return y+5*MM;
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB1 â€” PORTADA
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function portada(doc){
    var O=0;
    // Fondo negro total
    R(doc,0,0,LW,LH,BLK);
    // Franja plateada lateral izq
    R(doc,0,0,6*MM,LH,SLV);
    // Franja plateada lateral der
    R(doc,LW-3*MM,0,3*MM,LH,SLV2);
    // Bloque blanco inferior (zona texto)
    R(doc,6*MM,LH*0.55,LW-6*MM,LH*0.35,WHT);
    // LÃ­nea plateada divisoria
    R(doc,6*MM,LH*0.55,LW-6*MM,1.5*MM,SLV);

    // Logo centrado
    logo(doc,LW/2-30*MM,LH*0.08,60*MM);

    // Texto portada zona negra
    T(doc,"MICSA SAFETY DIVISION",10*MM,LH*0.34,22,WHT,FB,LW-14*MM);
    T(doc,"Seguridad Patrimonial Industrial",10*MM,LH*0.34+21,11,SLV,FS,LW-14*MM);
    // LÃ­nea plateada corta
    R(doc,10*MM,LH*0.34+36,45*MM,1.2*MM,SLV);
    T(doc,'"Disciplina, criterio y control operativo"',10*MM,LH*0.34+43,9.5,SLV2,FL,LW-14*MM);

    // Zona blanca
    T(doc,"PROPUESTA COMERCIAL 2026",10*MM,LH*0.58,9,BLK,FB,LW*0.55);
    T(doc,"Sector Automotriz Â· Industrial Â· LogÃ­stica",10*MM,LH*0.58+13,8,DGRY,FR,LW*0.55);
    T(doc,new Date().toLocaleDateString('es-MX',{}),10*MM,LH*0.58+26,8,DGRY,FL,LW*0.4);
    // LÃ­nea plateada separaciÃ³n derecha
    L(doc,LW*0.6,LH*0.57,LW*0.6,LH*0.55+LH*0.34,MGRY,0.5);
    T(doc,"Director de Seguridad",LW*0.62,LH*0.59,7,DGRY,FL,LW*0.33);
    T(doc,"Gerardo GuzmÃ¡n Alvarado",LW*0.62,LH*0.59+10,8.5,BLK,FB,LW*0.33);
    T(doc,"Fuerza Civil Â· ProtecciÃ³n Civil Â· Bomberos",LW*0.62,LH*0.59+21,7,DGRY,FR,LW*0.33);

    // Footer portada
    R(doc,0,LH-12*MM,LW,12*MM,BLK);
    R(doc,0,LH-12*MM,LW,0.7*MM,SLV);
    T(doc,"MICSA Safety Division  Â·  Seguridad Patrimonial Industrial  Â·  contacto@micsa.mx  Â·  800-MICSA-01",
      6*MM,LH-7*MM,6.5,MGRY,FL,LW-10*MM);
    T(doc,"1 / 6",LW-MRG-12,LH-7*MM,7,SLV,FS,14);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB2 â€” CULTURA OPERATIVA
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function cultura(doc){
    var O=1;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"CULTURA OPERATIVA",2);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"PILARES DE MICSA SAFETY DIVISION");
    T(doc,"Un sistema de seguridad es tan sÃ³lido como la cultura del personal que lo opera. Estos cuatro pilares definen cÃ³mo cada elemento piensa, actÃºa y decide â€” sin excepciÃ³n.",
      O*LW+MRG,y,8,DGRY,FL,LW-MRG*2); y+=13*MM;

    var pw=(LW-MRG*2)*0.48, ph=52*MM;
    var P=[
        {n:"01",t:"RESPETO\nOPERATIVO",
         d:"Cada acciÃ³n del personal refleja la imagen de MICSA Safety ante el cliente. El respeto a las personas, los activos y los procesos es el primer requisito de permanencia. Sin excepciÃ³n ni contexto."},
        {n:"02",t:"DISCIPLINA\nREAL",
         d:"Los protocolos no se interpretan â€” se ejecutan. La disciplina aplica en el turno, el uniforme, la bitÃ¡cora y el reporte. Un guardia que improvisa es un riesgo para el cliente y para el equipo."},
        {n:"03",t:"CRITERIO\nEN CAMPO",
         d:"La seguridad requiere decisiones bajo presiÃ³n. El criterio no se improvisa: se desarrolla con evaluaciÃ³n mensual, simulaciones reales y retroalimentaciÃ³n directa del supervisor de operaciones."},
        {n:"04",t:"CONFIABILIDAD\nDESDE EL ORIGEN",
         d:"Seleccionamos y evaluamos a cada elemento antes de asignarlo. La confiabilidad se construye en el reclutamiento y se sostiene con KPIs medibles, expedientes auditables y reemplazos garantizados."}
    ];
    for(var i=0;i<P.length;i++){
        var px=O*LW+MRG+(i%2)*(pw+MRG*0.5);
        var py=y+Math.floor(i/2)*(ph+4*MM);
        R(doc,px,py,pw,ph,WHT,MGRY,0.5);
        R(doc,px,py,pw,1.8*MM,BLK);
        // Cuadro nÃºmero negro
        R(doc,px+3*MM,py+4.5*MM,12*MM,13*MM,BLK);
        T(doc,P[i].n,px+3.5*MM,py+6*MM,9,SLV,FB,11*MM,"C");
        // LÃ­nea plateada fina
        R(doc,px+18*MM,py+5*MM,1*MM,11*MM,SLV);
        T(doc,P[i].t,px+21*MM,py+5.5*MM,8,BLK,FB,pw-24*MM);
        L(doc,px+21*MM,py+17*MM,px+pw-4*MM,py+17*MM,MGRY,0.4);
        T(doc,P[i].d,px+3*MM,py+19*MM,6.8,DGRY,FR,pw-6*MM);
    }
    y+=2*ph+4*MM+10*MM;

    // Frase cierre
    R(doc,O*LW+MRG,y,LW-MRG*2,16*MM,BLK);
    R(doc,O*LW+MRG,y,LW-MRG*2,1.2*MM,SLV);
    T(doc,'"Puede faltar flujo, pero nunca posiciÃ³n."',O*LW+MRG+5*MM,y+4.5*MM,10.5,WHT,FS,LW-MRG*2-10*MM);
    T(doc,"Gerardo GuzmÃ¡n Alvarado â€” Director de Seguridad, MICSA Safety Division",
      O*LW+MRG+5*MM,y+11.5*MM,6.5,SLV,FL,LW-MRG*2-10*MM);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB3 â€” PLAN DE CAPACITACIÃ“N
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function capacitacion(doc){
    var O=2;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"PLAN DE CAPACITACIÃ“N",3);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"PROGRAMA DE FORMACIÃ“N OPERATIVA â€” 5 FASES");
    T(doc,"Desde el ingreso hasta la certificaciÃ³n operativa. Cada fase tiene criterios de aprobaciÃ³n; sin cumplirlos el elemento no avanza al siguiente nivel.",
      O*LW+MRG,y,8,DGRY,FL,LW-MRG*2); y+=13*MM;

    var fases=[
        {n:"F1",t:"INDUCCIÃ“N",dias:"DÃ­as 1â€“3",
         items:["Historia y valores MICSA","Reglamento interno y cultura","PresentaciÃ³n instalaciones cliente","Cadena de mando y comunicaciÃ³n","Entrega uniforme y EPP completo"]},
        {n:"F2",t:"FORMACIÃ“N\nTÃ‰CNICA",dias:"DÃ­as 4â€“10",
         items:["Control de accesos vehicular/peatonal","Registro bitÃ¡cora digital y QR","Uso correcto de EPP (NOM-017)","Manejo de radio y comunicaciones","Protocolo de rondines certificado"]},
        {n:"F3",t:"ESPECIA-\nLIZACIÃ“N",dias:"DÃ­as 11â€“20",
         items:["NOM-030 â€” prevenciÃ³n y respuesta","OperaciÃ³n CCTV y monitoreo","Manejo de cuatrimoto certificado","Reporte de incidente formal","SimulaciÃ³n de emergencia real"]},
        {n:"F4",t:"EVALUACIÃ“N",dias:"DÃ­as 21â€“25",
         items:["Examen tÃ©cnico â‰¥ 80%","SimulaciÃ³n en campo real","Prueba psicomÃ©trica final","RevisiÃ³n de expediente completo","DecisiÃ³n: aprobado o baja"]},
        {n:"F5",t:"CAPACIT.\nCONTINUA",dias:"Mensual",
         items:["SesiÃ³n mensual obligatoria 60 min","RevisiÃ³n KPIs individuales","ActualizaciÃ³n NOM-030 y cliente","Simulacro semestral obligatorio","RetroalimentaciÃ³n del supervisor"]}
    ];
    var fw=LW-MRG*2, fh=40*MM;
    for(var i=0;i<fases.length;i++){
        var f=fases[i];
        var fx=O*LW+MRG, fy=y+i*(fh+3*MM);
        R(doc,fx,fy,fw,fh,WHT,MGRY,0.4);
        // Columna fase
        R(doc,fx,fy,26*MM,fh,BLK);
        T(doc,f.n,fx+1*MM,fy+5*MM,14,SLV,FB,24*MM,"C");
        T(doc,f.t,fx+1*MM,fy+18*MM,7,WHT,FS,24*MM,"C");
        T(doc,f.dias,fx+1*MM,fy+28*MM,6,SLV2,FL,24*MM,"C");
        L(doc,fx+26*MM,fy+5*MM,fx+26*MM,fy+fh-5*MM,SLV,0.6);
        // Items en 2 columnas
        var iw=(fw-30*MM)/2-2*MM;
        for(var j=0;j<f.items.length;j++){
            var ix=fx+29*MM+Math.floor(j/3)*(iw+2*MM);
            var iy=fy+4*MM+(j%3)*11.5*MM;
            R(doc,ix,iy+2*MM,2*MM,6*MM,SLV);
            T(doc,f.items[j],ix+3*MM,iy+2.5*MM,7,DGRY,FR,iw-1*MM);
        }
        // Conector
        if(i<fases.length-1){
            R(doc,fx+fw/2-3*MM,fy+fh,6*MM,3*MM,SLV2);
        }
    }
    y+=fases.length*(fh+3*MM)+8*MM;

    // Criterio de aprobaciÃ³n
    R(doc,O*LW+MRG,y,LW-MRG*2,14*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,2*MM,14*MM,BLK);
    T(doc,"CRITERIO DE APROBACIÃ“N: Puntaje â‰¥ 80% en matriz de evaluaciÃ³n. Elemento reprobado = baja inmediata sin recontrataciÃ³n en 90 dÃ­as.",
      O*LW+MRG+5*MM,y+4.5*MM,7.5,DGRY,FR,LW-MRG*2-8*MM);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB4 â€” VALOR AGREGADO
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function valor(doc){
    var O=3;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"VALOR AGREGADO",4);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"POR QUÃ‰ MICSA SAFETY â€” NO COMPETIMOS POR PRECIO");
    T(doc,"No somos una empresa de guardias. Somos un sistema de control operativo que se integra a la operaciÃ³n del cliente, reduce pÃ©rdidas y genera evidencia auditable de cada turno.",
      O*LW+MRG,y,8,DGRY,FL,LW-MRG*2); y+=14*MM;

    var V=[
        {n:"01",t:"PERSONAL CON CRITERIO",
         d:"Filtro psicomÃ©trico â‰¥80% + evaluaciÃ³n tÃ©cnica + verificaciÃ³n de antecedentes. No contratamos cuerpos â€” contratamos criterio. El guardia que no aprueba no ingresa y no hay excepciones.",
         m:"Filtro 5 fases Â· 0 excepciones"},
        {n:"02",t:"CONTROL OPERATIVO REAL",
         d:"BitÃ¡cora digital con timestamp, rondines con QR en puntos de control, monitoreo CCTV dedicado. El cliente tiene visibilidad total del 100% de la operaciÃ³n en tiempo real. Cero zonas grises.",
         m:"100% trazabilidad Â· dashboard mensual"},
        {n:"03",t:"INTEGRACIÃ“N INDUSTRIAL",
         d:"Nuestros supervisores conocen los ritmos de producciÃ³n automotriz. Coordinamos con superintendentes de planta. Protocolo de arranque en 15 dÃ­as hÃ¡biles sin interrumpir la operaciÃ³n del cliente.",
         m:"Arranque operativo en 15 dÃ­as hÃ¡biles"},
        {n:"04",t:"SUPERVISIÃ“N ACTIVA",
         d:"Jefe de Seguridad + Supervisor en sitio durante toda la operaciÃ³n. No hay autogestiÃ³n ni supervisiÃ³n remota. Alguien responsable presente en cada turno. Respuesta garantizada < 3 minutos.",
         m:"Respuesta < 3 min Â· supervisor en sitio"},
        {n:"05",t:"TRAZABILIDAD DOCUMENTAL",
         d:"Cada elemento tiene expediente completo auditables. Cada incidente tiene reporte escrito en â‰¤2 horas. Dashboard mensual de KPIs entregado al cliente. Disponible para auditorÃ­a STPS en tiempo real.",
         m:"Reporte escrito â‰¤ 2h Â· auditable STPS"}
    ];
    var vw=LW-MRG*2, vh=36*MM;
    for(var i=0;i<V.length;i++){
        var vx=O*LW+MRG, vy=y+i*(vh+3*MM);
        R(doc,vx,vy,vw,vh,WHT,MGRY,0.4);
        // Barra izq negra con nÃºmero
        R(doc,vx,vy,20*MM,vh,BLK);
        T(doc,V[i].n,vx+1*MM,vy+6*MM,12,SLV,FB,18*MM,"C");
        // LÃ­nea plateada
        L(doc,vx+20*MM,vy+6*MM,vx+20*MM,vy+vh-6*MM,SLV,0.6);
        // TÃ­tulo
        T(doc,V[i].t,vx+23*MM,vy+5*MM,8.5,BLK,FB,vw-26*MM);
        // Texto
        T(doc,V[i].d,vx+23*MM,vy+14.5*MM,7,DGRY,FR,vw-26*MM);
        // Badge mÃ©trica
        R(doc,vx+vw-60*MM,vy+vh-8*MM,57*MM,6.5*MM,LGRY,MGRY,0.3);
        R(doc,vx+vw-60*MM,vy+vh-8*MM,2*MM,6.5*MM,SLV);
        T(doc,V[i].m,vx+vw-57*MM,vy+vh-5*MM,6.5,DGRY,FS,54*MM);
    }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB5 â€” SISTEMA OPERATIVO / MANUALES
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function manuales(doc){
    var O=4;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"SISTEMA OPERATIVO â€” MANUALES",5);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"ARQUITECTURA DE DOCUMENTACIÃ“N OPERATIVA");
    T(doc,"Seis manuales alineados a NOM-STPS y al reglamento del cliente. Disponibles en formato digital e impreso. Actualizados semestralmente.",
      O*LW+MRG,y,8,DGRY,FL,LW-MRG*2); y+=12*MM;

    var M=[
        {cod:"M-01",t:"MANUAL GENERAL DE OPERACIONES",
         items:["Estructura organizacional y cadena de mando completa","DescripciÃ³n detallada de puestos por nivel","PolÃ­ticas de contrataciÃ³n, evaluaciÃ³n y baja","ComunicaciÃ³n interna y reportes formales","Protocolo de sustituciÃ³n de personal garantizado"]},
        {cod:"M-02",t:"MANUAL INTRAMUROS â€” PLANTA",
         items:["Control de acceso vehicular y peatonal con registro","GestiÃ³n de visitas, contratistas y proveedores","Protocolos cierre/apertura de instalaciones","Manejo de activos crÃ­ticos y materiales","CoordinaciÃ³n directa con seguridad del cliente"]},
        {cod:"M-03",t:"MANUAL DE MONITOREO CCTV",
         items:["OperaciÃ³n del sistema de videovigilancia 24/7","Protocolo de detecciÃ³n y clasificaciÃ³n de amenazas","Reporte con evidencia de incidentes en video","Mantenimiento bÃ¡sico de equipos CCTV","BitÃ¡cora de monitoreo por turno con timestamp"]},
        {cod:"M-04",t:"PROTOCOLO DE USO DE FUERZA",
         items:["Principios de proporcionalidad, necesidad y legalidad","Niveles de respuesta escalada a amenaza fÃ­sica","CoordinaciÃ³n con autoridades (policÃ­a / FGR)","DocumentaciÃ³n obligatoria post-incidente","Marco legal aplicable â€” derechos y responsabilidades"]},
        {cod:"M-05",t:"MANUAL DE RECLUTAMIENTO",
         items:["Perfil mÃ­nimo por puesto â€” requisitos verificables","Proceso 5 fases con criterio de aprobaciÃ³n â‰¥80%","DocumentaciÃ³n obligatoria de ingreso completa","PerÃ­odo de prueba 30 dÃ­as con evaluaciÃ³n semanal","Causales de baja y procedimiento formal documentado"]},
        {cod:"M-06",t:"REGLAMENTO INTERNO",
         items:["CÃ³digo de vestimenta y uniforme obligatorio","Uso de dispositivos mÃ³viles â€” prohibido en turno activo","BitÃ¡cora: sin firma = turno incompleto + sanciÃ³n","Cero tolerancia: alcohol, sustancias, propinas, regalos","Sanciones progresivas y causales de baja inmediata"]}
    ];
    var bw=(LW-MRG*2)*0.48, bh=55*MM;
    for(var i=0;i<M.length;i++){
        var bx=O*LW+MRG+(i%2)*(bw+MRG*0.5);
        var by=y+Math.floor(i/2)*(bh+4*MM);
        R(doc,bx,by,bw,bh,WHT,MGRY,0.4);
        R(doc,bx,by,bw,9.5*MM,BLK);
        R(doc,bx,by,2*MM,bh,SLV);
        T(doc,M[i].cod,bx+4*MM,by+2.5*MM,7,SLV,FS,20*MM);
        T(doc,M[i].t,bx+26*MM,by+2.5*MM,6.8,WHT,FB,bw-28*MM);
        for(var j=0;j<M[i].items.length;j++){
            T(doc,"â–¸  "+M[i].items[j],bx+4*MM,by+11.5*MM+j*7.5*MM,6.5,DGRY,FR,bw-8*MM);
        }
    }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB6 â€” PLAN DE IMPLEMENTACIÃ“N
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function implementacion(doc){
    var O=5;
    R(doc,O*LW,0,LW,LH,WHT);
    HDR(doc,O,"PLAN DE IMPLEMENTACIÃ“N",6);
    FTR(doc,O);
    var y=19*MM;
    y=SEC(doc,O,y,"RUTA CRÃTICA â€” ARRANQUE EN 15 DÃAS HÃBILES");
    T(doc,"Desde la firma del contrato hasta operaciÃ³n plena con KPIs activos. Cinco fases secuenciales con hitos verificables en cada etapa.",
      O*LW+MRG,y,8,DGRY,FL,LW-MRG*2); y+=12*MM;

    var F=[
        {n:"F1",t:"ESTRUCTURA",  d:"DÃ­as 1â€“3",
         items:["Alta sistema proveedores cliente","Firma contrato + PO emitida","IntegraciÃ³n REPSE y fiscal","AsignaciÃ³n Jefe de Seguridad","Carta inicio formal entregada"]},
        {n:"F2",t:"LEGAL",       d:"DÃ­as 4â€“6",
         items:["RevisiÃ³n reglamento cliente","NDA y confidencialidad firmados","Alta IMSS personal asignado","Registro STPS obligaciones","Constancias SAT ambas partes"]},
        {n:"F3",t:"OPERACIÃ“N",   d:"DÃ­as 7â€“10",
         items:["InducciÃ³n 3 dÃ­as personal","Entrega uniformes + EPP completo","Reconocimiento fÃ­sico planta","Acuerdo puntos control QR","ActivaciÃ³n bitÃ¡cora digital"]},
        {n:"F4",t:"IMPLEMENTACIÃ“N",d:"DÃ­as 11â€“13",
         items:["OperaciÃ³n supervisada en campo","ValidaciÃ³n protocolo con cliente","Ajuste turnos definitivos","ActivaciÃ³n monitoreo CCTV","Primera bitÃ¡cora semanal"]},
        {n:"F5",t:"CONTROL",     d:"DÃ­as 14â€“15",
         items:["Dashboard KPI primer perÃ­odo","ReuniÃ³n alineaciÃ³n cliente","Plan mejora continua acordado","Fecha revisiÃ³n mensual fija","OperaciÃ³n plena transferida"]}
    ];
    var fw=LW-MRG*2, fh=42*MM;
    for(var i=0;i<F.length;i++){
        var ff=F[i];
        var fx=O*LW+MRG, fy=y+i*(fh+3*MM);
        R(doc,fx,fy,fw,fh,WHT,MGRY,0.4);
        // NÃºmero
        R(doc,fx,fy,25*MM,fh,BLK);
        T(doc,ff.n,fx+1*MM,fy+5*MM,16,SLV,FB,23*MM,"C");
        T(doc,ff.t,fx+1*MM,fy+20*MM,7,WHT,FS,23*MM,"C");
        T(doc,ff.d,fx+1*MM,fy+28*MM,6.5,SLV2,FL,23*MM,"C");
        L(doc,fx+25*MM,fy+5*MM,fx+25*MM,fy+fh-5*MM,SLV,0.6);
        // Items 2 col
        var iw=(fw-29*MM)/2-2*MM;
        for(var j=0;j<ff.items.length;j++){
            var ix=fx+27*MM+Math.floor(j/3)*(iw+2*MM);
            var iy=fy+4*MM+(j%3)*12.5*MM;
            R(doc,ix,iy+2.5*MM,2*MM,7*MM,SLV);
            T(doc,ff.items[j],ix+4*MM,iy+3*MM,7,DGRY,FR,iw-2*MM);
        }
        // Conector a siguiente
        if(i<F.length-1){
            L(doc,fx+fw/2,fy+fh,fx+fw/2,fy+fh+3*MM,MGRY,1);
        }
    }
    y+=F.length*(fh+3*MM)+8*MM;

    // Cierre â€” bloque negro premium
    R(doc,O*LW+MRG,y,LW-MRG*2,26*MM,BLK);
    R(doc,O*LW+MRG,y,LW-MRG*2,1.5*MM,SLV);
    T(doc,"MICSA SAFETY DIVISION â€” SISTEMA INTEGRAL DE SEGURIDAD PATRIMONIAL",
      O*LW+MRG,y+5*MM,9,WHT,FB,LW-MRG*2,"C");
    R(doc,O*LW+LW/2-22*MM,y+14*MM,44*MM,0.7*MM,SLV);
    T(doc,'"Disciplina, criterio y control operativo"',O*LW+MRG,y+17*MM,9,SLV,FL,LW-MRG*2,"C");

    // Bloque firmas
    y+=32*MM;
    var sw=(LW-MRG*2)*0.45;
    R(doc,O*LW+MRG,y,sw,20*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+MRG,y,sw,1.5*MM,BLK);
    T(doc,"_________________________",O*LW+MRG+5*MM,y+10*MM,8,DGRY,FR,sw-10*MM);
    T(doc,"Director General â€” MICSA Safety Division",O*LW+MRG+5*MM,y+15*MM,6.5,DGRY,FL,sw-10*MM);

    R(doc,O*LW+LW-MRG-sw,y,sw,20*MM,LGRY,MGRY,0.4);
    R(doc,O*LW+LW-MRG-sw,y,sw,1.5*MM,BLK);
    T(doc,"_________________________",O*LW+LW-MRG-sw+5*MM,y+10*MM,8,DGRY,FR,sw-10*MM);
    T(doc,"Representante Legal â€” Cliente",O*LW+LW-MRG-sw+5*MM,y+15*MM,6.5,DGRY,FL,sw-10*MM);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// MAIN
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function main(){
    var doc=app.documents.add(DocumentColorSpace.CMYK, LW*6, LH);

    for(var i=1;i<6;i++) doc.artboards.add([i*LW,0,(i+1)*LW,-LH]);
    var N=["01_PORTADA","02_CULTURA","03_CAPACITACION","04_VALOR_AGREGADO","05_MANUALES","06_IMPLEMENTACION"];
    for(var i=0;i<6;i++) doc.artboards[i].name=N[i];

    portada(doc); cultura(doc); capacitacion(doc);
    valor(doc);   manuales(doc); implementacion(doc);

    var ai=new File(BASE+"MICSA_SAFETY_TACTICO.ai");
    var ao=new IllustratorSaveOptions();
    ao.compatibility=Compatibility.ILLUSTRATOR24;
    ao.saveMultipleArtboards=true;
    doc.saveAs(ai,ao);

    try{
        var pdf=new File(OUT+"MICSA_SAFETY_TACTICO.pdf");
        var po=new PDFSaveOptions();
        po.compatibility=PDFCompatibility.ACROBAT8;
        po.generateThumbnails=true;
        po.preserveEditability=false;
        po.saveMultipleArtboards=true;
        po.artboardRange="1-6";
        doc.saveAs(pdf,po);
    }catch(e){}

    alert("MICSA_SAFETY_TACTICO.ai â€” 6 artboards.\n\nPortada Â· Cultura Â· CapacitaciÃ³n Â· Valor Â· Manuales Â· ImplementaciÃ³n\n\nPaleta: Negro Rico + Plateado K25\nPDF exportado en PDFs/");
}
main();



