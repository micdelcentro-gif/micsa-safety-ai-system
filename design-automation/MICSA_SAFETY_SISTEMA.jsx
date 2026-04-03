/**
 * MICSA SAFETY DIVISION â€” Sistema Corporativo Premium
 * ExtendScript para Adobe Illustrator
 * 6 Artboards A4 â€” CMYK â€” Paleta Rojo + Blanco + Gris
 * USO: Archivo > Scripts > Otro script...
 */

var BASE_PATH = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\micsa-safety-ai-system\\";
var LOGO_FILE = BASE_PATH + "assets\\logo.png";
var OUT_PATH  = BASE_PATH + "output\\";

var MM  = 2.834645669;
var A4W = 210 * MM;
var A4H = 297 * MM;
var MRG = 20  * MM;

function cmyk(c,m,y,k){ var col=new CMYKColor(); col.cyan=c; col.magenta=m; col.yellow=y; col.black=k; return col; }
var RED    = cmyk(0,90,85,0);     // Rojo MICSA dominante
var RED2   = cmyk(0,65,60,0);     // Rojo claro / acento
var BLACK  = cmyk(60,40,40,100);  // Negro rico
var DGRAY  = cmyk(0,0,0,75);      // Gris oscuro
var MGRAY  = cmyk(0,0,0,30);      // Gris medio
var LGRAY  = cmyk(0,0,0,8);       // Gris muy claro
var WHITE  = cmyk(0,0,0,0);

var FB = "Montserrat-Bold";
var FS = "Montserrat-SemiBold";
var FR = "OpenSans-Regular";
var FL = "OpenSans-Light";

// â”€â”€â”€ HELPERS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function rect(doc,x,y,w,h,fill,stroke,sw){
    var r=doc.pathItems.rectangle(-y,x,w,h);
    r.filled=true; r.fillColor=fill;
    if(stroke){r.stroked=true; r.strokeColor=stroke; r.strokeWidth=sw||0.5;}
    else r.stroked=false; return r;
}
function txt(doc,s,x,y,sz,col,font,w,align){
    var t=doc.textFrames.add();
    t.contents=s; t.left=x; t.top=-y;
    if(w) t.width=w;
    t.textRange.characterAttributes.size=sz;
    t.textRange.characterAttributes.fillColor=col;
    try{t.textRange.characterAttributes.textFont=app.textFonts.getByName(font);}catch(e){}
    if(align==="C") t.textRange.paragraphAttributes.justification=Justification.CENTER;
    else if(align==="R") t.textRange.paragraphAttributes.justification=Justification.RIGHT;
    return t;
}
function line(doc,x1,y1,x2,y2,col,w){
    var l=doc.pathItems.add();
    l.setEntirePath([[x1,-y1],[x2,-y2]]);
    l.stroked=true; l.strokeColor=col; l.strokeWidth=w||0.5; l.filled=false; return l;
}
function logo(doc,x,y,w){
    try{
        var f=new File(LOGO_FILE); if(!f.exists) return;
        var p=doc.placedItems.add(); p.file=f;
        var sc=w/p.width; p.width=w; p.height=p.height*sc;
        p.left=x; p.top=-y;
    }catch(e){}
}
function header(doc,O,titulo,pgNum){
    // Full Page Background Reset
    rect(doc,O*A4W,0,A4W,A4H,WHITE);
    
    // Tactical Sidebar
    rect(doc,O*A4W,0,4*MM,A4H,BLACK);
    rect(doc,O*A4W+4*MM,0,0.8*MM,A4H,RED);

    // Header HUD
    rect(doc,O*A4W+4.8*MM,0,A4W-4.8*MM,15*MM,BLACK);
    line(doc,O*A4W+4.8*MM,15*MM,O*A4W+A4W,15*MM,RED,0.5);
    
    logo(doc,O*A4W+8*MM,2*MM,14*MM);
    txt(doc,"MICSA SAFETY DIVISION",O*A4W+25*MM,4.5*MM,8,WHITE,FB,A4W*0.5);
    txt(doc,titulo.toUpperCase(),O*A4W+25*MM,10.5*MM,6.5,LGRAY,FL,A4W*0.55);
    
    // Page Signal
    rect(doc,(O+1)*A4W-15*MM,2*MM,10*MM,11*MM,RED);
    txt(doc,String(pgNum),(O+1)*A4W-10*MM,5.5*MM,10,WHITE,FB,8*MM,"C");
}
function footer(doc,O){
    rect(doc,O*A4W+4.8*MM,A4H-10*MM,A4W-4.8*MM,10*MM,BLACK);
    rect(doc,O*A4W+4.8*MM,A4H-10*MM,A4W-4.8*MM,0.5*MM,RED);
    txt(doc,"CONFIDENCIAL  |  MICSA Safety Division  |  Monclova, Coahuila  |  micsasafety.com.mx",
        O*A4W+10*MM,A4H-6*MM,5.5,MGRAY,FL,A4W-MRG*2);
}
function sec(doc,O,y,titulo){
    rect(doc,O*A4W+MRG,y,A4W-MRG*2,8*MM,LGRAY,RED,0.3);
    rect(doc,O*A4W+MRG,y,2.5*MM,8*MM,BLACK);
    txt(doc,titulo.toUpperCase(),O*A4W+MRG+6*MM,y+1.5*MM,9,BLACK,FB,A4W-MRG*2-8*MM);
    y+=10*MM;
    line(doc,O*A4W+MRG,y,O*A4W+A4W-MRG,y,MGRAY,0.5);
    return y+6*MM;
}

// Tabla
function tabla(doc,O,y,hdrs,rows,cols){
    var x=O*A4W+MRG, rh=10*MM, tw=A4W-MRG*2;
    var cx=x;
    for(var h=0;h<hdrs.length;h++){
        var cw=cols[h]*tw;
        rect(doc,cx,y,cw,rh,RED);
        line(doc,cx,y,cx,y+rh,RED2,0.3);
        txt(doc,hdrs[h],cx+2*MM,y+3*MM,6.5,WHITE,FB,cw-4*MM);
        cx+=cw;
    }
    y+=rh;
    for(var r=0;r<rows.length;r++){
        cx=x;
        var bg=r%2===0?LGRAY:WHITE;
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

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB1 â€” PORTADA
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function portada(doc){
    var O=0;
    // Fondo dividido: rojo superior 60%, blanco inferior
    rect(doc,0,0,A4W,A4H*0.60,RED);
    rect(doc,0,A4H*0.60,A4W,A4H*0.40,WHITE);
    // Franja negra vertical izq decorativa
    rect(doc,0,0,7*MM,A4H,BLACK);
    // LÃ­nea blanca divisoria
    rect(doc,0,A4H*0.60,A4W,1.2*MM,BLACK);
    // Logo
    logo(doc,A4W/2-25*MM,A4H*0.08,50*MM);
    // TÃ­tulo
    txt(doc,"MICSA SAFETY DIVISION",10*MM,A4H*0.30,26,WHITE,FB,A4W-14*MM);
    txt(doc,"Sistema Integral de Seguridad Patrimonial e Industrial",10*MM,A4H*0.30+22,11,WHITE,FL,A4W-14*MM);
    // LÃ­nea roja corta â†’ blanca al cambiar fondo
    rect(doc,10*MM,A4H*0.30+38,50*MM,1.5*MM,WHITE);
    txt(doc,'"No cuidamos instalaciones. Controlamos operaciones."',10*MM,A4H*0.30+46,10.5,WHITE,FS,A4W-14*MM);
    // Bloque info inferior
    rect(doc,10*MM,A4H*0.63,A4W-20*MM,A4H*0.25,LGRAY);
    rect(doc,10*MM,A4H*0.63,2*MM,A4H*0.25,RED);
    txt(doc,"PROPUESTA COMERCIAL",14*MM,A4H*0.65,9,BLACK,FB,A4W*0.5);
    txt(doc,"Sector Automotriz e Industrial â€” 2026",14*MM,A4H*0.65+13,8,DGRAY,FR,A4W*0.5);
    txt(doc,new Date().toLocaleDateString('es-MX',{}),14*MM,A4H*0.65+26,8,DGRAY,FL,A4W*0.4);
    // Footer portada
    rect(doc,0,A4H-13*MM,A4W,13*MM,BLACK);
    rect(doc,0,A4H-13*MM,A4W,0.8*MM,RED);
    txt(doc,"MICSA Safety Division  Â·  Seguridad Patrimonial Industrial  Â·  contacto@micsa.mx",10*MM,A4H-8*MM,6.5,MGRAY,FL,A4W-20*MM);
    txt(doc,"1 / 6",A4W-MRG-16,A4H-8*MM,7,RED,FS,16);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB2 â€” FILOSOFÃA Y CULTURA OPERATIVA
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function filosofia(doc){
    var O=1;
    header(doc,O,"FILOSOFÃ A Y CULTURA OPERATIVA",2);
    footer(doc,O);

    var y=22*MM;
    y=sec(doc,O,y,"MISIÃ“N CORPORATIVA");

    rect(doc,O*A4W,0,A4W,A4H,WHITE);
    header(doc,O,"FILOSOFÃA Y CULTURA OPERATIVA",2);
    footer(doc,O);

    var y=20*MM;
    txt(doc,"MISIÃ“N",O*A4W+MRG,y,12,RED,FB,A4W-MRG*2); y+=7*MM;
    line(doc,O*A4W+MRG,y,O*A4W+A4W-MRG,y,RED,1); y+=5*MM;
    rect(doc,O*A4W+MRG,y,A4W-MRG*2,18*MM,LGRAY);
    rect(doc,O*A4W+MRG,y,2*MM,18*MM,RED);
    txt(doc,"Proteger los activos fÃ­sicos, humanos e intangibles de nuestros clientes mediante un sistema integral de seguridad basado en anÃ¡lisis de riesgos, personal certificado y procesos auditables en tiempo real.",
        O*A4W+MRG+5*MM,y+4*MM,8,DGRAY,FR,A4W-MRG*2-8*MM);
    y+=22*MM;

    y=sec(doc,O,y,"LOS 4 PILARES DE MICSA SAFETY");

    var pilares=[
        {n:"01",t:"RESPETO OPERATIVO",
         d:"Cada interacciÃ³n refleja la imagen de MICSA Safety. El respeto con el personal del cliente, visitantes y compaÃ±eros es el requisito mÃ­nimo para pertenecer al equipo. No es negociable."},
        {n:"02",t:"DISCIPLINA REAL",
         d:"Los protocolos existen para cumplirse. No hay interpretaciones ni excepciones. La disciplina aplica en turno, uniforme, bitÃ¡cora y reporte â€” sin importar el contexto ni la hora."},
        {n:"03",t:"CRITERIO EN CAMPO",
         d:"Un guardia entrenado toma decisiones correctas bajo presiÃ³n. El criterio se desarrolla con evaluaciÃ³n mensual, simulaciones y retroalimentaciÃ³n directa del supervisor de operaciones."},
        {n:"04",t:"CONFIABILIDAD DESDE EL ORIGEN",
         d:"Seleccionamos, evaluamos y supervisamos a cada elemento. La confiabilidad comienza en el reclutamiento y se sostiene con KPIs medibles, expediente auditables y reemplazos garantizados."}
    ];
    var pw=(A4W-MRG*2)*0.48;
    var ph=50*MM;
    for(var i=0;i<pilares.length;i++){
        var px=O*A4W+MRG+(i%2)*(pw+MRG*0.5);
        var py=y+Math.floor(i/2)*(ph+4*MM);
        rect(doc,px,py,pw,ph,WHITE,MGRAY,0.5);
        rect(doc,px,py,pw,2*MM,RED);
        // NÃºmero en cuadro rojo
        rect(doc,px+3*MM,py+5*MM,12*MM,12*MM,RED);
        txt(doc,pilares[i].n,px+3.5*MM,py+6*MM,9,WHITE,FB,11*MM,"C");
        txt(doc,pilares[i].t,px+18*MM,py+6*MM,8,BLACK,FB,pw-22*MM);
        line(doc,px+18*MM,py+15.5*MM,px+pw-4*MM,py+15.5*MM,MGRAY,0.4);
        txt(doc,pilares[i].d,px+3*MM,py+18*MM,6.8,DGRAY,FR,pw-6*MM);
    }
    y+=2*ph+4*MM+8*MM;

    // Frase cierre
    rect(doc,O*A4W+MRG,y,A4W-MRG*2,16*MM,BLACK);
    rect(doc,O*A4W+MRG,y,A4W-MRG*2,1*MM,RED);
    txt(doc,'"Puede faltar flujo, pero nunca posiciÃ³n."',O*A4W+MRG+5*MM,y+5*MM,10,WHITE,FS,A4W-MRG*2-10*MM);
    txt(doc,"Principio operativo MICSA Safety Division â€” Director: Gerardo GuzmÃ¡n Alvarado",
        O*A4W+MRG+5*MM,y+11.5*MM,6.5,RED2,FL,A4W-MRG*2-10*MM);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB3 â€” MODELO DE OPERACIÃ“N 3 CAPAS
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function modelo(doc){
    var O=2;
    rect(doc,O*A4W,0,A4W,A4H,WHITE);
    header(doc,O,"MODELO DE OPERACIÃ“N â€” SISTEMA DE 3 CAPAS",3);
    footer(doc,O);

    var y=22*MM;
    y=sec(doc,O,y,"SISTEMA INTEGRAL DE CONTROL DE RIESGOS");
    txt(doc,"Tres capas independientes y complementarias que actÃºan en secuencia para neutralizar cualquier amenaza antes de que se convierta en pÃ©rdida.",
        O*A4W+MRG,y,8,DGRAY,FL,A4W-MRG*2); y+=14*MM;

    var capas=[
        {n:"CAPA 1",t:"PREVENCIÃ“N",col:RED,
         items:[
             "AnÃ¡lisis de riesgos inicial y periÃ³dico","Control de accesos vehicular y peatonal",
             "Filtros documentales de visitantes y contratistas","IdentificaciÃ³n de vulnerabilidades fÃ­sicas",
             "Protocolo de horarios crÃ­ticos de operaciÃ³n","Reporte preventivo semanal al cliente"
         ],
         leyenda:"Primera lÃ­nea de defensa â€” ActÃºa antes de que ocurra el incidente"},
        {n:"CAPA 2",t:"DETECCIÃ“N",col:BLACK,
         items:[
             "Monitoreo CCTV con operador dedicado 24/7","Rondines con registro QR en puntos de control",
             "Sistema de bitÃ¡cora digital en tiempo real","DetecciÃ³n temprana de comportamientos irregulares",
             "ComunicaciÃ³n directa supervisorâ€“monitorista","Alertas inmediatas a cadena de mando"
         ],
         leyenda:"Vigilancia activa â€” Detecta antes de que el riesgo escale"},
        {n:"CAPA 3",t:"RESPUESTA",col:DGRAY,
         items:[
             "IntervenciÃ³n inmediata (tiempo mÃ¡x. 3 min)","ContenciÃ³n y aislamiento del Ã¡rea afectada",
             "NotificaciÃ³n a Jefe de Seguridad y cliente","CoordinaciÃ³n con autoridades si aplica",
             "Reporte escrito de incidente en mÃ¡x. 2 horas","AnÃ¡lisis post-incidente y mejora de protocolo"
         ],
         leyenda:"Respuesta controlada â€” Minimiza el daÃ±o y documenta todo"}
    ];

    var cw=A4W-MRG*2;
    var ch=62*MM;
    for(var i=0;i<capas.length;i++){
        var ca=capas[i];
        var cx=O*A4W+MRG;
        var cy=y+i*(ch+4*MM);
        // Contenedor
        rect(doc,cx,cy,cw,ch,WHITE,MGRAY,0.5);
        // Columna izq coloreada
        rect(doc,cx,cy,28*MM,ch,ca.col);
        txt(doc,ca.n,cx+2*MM,cy+8*MM,8,WHITE,FB,24*MM,"C");
        txt(doc,ca.t,cx+2*MM,cy+17*MM,11,WHITE,FB,24*MM,"C");
        // Flecha indicadora
        if(i<2){
            txt(doc,"â–¼",cx+cw/2-3*MM,cy+ch+1*MM,8,RED,FB,8*MM,"C");
        }
        // Items en dos columnas
        var iw=(cw-32*MM)/2-3*MM;
        for(var j=0;j<ca.items.length;j++){
            var ix=cx+30*MM+(Math.floor(j/3))*(iw+3*MM);
            var iy=cy+5*MM+(j%3)*12*MM;
            rect(doc,ix,iy,2.5*MM,8*MM,ca.col);
            txt(doc,ca.items[j],ix+4*MM,iy+1.5*MM,7,DGRAY,FR,iw-2*MM);
        }
        // Leyenda inferior
        rect(doc,cx+28*MM,cy+ch-9*MM,cw-28*MM,9*MM,LGRAY);
        txt(doc,"â†’ "+ca.leyenda,cx+31*MM,cy+ch-5.5*MM,6.5,DGRAY,FL,cw-32*MM);
    }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB4 â€” LIDERAZGO Y ESTRUCTURA DE MANDO
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function liderazgo(doc){
    var O=3;
    rect(doc,O*A4W,0,A4W,A4H,WHITE);
    header(doc,O,"LIDERAZGO Y ESTRUCTURA DE MANDO",4);
    footer(doc,O);

    var y=22*MM;
    y=sec(doc,O,y,"CADENA DE MANDO OPERATIVO");

    // Organigrama vertical
    var bw=70*MM, bh=11*MM;
    var cx=O*A4W+A4W/2-bw/2;
    var niveles=[
        {t:"DIRECCIÃ“N GENERAL",st:"Grupo MICSA",col:BLACK,tc:WHITE},
        {t:"DIRECCIÃ“N DE SEGURIDAD",st:"Gerardo GuzmÃ¡n Alvarado",col:RED,tc:WHITE},
        {t:"JEFE DE OPERACIONES",st:"CoordinaciÃ³n tÃ¡ctica en campo",col:DGRAY,tc:WHITE},
        {t:"SUPERVISORES DE TURNO",st:"Turno A  Â·  Turno B  Â·  SupervisiÃ³n FCA",col:MGRAY,tc:BLACK},
        {t:"OFICIALES / MONITORISTAS",st:"Personal operativo certificado",col:LGRAY,tc:BLACK}
    ];
    for(var i=0;i<niveles.length;i++){
        var ny=y+i*(bh+8*MM);
        rect(doc,cx,ny,bw,bh,niveles[i].col,MGRAY,0.4);
        rect(doc,cx,ny,2.5*MM,bh,RED);
        txt(doc,niveles[i].t,cx+5*MM,ny+3*MM,7,niveles[i].tc,FB,bw-8*MM);
        txt(doc,niveles[i].st,cx+5*MM,ny+8*MM,5.5,i<3?RED2:MGRAY,FL,bw-8*MM);
        // LÃ­nea conector
        if(i<niveles.length-1){
            line(doc,cx+bw/2,ny+bh,cx+bw/2,ny+bh+8*MM,MGRAY,0.8);
        }
    }
    y += niveles.length*(bh+8*MM)+5*MM;

    // Perfil director
    y=sec(doc,O,y,"DIRECTOR DE SEGURIDAD â€” PERFIL EJECUTIVO");

    var ph=65*MM, pw=A4W-MRG*2;
    rect(doc,O*A4W+MRG,y,pw,ph,LGRAY,MGRAY,0.5);
    rect(doc,O*A4W+MRG,y,2*MM,ph,RED);
    // Foto placeholder
    var fpw=28*MM, fph=35*MM;
    rect(doc,O*A4W+MRG+4*MM,y+5*MM,fpw,fph,WHITE,MGRAY,0.5);
    txt(doc,"FOTO",O*A4W+MRG+4*MM,y+19*MM,7,MGRAY,FL,fpw,"C");
    txt(doc,"Corporativa",O*A4W+MRG+4*MM,y+26*MM,6,MGRAY,FL,fpw,"C");
    // Datos
    var dx=O*A4W+MRG+36*MM;
    txt(doc,"Gerardo GuzmÃ¡n Alvarado",dx,y+7*MM,12,BLACK,FB,pw-40*MM);
    rect(doc,dx,y+18*MM,40*MM,1*MM,RED);
    txt(doc,"Director de Seguridad â€” MICSA Safety Division",dx,y+21*MM,8,DGRAY,FS,pw-40*MM);
    var exp=[
        "â–¸  5 aÃ±os en Fuerza Civil â€” Nuevo LeÃ³n",
        "â–¸  Servicio en ProtecciÃ³n Civil y Bomberos",
        "â–¸  GestiÃ³n de crisis y situaciones de riesgo real",
        "â–¸  CertificaciÃ³n en control de operaciones de alto riesgo",
        "â–¸  FormaciÃ³n en manejo de incidentes NFPA / STPS"
    ];
    for(var e=0;e<exp.length;e++){
        txt(doc,exp[e],dx,y+30*MM+e*6.5*MM,7,DGRAY,FR,pw-40*MM);
    }
    y+=ph+10*MM;

    // KPIs de liderazgo
    rect(doc,O*A4W+MRG,y,A4W-MRG*2,20*MM,RED);
    rect(doc,O*A4W+MRG,y,A4W-MRG*2,1.2*MM,BLACK);
    var kw=(A4W-MRG*2)/4;
    var kpis=[
        {v:"100%",l:"Cobertura\n24/7"},
        {v:"<3min",l:"Tiempo de\nrespuesta"},
        {v:"â‰¤5%",l:"RotaciÃ³n\nmensual"},
        {v:"0",l:"Incidentes\nno reportados"}
    ];
    for(var k=0;k<kpis.length;k++){
        var kx=O*A4W+MRG+k*kw;
        if(k>0) line(doc,kx,y+3*MM,kx,y+17*MM,RED2,0.5);
        txt(doc,kpis[k].v,kx+2*MM,y+5*MM,13,WHITE,FB,kw-4*MM,"C");
        txt(doc,kpis[k].l,kx+2*MM,y+14*MM,5.5,RED2,FL,kw-4*MM,"C");
    }
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB5 â€” SERVICIOS Y VALOR AGREGADO
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function servicios(doc){
    var O=4;
    rect(doc,O*A4W,0,A4W,A4H,WHITE);
    header(doc,O,"SERVICIOS Y PROPUESTA DE VALOR",5);
    footer(doc,O);

    var y=22*MM;
    y=sec(doc,O,y,"POR QUÃ‰ MICSA SAFETY â€” CONTROL REAL");
    txt(doc,"No vendemos guardias. Entregamos un sistema de control operativo que se integra a su operaciÃ³n industrial y genera evidencia auditble de cada turno.",
        O*A4W+MRG,y,8,DGRAY,FL,A4W-MRG*2); y+=14*MM;

    var srvs=[
        {n:"01",t:"CONTROL, NO SOLO VIGILANCIA",
         d:"BitÃ¡cora digital, rondines con QR, monitoreo CCTV con operador dedicado. El cliente tiene visibilidad del 100% de la operaciÃ³n en tiempo real. Cero zonas grises, cero turnos sin evidencia.",
         metric:"100% trazabilidad documental"},
        {n:"02",t:"INTEGRACIÃ“N INDUSTRIAL",
         d:"Conocemos los ritmos de producciÃ³n automotriz. Nuestros supervisores coordinan con superintendentes de planta. Nos adaptamos al sistema del cliente, no al revÃ©s. Protocolo de arranque en 15 dÃ­as.",
         metric:"15 dÃ­as hÃ¡biles arranque pleno"},
        {n:"03",t:"REDUCCIÃ“N DE PÃ‰RDIDAS ECONÃ“MICAS",
         d:"Un paro de lÃ­nea FCA cuesta $420,000â€“$1.2M MXN/hora. Nuestro costo mensual es menor que 1 hora de paro. AnÃ¡lisis de ROI disponible para presentaciÃ³n a DirecciÃ³n General.",
         metric:"ROI positivo desde el Mes 1"},
        {n:"04",t:"TRAZABILIDAD EN TIEMPO REAL",
         d:"Cada guardia, cada rondÃ­n, cada acceso queda registrado con timestamp y evidencia. Dashboard mensual de KPIs entregado al Ã¡rea de compras y seguridad del cliente. Disponible para auditorÃ­a STPS.",
         metric:"Dashboard mensual garantizado"}
    ];
    var sw=A4W-MRG*2;
    var sh=42*MM;
    for(var i=0;i<srvs.length;i++){
        var s=srvs[i];
        var sx=O*A4W+MRG;
        var sy=y+i*(sh+3*MM);
        rect(doc,sx,sy,sw,sh,WHITE,MGRAY,0.4);
        // Barra izq roja con nÃºmero
        rect(doc,sx,sy,22*MM,sh,RED);
        txt(doc,s.n,sx+2*MM,sy+5*MM,14,WHITE,FB,18*MM,"C");
        // LÃ­nea divisoria
        line(doc,sx+22*MM,sy+5*MM,sx+22*MM,sy+sh-5*MM,RED2,0.5);
        // Contenido
        txt(doc,s.t,sx+25*MM,sy+5*MM,8.5,BLACK,FB,sw-28*MM);
        txt(doc,s.d,sx+25*MM,sy+14*MM,7,DGRAY,FR,sw-28*MM);
        // Badge mÃ©trica
        rect(doc,sx+sw-45*MM,sy+sh-8*MM,42*MM,6*MM,LGRAY,MGRAY,0.3);
        txt(doc,"â†’ "+s.metric,sx+sw-44*MM,sy+sh-5.5*MM,6,RED,FS,40*MM);
    }
    y+=4*sh+3*MM*3+8*MM;

    // Tabla comparativa
    txt(doc,"MICSA SAFETY vs. SERVICIO ESTÃNDAR",O*A4W+MRG,y,11,BLACK,FB,A4W-MRG*2); y+=7*MM;
    line(doc,O*A4W+MRG,y,O*A4W+A4W-MRG,y,RED,1); y+=5*MM;
    var comp=[
        ["Trazabilidad documental","100% bitÃ¡cora digital","Registro en papel, sin respaldo"],
        ["Tiempo de respuesta","< 3 minutos garantizados","Sin protocolo definido"],
        ["SustituciÃ³n de elemento","< 4 horas garantizadas","Variable, sin compromiso"],
        ["Reportes al cliente","Dashboard mensual + incidentes","Verbal o a solicitud"],
        ["Cumplimiento normativo","NOM-030 + REPSE vigente","Sin verificaciÃ³n documentada"]
    ];
    tabla(doc,O,y,["CRITERIO","MICSA SAFETY DIVISION","SERVICIO ESTÃNDAR"],comp,[0.33,0.33,0.34]);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// AB6 â€” KPIs Y CONTROL DE RENDIMIENTO
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function kpis(doc){
    var O=5;
    rect(doc,O*A4W,0,A4W,A4H,WHITE);
    header(doc,O,"KPIs Y CONTROL DE RENDIMIENTO",6);
    footer(doc,O);

    var y=22*MM;
    y=sec(doc,O,y,"DASHBOARD DE CONTROL MENSUAL");
    txt(doc,"Cada mes el cliente recibe este dashboard completo. Si algÃºn indicador baja de la meta, MICSA Safety activa plan de mejora en 48 horas.",
        O*A4W+MRG,y,8,DGRAY,FL,A4W-MRG*2); y+=12*MM;

    // MÃ©tricas visuales
    var mets=[
        {v:"â‰¥ 97%",t:"PUNTUALIDAD",d:"Tolerancia: 5 min\n3 faltas en 30 dÃ­as = baja"},
        {v:"â‰¥ 95%",t:"RONDINES\nCUMPLIDOS",d:"QR por punto de control\nRegistro timestamp"},
        {v:"100%",t:"BITÃCORAS\nCOMPLETAS",d:"Sin firma = turno incompleto\nSanciÃ³n directa"},
        {v:"â‰¤ 5%",t:"ROTACIÃ“N\nMENSUAL",d:"EvaluaciÃ³n semanal\nExpediente individual"}
    ];
    var mw=(A4W-MRG*2)/4;
    var mh=38*MM;
    for(var m=0;m<mets.length;m++){
        var mx=O*A4W+MRG+m*mw;
        var my=y;
        rect(doc,mx,my,mw-1*MM,mh,WHITE,MGRAY,0.5);
        rect(doc,mx,my,mw-1*MM,2*MM,RED);
        txt(doc,mets[m].v,mx+1*MM,my+6*MM,15,RED,FB,(mw-3*MM),"C");
        line(doc,mx+3*MM,my+17.5*MM,mx+mw-4*MM,my+17.5*MM,MGRAY,0.4);
        txt(doc,mets[m].t,mx+1*MM,my+20*MM,7,BLACK,FB,mw-3*MM,"C");
        txt(doc,mets[m].d,mx+2*MM,my+27*MM,5.5,DGRAY,FL,mw-4*MM,"C");
    }
    y+=mh+8*MM;

    // Tabla KPI operativo
    y=sec(doc,O,y,"TABLA DE SEGUIMIENTO OPERATIVO MENSUAL");
    var kRows=[
        ["Puntualidad (%)","â‰¥ 97%","_____ %","_____","_____"],
        ["BitÃ¡coras completas","100%","_____ %","_____","_____"],
        ["Rondines cumplidos","â‰¥ 95%","_____ %","_____","_____"],
        ["Incidentes reportados","0 omisiones","_____","_____","_____"],
        ["Incidentes detectados","MÃ¡x. detecciÃ³n","_____","_____","_____"],
        ["Tiempo de respuesta","< 3 minutos","_____ min","_____","_____"],
        ["RotaciÃ³n de personal","â‰¤ 5%","_____ %","_____","_____"],
        ["Sustituciones en tiempo","100% < 4h","_____ %","_____","_____"]
    ];
    y=tabla(doc,O,y,["KPI","META","REAL","DESV.","ACCIÃ“N"],kRows,[0.32,0.14,0.14,0.12,0.28]);
    y+=8*MM;

    // Firma supervisiÃ³n
    var fw=A4W-MRG*2;
    rect(doc,O*A4W+MRG,y,fw,22*MM,LGRAY,MGRAY,0.4);
    rect(doc,O*A4W+MRG,y,fw,1.2*MM,RED);
    txt(doc,"Supervisor responsable: _______________________________",O*A4W+MRG+4*MM,y+6*MM,8,DGRAY,FR,fw*0.55);
    txt(doc,"Firma: ___________________",O*A4W+MRG+fw*0.58,y+6*MM,8,DGRAY,FR,fw*0.38);
    txt(doc,"Fecha de reporte: _____________________   Mes de referencia: ______________________",O*A4W+MRG+4*MM,y+14*MM,7.5,DGRAY,FR,fw-8*MM);
    y+=28*MM;

    // Bloque cierre
    rect(doc,O*A4W+MRG,y,fw,22*MM,RED);
    rect(doc,O*A4W+MRG,y,fw,1.2*MM,BLACK);
    txt(doc,'"No cuidamos instalaciones. Controlamos operaciones."',O*A4W+MRG+5*MM,y+5*MM,10,WHITE,FS,fw-10*MM);
    txt(doc,"contacto@micsa.mx  Â·  800-MICSA-01  Â·  www.micsa.mx  Â·  REPSE Vigente  Â·  NOM-030 Certificado",
        O*A4W+MRG+5*MM,y+14*MM,7,WHITE,FL,fw-10*MM);
}

// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
// MAIN
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
function main(){
    var doc=app.documents.add(DocumentColorSpace.CMYK, A4W*6, A4H);

    for(var i=1;i<6;i++) doc.artboards.add([i*A4W,0,(i+1)*A4W,-A4H]);
    var names=["01_PORTADA","02_FILOSOFIA","03_MODELO_3CAPAS","04_LIDERAZGO","05_SERVICIOS","06_KPIS"];
    for(var i=0;i<6;i++) doc.artboards[i].name=names[i];

    portada(doc);
    filosofia(doc);
    modelo(doc);
    liderazgo(doc);
    servicios(doc);
    kpis(doc);

    var saveFile=new File(BASE_PATH+"MICSA_SAFETY_SISTEMA.ai");
    var opts=new IllustratorSaveOptions();
    opts.compatibility=Compatibility.ILLUSTRATOR24;
    opts.saveMultipleArtboards=true;
    doc.saveAs(saveFile,opts);

    try{
        var pdfFile=new File(OUT_PATH+"MICSA_SAFETY_SISTEMA.pdf");
        var pdfOpts=new PDFSaveOptions();
        pdfOpts.compatibility=PDFCompatibility.ACROBAT8;
        pdfOpts.generateThumbnails=true;
        pdfOpts.preserveEditability=false;
        pdfOpts.saveMultipleArtboards=true;
        pdfOpts.artboardRange="1-6";
        doc.saveAs(pdfFile,pdfOpts);
    }catch(e){}

    alert("MICSA_SAFETY_SISTEMA.ai generada â€” 6 artboards.\n\nPortada Â· FilosofÃ­a Â· Modelo 3 Capas Â· Liderazgo (Gerardo GuzmÃ¡n) Â· Servicios Â· KPIs\n\nPDF exportado en PDFs/");
}
main();



