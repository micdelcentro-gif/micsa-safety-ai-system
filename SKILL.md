---
name: micsa-safety-ai-system
description: Usa esta habilidad cuando el usuario solicite diseñar o maquetar el sistema editorial corporativo, la presentación comercial, el organigrama o los manuales operativos de MICSA Safety Division en formato Adobe Illustrator (.AI). 
---

# 🎯 OBJETIVO PRINCIPAL

Diseñar la estructura visual, el layout y la composición de documentos corporativos de nivel premium en formato Adobe Illustrator (.AI editable) para MICSA SAFETY DIVISION. El documento es una herramienta de ventas y control para clientes del sector industrial y automotriz.

## 🗺️ MAPA DE RECURSOS (SCRIPTS .JSX)

Los scripts para Adobe Illustrator se encuentran ahora centralizados en la carpeta de automatización del repositorio:
`C:\Users\Tecnolaps\OneDrive\Desktop\micsa-safety-ai-system\design-automation\`

| Script | Identidad Visual | Páginas | Propósito |
| :--- | :--- | :--- | :--- |
| `MICSA_SAFETY_SISTEMA.jsx` | **Corporate (Rojo/Negro)** | 6 | Sistema operativo integral, filosofía, modelo 3 capas, liderazgo (Gerardo Guzmán). |
| `MICSA_SAFETY_TACTICO.jsx` | **Tactical (Negro/Plata)** | 6 | Presentación técnica, cultura operativa, plan de capacitación 5 fases, ruta crítica. |
| `MICSA_SAFETY_PRESENTACION.jsx` | **Premium (Navy/Gold)** | 7 | Presentación corporativa completa, matriz de evaluación, mockup de papelería. |
| `MICSA_COTIZACION_FCA.jsx` | **Tactical (Negro/Plata)** | 7 | Cotización formal Stellantis/FCA, desglose de nómina por plaza (21 elementos), resumen financiero. |
| `MICSA_MANUAL_SEGURIDAD.jsx`| **Tactical (Negro/Plata)** | 8 | Manual operativo industrial, marco legal NOM, protocolos de emergencia y KPIs. |
| `MICSA_CULTURA_COMPLETO.jsx`| **Tactical (Negro/Plata)** | 8 | Identidad y cultura organizacional, visión interna, valores y plan de carrera. |
| `MICSA_PLAN_IMPLEMENTACION.jsx` | **Tactical (Negro/Plata)** | 7 | Ruta crítica Gantt 8 semanas, checklist arranque, protocolo emergencias, KPIs operativos. |

## 🎨 IDENTIDADES VISUALES Y ESTILOS

* **Corporate (RED):** Rojo dominante (C0 M90 Y85 K0), Blanco, Negro rico. Tipografía: Montserrat Bold / Open Sans.
* **Tactical (SILVER):** Negro dominante, Plateado (K25), Gris oscuro. Estilo minimalista industrial/militar.
* **Premium (GOLD):** Azul Marino (C85 M65 Y0 K40), Dorado metálico. Estilo institucional de alto nivel.

## 🏗️ ESTRUCTURA Y COMPONENTES

Cuando la IA (Cline/Copilot/Claude) edite estos scripts, debe seguir estas funciones:

* `rect(doc,x,y,w,h,fill,stroke,sw)`: Creación de bloques y fondos.
* `txt(doc,s,x,y,sz,col,font,w,align)`: Gestión tipográfica avanzada.
* `logoPlace(doc,x,y,w)`: Inserción dinámica del logotipo de MICSA.
* `tabla(doc,O,y,hdrs,rows,cols)`: Generación de dashboards y comparativas de rendimiento.
* `FILASUM(doc,O,y,label,valor,highlight)`: Variante de tabla para totales financieros con alineación derecha.

## 📋 DATOS CLAVE DEL SISTEMA

* **Director de Seguridad:** Gerardo Guzmán Alvarado.
* **Experiencia:** 5 años Fuerza Civil + Protección Civil + Bomberos.
* **Frase Portada:** "No cuidamos instalaciones. Controlamos operaciones."
* **Frase Cierre:** "Puede faltar flujo, pero nunca posición."

## 🤖 INSTRUCCIONES PARA LA IA

1. **Consistencia de Marca:** No mezcles paletas de colores entre scripts.
2. **Rutas de Archivo:** Las rutas reales en todos los scripts son:
   - `BASE = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\MICSA_Safety\\"` — carpeta de trabajo y salida de `.ai`
   - `LOGO = "C:\\Users\\Tecnolaps\\OneDrive\\Desktop\\adobe\\ID-5608-20260329T034218Z-1-001\\ID-5608\\LOGO 7\\editable\\editable.ai"` — logotipo vectorial
   - `OUT  = BASE + "PDFs\\"` — carpeta de exportación PDF
   No cambies estas rutas a menos que el usuario mueva físicamente los archivos.
3. **Artboards:** Mantén el tamaño Carta o A4 según el script original.
4. **Layout:** Usa la retícula (grid) definida por las constantes `MRG` (margen) y `MM` (milímetros).
5. **Ubicación de Scripts:** Todos los `.jsx` deben guardarse siempre en `design-automation/`.
