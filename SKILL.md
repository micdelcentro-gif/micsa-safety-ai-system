---
name: micsa-safety-ai-system
description: Usa esta habilidad cuando el usuario solicite diseñar o maquetar el sistema editorial corporativo, la presentación comercial, el organigrama o los manuales operativos de MICSA Safety Division en formato Adobe Illustrator (.AI). 
---

# OBJETIVO PRINCIPAL
Diseñar la estructura visual, el layout y la composición de documentos corporativos de nivel premium en formato Adobe Illustrator (.AI editable) para MICSA SAFETY DIVISION. El documento es una herramienta de ventas y control para clientes del sector industrial y automotriz.

# 🗺️ MAPA DE RECURSOS (SCRIPTS .JSX)
Los scripts para Adobe Illustrator se encuentran en `C:\Users\Tecnolaps\OneDrive\Desktop\MICSA_Safety\`:

| Script | Identidad Visual | Páginas | Propósito |
| :--- | :--- | :--- | :--- |
| `MICSA_SAFETY_SISTEMA.jsx` | **Corporate (Rojo/Negro)** | 6 | Sistema operativo integral, filosofía, modelo 3 capas, liderazgo (Gerardo Guzmán). |
| `MICSA_SAFETY_TACTICO.jsx` | **Tactical (Negro/Plata)** | 6 | Presentación técnica, cultura operativa, plan de capacitación 5 fases, ruta crítica. |
| `MICSA_SAFETY_PRESENTACION.jsx` | **Premium (Navy/Gold)** | 7 | Presentación corporativa completa, matriz de evaluación, mockup de papelería (tarjetas, credenciales). |

# 🎨 IDENTIDADES VISUALES Y ESTILOS
* **Corporate (RED):** Rojo dominante (C0 M90 Y85 K0), Blanco, Negro rico. Tipografía: Montserrat Bold / Open Sans.
* **Tactical (SILVER):** Negro dominante, Plateado (K25), Gris oscuro. Estilo minimalista industrial/militar.
* **Premium (GOLD):** Azul Marino (C85 M65 Y0 K40), Dorado metálico. Estilo institucional de alto nivel.

# 🏗️ ESTRUCTURA Y COMPONENTES
Cuando la IA (Cline/Copilot/Claude) edite estos scripts, debe seguir estas funciones:
* `rect(doc,x,y,w,h,fill,stroke,sw)`: Creación de bloques y fondos.
* `txt(doc,s,x,y,sz,col,font,w,align)`: Gestión tipográfica avanzada.
* `logoPlace(doc,x,y,w)`: Inserción dinámica del logotipo de MICSA.
* `tabla(doc,O,y,hdrs,rows,cols)`: Generación de dashboards y comparativas de rendimiento.

# 📋 DATOS CLAVE DEL SISTEMA
* **Director de Seguridad:** Gerardo Guzmán Alvarado.
* **Experiencia:** 5 años Fuerza Civil + Protección Civil + Bomberos.
* **Frase Portada:** "No cuidamos instalaciones. Controlamos operaciones."
* **Frase Cierre:** "Puede faltar flujo, pero nunca posición."

# 🤖 INSTRUCCIONES PARA LA IA
1.  **Consistencia de Marca:** No mezcles paletas de colores entre scripts.
2.  **Rutas de Archivo:** Asegura que `BASE_PATH` y `LOGO_FILE` sean correctos antes de enviar al usuario.
3.  **Artboards:** Mantén el tamaño Carta o A4 según el script original.
4.  **Layout:** Usa la retícula (grid) definida por las constantes `MRG` (margen) y `MM` (milímetros).
