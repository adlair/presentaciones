#!/usr/bin/env python3
"""Genera la charla GNOME Honeycomb (26 diapositivas; 50 minutos).
Requiere python-pptx. No modifica el repositorio de la extensión.
"""
from pathlib import Path
from io import BytesIO
from urllib.request import urlopen
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "GNOME_Honeycomb_Charla_50_Minutos.pptx"
BG = RGBColor(13, 20, 34)
WHITE = RGBColor(238, 243, 250)
MUTED = RGBColor(168, 186, 205)
CYAN = RGBColor(55, 216, 230)
FOOT = "GNOME HONEYCOMB  •  DE LA IDEA A LA REVISIÓN TÉCNICA"

# Formato: título, ideas (separadas por |), imagen opcional.
DATA = [
("GNOME Honeycomb", "De una cuadrícula plana a una experiencia orbital|Diseño · GJS · Matemáticas · Validación · EGO", "orbital-1.jpg"),
("La idea: ¿por qué cambiar Applications?", "Una cuadrícula paginada funciona, pero limita la exploración visual.|Quería un espacio continuo, rotatorio y centrado.|La meta no era reemplazar GNOME: era integrarse con su comportamiento.", "orbital-2.jpg"),
("Qué construimos", "Una extensión de GNOME Shell escrita en JavaScript/GJS.|Distribución determinista de aplicaciones sobre una esfera virtual.|Arrastre para rotar, magnetismo del puntero y profundidad visual.|Búsqueda, Overview, lanzamientos y paginación nativa preservados.", "orbital-3.jpg"),
("Demostración inicial: 3 minutos", "Abrir Applications y rotar el panal.|Acercar el puntero y seleccionar una aplicación.|Buscar con teclado y regresar con Escape.|Mostrar la cuadrícula nativa al deshabilitar.", "orbital-1.jpg"),
("El ecosistema técnico", "GNOME Shell 50.5 y Mutter 50.5 en Wayland.|GJS: JavaScript integrado con bibliotecas GNOME.|Clutter/St: actores, transformación, eventos y renderizado.|Shell.AppSystem, AppDisplay y ParentalControlsManager.", None),
("Arquitectura en una imagen", "Overview → ControlsManager → AppDisplay.|AppDisplay contiene _scrollView y _grid.|Honeycomb crea viewport, canvas e iconos propios.|Los actores nativos se ocultan temporalmente y se restauran.", None),
("Decisión crucial: no mover los iconos nativos", "El primer enfoque dejó rastros de reparenting.|La arquitectura consolidada crea sus propios actores visuales.|Los tiles originales siguen siendo propiedad de GNOME.|Esto simplifica la restauración y reduce interferencias.", None),
("Matemáticas: esfera de Fibonacci", "Distribuye N puntos aproximadamente de forma uniforme.|El índice de cada aplicación determina una posición reproducible.|No exige filas, columnas ni páginas en el panal.|El número de aplicaciones puede variar.", None),
("De la esfera 3D a una pantalla 2D", "Yaw y pitch rotan el universo sin desplazar su centro.|La proyección calcula coordenadas de pantalla.|Profundidad controla tamaño, opacidad y orden visual.|Pivote centrado: cada icono escala alrededor de su centro.", None),
("Interacción: magnetismo y rotación", "El arrastre modifica yaw y pitch.|La cercanía del puntero aumenta el protagonismo local.|La interpolación suaviza las transiciones.|Después de arrastrar, se suprime el clic accidental.", "orbital-2.jpg"),
("La primera dificultad: geometría real", "Un actor puede existir sin tener allocation válida.|Tamaño solicitado no equivale a tamaño ya calculado.|La sincronización con Clutter determina cuándo presentar.|Un estado prematuro puede provocar picking incorrecto.", None),
("Búsqueda: el problema más difícil", "GNOME mantiene su propia búsqueda y sus resultados.|Honeycomb debe dejar de capturar input cuando aparece Search.|Al salir, prepara contenido invisible y no interactivo.|Solo presenta el panal cuando las condiciones vuelven a ser válidas.", None),
("El salto visual que todavía vemos", "Primera apertura: transición visible de grid a panal.|Tras una búsqueda + Escape: el salto se repite.|La funcionalidad fue aceptada en pruebas manuales.|No se afirma una experiencia completamente libre de flicker.", None),
("Paginación: preservar el estado del usuario", "No forzar página cero.|Medir el _scrollView estable, no la cuadrícula desplazada.|Guardar y restaurar visibilidad/opacidad/reactividad.|Conservar controles, scroll y navegación nativa.", None),
("Política parental: GNOME decide", "Descubrimiento desde Shell.AppSystem.|Filtrado con el administrador parental nativo.|Si el filtro falla: no mostrar la aplicación.|Revalidar permisos justo antes del lanzamiento.", None),
("Lifecycle: enable, disable y cleanup", "Cada señal, fuente y actor necesita propietario.|Cancelar callbacks pendientes y generaciones obsoletas.|Restaurar actores nativos antes de liberar recursos propios.|Evitar acceder a objetos destruidos o dejar handlers vivos.", None),
("Cómo probamos sin arriesgar el escritorio", "Tests Node con mocks y regresiones RED/GREEN.|QEMU/KVM con Fedora y GNOME 50.5.|TigerVNC sobre sockets Unix privados; sin red invitada.|Snapshots, hashes y apagado supervisado.", None),
("Lo que nos enseñaron las máquinas virtuales", "Un test puede fallar por readiness, no por el producto.|Un timeout 124 no equivale a apagado limpio.|La memoria disponible y los permisos QCOW2 también importan.|Evidencia aislada y reproducible evita conclusiones falsas.", None),
("Alternative D: experimentar y saber detenerse", "Se probó una composición alternativa de búsqueda.|La evaluación humana la percibió ligeramente mejor.|Persistieron saltos visuales y dudas de atribución.|Decisión: congelar D y no integrarla sin evidencia suficiente.", None),
("La revisión de GNOME Extensions", "Primera revisión de EGO: tres objeciones.|Exceso de try/catch vacíos.|Comprobaciones redundantes de APIs garantizadas.|Pregunta sobre procedencia y mantenibilidad del código.", None),
("La remediación P1 → P5", "P1: APIs y tipos reales de Clutter/St.|P2: ruta canónica de AppDisplay y _scrollView.|P3: AppInfo, Shell.App y filtro parental.|P4: propiedad de recursos y eliminación de catches vacíos.|P5: retirar estado inerte y mejorar documentación.", None),
("Las pruebas crecieron con cada corrección", "145 pruebas en la auditoría A/C.|167 tras la primera limpieza; 271 antes de P1.|288 → 305 → 326 → 359 → 375 en P1–P5.|378 pruebas automatizadas al cierre del último ajuste.", None),
("Preparar una entrega responsable", "ZIP de cuatro archivos: extension.js, metadata, CSS y LICENSE.|Hashes SHA-256 y comparación byte a byte.|Respaldo reversible antes de instalación local.|Logout/login Wayland para cargar el módulo nuevo.", None),
("Estado real del proyecto", "Candidato final instalado y verificado en disco.|Último smoke del candidato aún pendiente en el punto documentado.|Salto visual conocido y aceptado como limitación.|Sin publicación nueva confirmada en EGO en la evidencia disponible.", None),
("Cinco aprendizajes para otros desarrolladores", "Diseñar alrededor de contratos reales, no de APIs imaginadas.|El ciclo de vida importa tanto como el renderizado.|Un PASS de mocks no sustituye interacción humana.|Documentar fallos y límites da credibilidad.|La revisión de código es parte del producto.", None),
("Cierre y preguntas", "GNOME Honeycomb es una exploración de interacción espacial.|La experiencia se construyó mediante prototipos, pruebas y correcciones.|El siguiente paso es completar el smoke final y la revisión de EGO.|Repositorio: github.com/adlair/gnome-honeycomb", "orbital-3.jpg"),
]

# Distribución 50 minutos incluyendo 3 de demostración y 3 para preguntas.
MINUTES = [2,2,2,3,1,2,2,2,2,2,2,3,2,2,2,2,2,1,2,2,3,2,1,1,1,3]
assert len(DATA) == 26 and len(MINUTES) == 26
# Ajuste explícito a 50 minutos: la suma de tiempos se recalcula abajo.
MINUTES = [1,2,2,3,2,2,2,2,2,2,2,3,2,2,2,2,2,1,1,2,3,2,1,1,1,3]
assert sum(MINUTES) == 50, sum(MINUTES)

def box(slide, x, y, w, h, text, size, color=WHITE, bold=False, spacing=1.05):
    sh = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = sh.text_frame
    tf.clear()
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.space_after = Pt(0)
    for run in p.runs:
        run.font.name = "Cantarell"
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
    return sh

def picture(slide, filename, x, y, w, h):
    raw = "https://raw.githubusercontent.com/adlair/gnome-honeycomb/master/screenshots/" + filename
    try:
        b = urlopen(raw, timeout=30).read()
        slide.shapes.add_picture(BytesIO(b), Inches(x), Inches(y), width=Inches(w), height=Inches(h))
        return True
    except Exception as ex:
        box(slide, x, y+0.5, w, 1.2, "Captura orbital disponible en github.com/adlair/gnome-honeycomb", 17, MUTED)
        print("Imagen no disponible:", filename, ex)
        return False

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
for number, (title, body, photo) in enumerate(DATA, 1):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = BG
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.44), Inches(0.6), Inches(0.07), Inches(0.63))
    accent.fill.solid(); accent.fill.fore_color.rgb = CYAN; accent.line.fill.background()
    box(slide, 0.69, 0.53, 12.0, 0.78, title, 29, WHITE, True)
    box(slide, 0.72, 7.06, 11.0, 0.25, FOOT, 8, MUTED)
    box(slide, 12.16, 7.06, 0.52, 0.25, f"{number:02d}", 9, CYAN)
    if number == 1:
        box(slide, 0.72, 1.36, 11.8, 0.45, "Cómo construí un universo de aplicaciones para GNOME Shell · Charla técnica · 50 minutos", 13, MUTED)
        box(slide, 0.9, 2.50, 5.2, 2.7, body.split("|")[0], 29, WHITE, True)
        box(slide, 0.9, 5.8, 5.2, 0.55, body.split("|")[1], 15, CYAN)
        picture(slide, photo, 6.28, 2.2, 6.33, 4.02)
    else:
        parts = body.split("|")
        with_image = photo is not None
        x, w = 0.83, (7.1 if with_image else 11.9)
        font = 19 if len(parts) <= 4 else 18
        y = 1.95
        for t in parts:
            box(slide, x, y, w, 0.92 if len(t) > 75 else 0.72, "• " + t, font)
            y += 1.08 if len(t) > 75 else 0.92
        if with_image:
            picture(slide, photo, 8.35, 2.92, 4.25, 2.74)
    note = slide.notes_slide.notes_text_frame
    note.text = (f"Duración sugerida: {MINUTES[number-1]} minuto(s).\n"
                 + ("Mostrar el comportamiento real; no afirmar que el salto visual desapareció." if number in (4,13) else
                    "Distinguir pruebas automatizadas de validaciones nativas reales." if number in (17,18,22,24) else
                    "Explicar la decisión técnica con ejemplos y conectar con la siguiente diapositiva."))
OUT.parent.mkdir(parents=True, exist_ok=True)
prs.save(OUT)
print(f"Generado: {OUT} | diapositivas={len(prs.slides)} | duración={sum(MINUTES)} minutos")
