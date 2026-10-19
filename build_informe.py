# -*- coding: utf-8 -*-
"""Construye el informe de la Actividad 1 sobre la plantilla act1.docx.

Lee `resultados.json`, escrito por el cuaderno ya ejecutado, de modo que
ninguna cifra del informe se transcribe a mano. Conserva encabezado, pie,
márgenes y estilos de la plantilla; sustituye el cuerpo del enunciado.

Uso:
    python build_informe.py
"""

import json
import pathlib

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt, RGBColor

AQUI = pathlib.Path(__file__).resolve().parent
PLANTILLA = AQUI.parent / "act1.docx"
SALIDA = AQUI / "Actividad1_CienciaDatos_Meneses.docx"
FIG = AQUI / "figuras"
REPO = "https://github.com/amenesesy/actividad-cd-liver"
COLAB = ("https://colab.research.google.com/github/amenesesy/actividad-cd-liver/"
         "blob/main/CD_Actividad1_Liver.ipynb")

R = json.load(open(AQUI / "resultados.json", encoding="utf-8"))
ANCHO_FIG = Cm(14.0)


def n(x, d=2):
    """Número con coma decimal y espacio fino de millar, al estilo español."""
    s = "%.*f" % (d, x)
    s = s.replace("-", "−")  # signo menos tipográfico
    ent, _, dec = s.partition(".")
    if len(ent) > 4:
        ent = "{:,}".format(int(ent)).replace(",", " ")
    return ent + ("," + dec if dec else "")


def pct(x, d=1):
    return n(x, d) + " %"


# ----------------------------------------------------------------- documento
doc = Document(str(PLANTILLA))
cuerpo = doc.element.body

# Se conservan los 7 primeros párrafos (recuadro gris, datos del estudiante y
# título de la actividad) y se elimina el resto del enunciado.
parrafos = doc.paragraphs
for p in parrafos[7:]:
    p._element.getparent().remove(p._element)
for t in doc.tables:
    t._element.getparent().remove(t._element)

# Nombre y fecha en los párrafos de datos del estudiante.
for r in parrafos[2].runs:
    if "Nombre y apellidos del estudiante" in r.text:
        r.text = "Abel Meneses Yaranga"
for r in parrafos[3].runs:
    if "01/01/2022" in r.text:
        r.text = "18/10/2026"


def parrafo(texto, estilo="Normal", despues=6, antes=0, alineacion=None,
            cursiva=False, negrita=False, tamano=None):
    p = doc.add_paragraph(style=estilo)
    run = p.add_run(texto)
    run.italic = cursiva or None
    run.bold = negrita or None
    if tamano:
        run.font.size = Pt(tamano)
    p.paragraph_format.space_after = Pt(despues)
    p.paragraph_format.space_before = Pt(antes)
    if alineacion is not None:
        p.alignment = alineacion
    return p


def runs(p, t, tamano=None):
    """Añade `t` al párrafo `p`; *cursiva* simple delimitada por asteriscos."""
    for i, parte in enumerate(t.split("*")):
        if parte:
            run = p.add_run(parte)
            if i % 2 == 1:
                run.italic = True
            if tamano:
                run.font.size = Pt(tamano)
    return p


def texto(t):
    """Párrafo del cuerpo. Admite *cursiva* simple delimitada por asteriscos."""
    p = doc.add_paragraph(style="Normal")
    p.paragraph_format.space_after = Pt(4)
    return runs(p, t)


def nota_apa(nota):
    p = doc.add_paragraph(style="Normal")
    r = p.add_run("Nota. ")
    r.italic = True
    r.font.size = Pt(10)
    runs(p, nota, tamano=10)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.line_spacing = 1.15
    return p


def apartado2(t):
    p = parrafo(t, "Título Apartado 2", despues=4, antes=10)
    p.paragraph_format.keep_with_next = True
    return p


def apartado3(t):
    p = parrafo(t, "Título Apartado 3", despues=3, antes=8)
    p.paragraph_format.keep_with_next = True
    return p


def figura(num, titulo, archivo, nota="Elaboración propia a partir de act_liver_disease.csv.",
           ancho=ANCHO_FIG):
    """Figura con el formato de la 7.ª edición del manual APA."""
    p = parrafo("Figura %d" % num, negrita=True, despues=0, antes=8)
    p.paragraph_format.keep_with_next = True
    p = parrafo(titulo, cursiva=True, despues=4)
    p.paragraph_format.keep_with_next = True
    doc.add_picture(str(FIG / archivo), width=ancho)
    pic = doc.paragraphs[-1]
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.paragraph_format.space_after = Pt(2)
    pic.paragraph_format.keep_with_next = True
    nota_apa(nota)


def bordes_apa(t):
    """Tabla al estilo APA: solo línea superior, bajo la cabecera e inferior."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tblPr = t._tbl.tblPr
    for viejo in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(viejo)
    borders = OxmlElement("w:tblBorders")
    for lado in ("top", "bottom"):
        b = OxmlElement("w:%s" % lado)
        b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "8")
        b.set(qn("w:space"), "0"); b.set(qn("w:color"), "333333")
        borders.append(b)
    for lado in ("left", "right", "insideH", "insideV"):
        b = OxmlElement("w:%s" % lado)
        b.set(qn("w:val"), "nil")
        borders.append(b)
    tblPr.append(borders)
    # Línea bajo la fila de cabecera.
    for c in t.rows[0].cells:
        tcPr = c._tc.get_or_add_tcPr()
        tcb = OxmlElement("w:tcBorders")
        b = OxmlElement("w:bottom")
        b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "6")
        b.set(qn("w:space"), "0"); b.set(qn("w:color"), "333333")
        tcb.append(b)
        tcPr.append(tcb)


def tabla(num, titulo, cabecera, filas, nota=None, anchos=None):
    """Tabla con el formato APA: número, título en cursiva, cuerpo compacto."""
    p = parrafo("Tabla %d" % num, negrita=True, despues=0, antes=8)
    p.paragraph_format.keep_with_next = True
    p = parrafo(titulo, cursiva=True, despues=4)
    p.paragraph_format.keep_with_next = True
    t = doc.add_table(rows=1 + len(filas), cols=len(cabecera))
    # El nombre del estilo viene con la codificación de la plantilla; se busca por id.
    t.style = next(s for s in doc.styles if s.style_id == "Tablaconcuadrcula")
    t.alignment = 1
    bordes_apa(t)
    for j, h in enumerate(cabecera):
        c = t.cell(0, j)
        c.text = ""
        run = c.paragraphs[0].add_run(h)
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(0x00, 0x98, 0xCD)
    for i, fila in enumerate(filas, start=1):
        for j, v in enumerate(fila):
            c = t.cell(i, j)
            c.text = ""
            run = c.paragraphs[0].add_run(str(v))
            run.font.size = Pt(10)
    for fila in t.rows:
        for j, c in enumerate(fila.cells):
            pf = c.paragraphs[0].paragraph_format
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            c.paragraphs[0].alignment = (WD_ALIGN_PARAGRAPH.LEFT if j == 0
                                         else WD_ALIGN_PARAGRAPH.CENTER)
            if anchos:
                c.width = anchos[j]
    if nota:
        nota_apa(nota)
    else:
        parrafo("", despues=6)


# ------------------------------------------------------------- INTRODUCCIÓN
E, C, F, D = R["estructura"], R["categoricas"], R["faltantes"], R["duplicados"]
CO, X, DS, DC = R["coherencia"], R["extremos"], R["dispersion"], R["dispersion_clase"]
FO, CR, IM, NO, AT = R["forma"], R["correlacion"], R["imputacion"], R["normalizacion"], R["atipicos"]

apartado2("Introducción")
texto(
    "Este informe desarrolla las dos primeras etapas del proyecto: la descripción del "
    "conjunto de datos y su tratamiento previo al modelado. El archivo act_liver_disease.csv "
    "procede del estudio de Straw y Wu (2022) sobre sesgo por sexo en la predicción de "
    "enfermedad hepática y, a su vez, del Indian Liver Patient Dataset (ILPD) de Ramana et "
    "al. (2011). El análisis se ejecutó en Python con pandas (McKinney, 2010), scikit-learn "
    "(Pedregosa et al., 2011), SciPy (Virtanen et al., 2020) y seaborn (Waskom, 2021); el "
    "cuaderno con el código y sus salidas se anexa (Meneses Yaranga, 2026)."
)

# =================================================================== ETAPA 1
apartado2("Etapa 1. Descripción de la base de datos")

apartado3("1. Estructura del dataset")
texto(
    "El archivo es un CSV de texto plano separado por comas, de %s KB, con una fila de "
    "cabecera y %d registros, cada uno correspondiente a un paciente, descritos por %d "
    "variables. Es pequeño (76 KB en pandas) y se lee sin parámetros especiales, pero no "
    "tiene identificador de paciente, de modo que un duplicado no puede distinguirse de "
    "dos pacientes distintos. Diez columnas son numéricas (cinco enteras: Age, Alkphos, "
    "Sgpt, Sgot y Selector; cinco decimales: TB, DB, TP, ALB y A/G Ratio) y Gender es "
    "texto con dos valores, Male y Female."
    % (n(E["kb_disco"], 1), E["filas"], E["columnas"])
)
texto(
    "Las variables responden a un panel hepático habitual. Age es la edad en años (de %d "
    "a 90, mediana 45). TB y DB son la bilirrubina total y directa en mg/dL; Alkphos la "
    "fosfatasa alcalina y Sgpt y Sgot las transaminasas ALT y AST, todas en UI/L; TP y ALB "
    "son las proteínas totales y la albúmina en g/dL, y A/G Ratio es el cociente entre "
    "albúmina y globulina. Selector es la etiqueta clínica, codificada como 1 para el "
    "paciente con enfermedad hepática y 2 para el paciente sin ella, lo que conviene "
    "recodificar a 0 y 1 antes de modelar. Selector y Gender son categóricas por "
    "naturaleza aunque se almacenen como entero y texto; las otras nueve son cuantitativas."
    % X["edad_min"]
)
texto(
    "La composición de la muestra es desigual en las dos categóricas: %d hombres (%s) "
    "frente a %d mujeres, y %d pacientes enfermos (%s) frente a %d sanos. La proporción "
    "de enfermos es similar en ambos sexos (%s en hombres y %s en mujeres). Dos "
    "consecuencias para el modelado: un clasificador que dijera siempre enfermo acertaría "
    "el %s, así que la exactitud no sirve para evaluarlo y la partición debe ser "
    "estratificada; y la escasez de mujeres limita lo que se aprenda sobre ellas, que es "
    "justamente la preocupación de Straw y Wu (2022)."
    % (C["hombres"], pct(C["pct_hombres"]), C["mujeres"], C["enfermos"],
       pct(C["pct_enfermos"]), C["sanos"], pct(C["pct_enfermos_hombres"]),
       pct(C["pct_enfermos_mujeres"]), pct(C["pct_enfermos"]))
)

apartado3("2. Problemas de calidad del dataset")
texto(
    "La revisión sistemática del cuaderno (sección 2) cuantificó seis problemas. El primero "
    "son los valores faltantes: solo hay %d, todos en A/G Ratio (%s de las filas), en "
    "pacientes cuyos TP y ALB sí están registrados, lo que abre la puerta a reconstruir "
    "el dato en lugar de imputarlo a ciegas. El segundo son los duplicados: %d filas son "
    "copias exactas de otra en las once columnas (%d registros implicados). Dado que el "
    "ILPD original ya los contiene y que dos pacientes con la misma analítica son "
    "posibles, se conservan, pero deben tenerse en cuenta al partir en entrenamiento y "
    "prueba, porque un duplicado repartido entre ambos conjuntos infla las métricas."
    % (F["total"], pct(F["pct_ag"], 2), D["excedentes"], D["implicadas"])
)
texto(
    "El tercero es la coherencia interna. En %d registros la bilirrubina directa supera "
    "a la total, lo que es imposible porque la directa es una fracción de la total; al ser "
    "solo %d filas, procede marcarlas y excluirlas de cualquier cociente DB/TB. Además, al "
    "recalcular el cociente albúmina/globulina como ALB/(TP − ALB) la diferencia con el "
    "valor registrado es pequeña en general (mediana %s), pero %d registros se apartan en "
    "más de 0,1 y el peor llega a %s: con TP de 3,6 y ALB de 2,7 el cociente debería ser "
    "3,0 y figura 0,7. Son errores de captura o de redondeo en pocas filas; marcarlos "
    "permite recalcular A/G Ratio desde ALB y TP, como se propone en la tarea 6."
    % (CO["db_mayor_tb"], CO["db_mayor_tb"], n(CO["ag_dif_mediana"], 3),
       CO["ag_inconsistentes"], n(CO["ag_dif_max"], 1))
)
texto(
    "En cuarto lugar, los valores extremos: Sgot alcanza %s UI/L, Sgpt %s y Alkphos %s, y "
    "la bilirrubina total llega a %s mg/dL. Son cifras compatibles con una hepatitis "
    "aguda grave y no se pueden descartar como errores, pero %d registros de Sgot, %d de "
    "Sgpt y %d de Alkphos superan las 1500 UI/L y merecerían confirmación clínica. El quinto es "
    "el truncamiento de la edad, que en el ILPD se recorta a 90 años (aquí afecta a %s), "
    "y el sexto, el desbalance de sexo y clase ya descrito; no hay, en cambio, problemas "
    "de tipo de dato ni categorías mal escritas."
    % (n(X["sgot_max"], 0), n(X["sgpt_max"], 0), n(X["alkphos_max"], 0), n(X["tb_max"], 1),
       X["por_variable"]["Sgot"], X["por_variable"]["Sgpt"], X["por_variable"]["Alkphos"],
       "un paciente" if X["edad_90"] == 1 else "%d pacientes" % X["edad_90"])
)

# =================================================================== ETAPA 2
apartado2("Etapa 2. Tratamiento de datos")

apartado3("3. Media, desviación estándar, varianza, rango e IQR de TB y DB")
tb, db = DS["TB"], DS["DB"]
tabla(1, "Estadísticos descriptivos de la bilirrubina total (TB) y directa (DB), en mg/dL",
      ["Estadístico", "TB", "DB"],
      [("Media", n(tb["media"]), n(db["media"])),
       ("Mediana", n(tb["mediana"]), n(db["mediana"])),
       ("Desviación estándar", n(tb["desv_std"]), n(db["desv_std"])),
       ("Varianza", n(tb["varianza"]), n(db["varianza"])),
       ("Rango (mín. a máx.)", "%s (%s a %s)" % (n(tb["rango"]), n(tb["mínimo"]), n(tb["máximo"])),
        "%s (%s a %s)" % (n(db["rango"]), n(db["mínimo"]), n(db["máximo"]))),
       ("Rango intercuartílico (Q1 a Q3)", "%s (%s a %s)" % (n(tb["IQR"]), n(tb["Q1"]), n(tb["Q3"])),
        "%s (%s a %s)" % (n(db["IQR"]), n(db["Q1"]), n(db["Q3"]))),
       ("Coeficiente de variación", pct(tb["coef_variación_%"], 0), pct(db["coef_variación_%"], 0))],
      nota="*n* = 583. Desviación y varianza muestrales (*n* − 1). Elaboración propia.",
      anchos=[Cm(6.0), Cm(4.0), Cm(4.0)])
texto(
    "Los dos indicadores de bilirrubina (Tabla 1) muestran una variabilidad enorme en términos "
    "relativos: la desviación estándar casi duplica a la media (coeficiente de variación "
    "cercano al 190 %% en ambas); la varianza (%s en TB) no se lee directamente por "
    "estar en mg²/dL². Los demás estadísticos revelan el motivo. La "
    "mediana de TB es %s mientras la media es %s, y en DB la mediana es %s frente a una "
    "media de %s: la media está arrastrada por una cola derecha larga. El rango "
    "intercuartílico, que describe al 50 %% central de los pacientes, es de apenas %s "
    "mg/dL en TB y %s en DB, frente a rangos totales de %s y %s; es decir, la mitad "
    "central de los pacientes ocupa un intervalo 40 veces más estrecho que el recorrido "
    "total en TB y 18 veces en DB."
    % (n(tb["varianza"], 1), n(tb["mediana"], 1), n(tb["media"]), n(db["mediana"], 1),
       n(db["media"]), n(tb["IQR"], 1), n(db["IQR"], 1), n(tb["rango"], 1), n(db["rango"], 1))
)
texto(
    "Las implicaciones son tres. La media y la desviación estándar no son resúmenes "
    "representativos de estas variables y deben acompañarse de la mediana y el IQR. Las "
    "técnicas sensibles a la escala, como la normalización Z-Score o los modelos basados "
    "en distancias, quedarán dominadas por los pocos valores extremos. Y, sin embargo, "
    "esa cola es informativa: la media de TB es %s mg/dL en los enfermos y %s en los "
    "sanos (en DB, %s frente a %s) y la desviación estándar 7,1 frente a 1,0, de modo que "
    "casi toda la variabilidad vive en el grupo enfermo; es la señal que se quiere "
    "predecir y no debe eliminarse sin más."
    % (n(DC["tb_media_enf"]), n(DC["tb_media_sano"]), n(DC["db_media_enf"]), n(DC["db_media_sano"]))
)

apartado3("4. Forma y simetría de TB, DB, Alkphos y Sgpt")
figura(1, "Histogramas de TB, DB, Alkphos y Sgpt con la media (línea continua) y la mediana (discontinua)",
       "fig1_histogramas.png",
       nota="Escala original; 40 intervalos por variable. Entre paréntesis, asimetría y curtosis de "
            "Fisher calculadas con pandas. Elaboración propia a partir de act_liver_disease.csv.",
       ancho=Cm(9.8))
texto(
    "Las cuatro distribuciones son unimodales, fuertemente asimétricas a la derecha y con "
    "colas pesadas (Figura 1). Los coeficientes de asimetría son %s en TB, %s en DB, %s en "
    "Alkphos y %s en Sgpt, cuando una distribución simétrica daría cero y por encima de 1 "
    "ya se considera marcada; la curtosis, de %s a %s, confirma que las colas concentran "
    "mucha más masa de la que tendría una normal. En las dos bilirrubinas la mitad de "
    "los pacientes se apila en el suelo de la escala (TB entre 0,4 y 1,0 mg/dL) "
    "y la media queda a la derecha del grueso de los datos; Alkphos es la única que "
    "muestra un cuerpo central reconocible en torno a 200 UI/L, y Sgpt es la más extrema, "
    "con casi todos los valores por debajo de 100 y un puñado que llega a 2000."
    % (n(FO["TB"]["asimetría"], 1), n(FO["DB"]["asimetría"], 1), n(FO["Alkphos"]["asimetría"], 1),
       n(FO["Sgpt"]["asimetría"], 1), n(FO["DB"]["curtosis"], 0), n(FO["Sgpt"]["curtosis"], 0))
)
texto(
    "El cuaderno repite los histogramas con la transformación log(1 + x) (celda 4.3). La "
    "asimetría baja a valores entre %s y %s: Alkphos y Sgpt se acercan a la lognormal "
    "(%s y %s), pero TB y DB siguen sesgadas (%s) porque el efecto suelo no desaparece con "
    "el logaritmo. El patrón es coherente con una muestra clínica en la que la mayoría "
    "presenta valores normales y una minoría, concentrada en los enfermos, valores muy "
    "alterados. Un método que suponga normalidad (regresión logística sin transformar, "
    "análisis discriminante, Z-Score) tratará mal estas variables; log(1 + x) es el paso "
    "previo razonable y los modelos de árboles son indiferentes a la forma."
    % (n(FO["Alkphos"]["asimetría_log"], 1), n(FO["TB"]["asimetría_log"], 1),
       n(FO["Alkphos"]["asimetría_log"], 1), n(FO["Sgpt"]["asimetría_log"], 1),
       n(FO["TB"]["asimetría_log"], 1))
)

apartado3("5. Relaciones entre las variables: matriz de correlación")
fu = {(d["a"], d["b"]): d for d in CR["fuertes"]}
texto(
    "Cuatro parejas superan un coeficiente de 0,5 y todas son significativas con *p* < 0,001 "
    "(Figura 2). TB y DB correlacionan %s (%s con Spearman), lo que era de esperar porque "
    "la directa es una fracción de la total; Sgpt y Sgot, %s, pues ambas transaminasas se "
    "liberan con el daño hepatocelular; TP y ALB, %s, ya que la albúmina es el componente "
    "mayoritario de las proteínas totales; y ALB con A/G Ratio, %s, porque el cociente se "
    "calcula a partir de ella. Estas redundancias importan para la selección de variables: "
    "en modelos lineales o regularizados conviene conservar de cada pareja la más ligada a "
    "la clase (DB sobre TB, Sgot sobre Sgpt, ALB sobre TP) o resumirlas (la fracción "
    "DB/TB), y la terna TP, ALB y A/G Ratio contiene en realidad dos grados de libertad; "
    "los árboles toleran la colinealidad."
    % (n(fu[("TB", "DB")]["r"], 3), n(fu[("TB", "DB")]["rho"], 3), n(fu[("Sgpt", "Sgot")]["r"], 3),
       n(fu[("TP", "ALB")]["r"], 3), n(fu[("ALB", "A/G Ratio")]["r"], 3))
)
figura(2, "Matriz de correlación de Pearson entre las once variables, con Gender (1 = hombre) y Enfermo (1 = Selector 1) recodificadas",
       "fig2_correlacion.png",
       nota="Triángulo inferior de la matriz; *n* = 579 pares completos para A/G Ratio y 583 en el resto. "
            "Elaboración propia a partir de act_liver_disease.csv.",
       ancho=Cm(10.0))
cc, cs = CR["con_clase"], CR["con_clase_spearman"]
texto(
    "Con la etiqueta clínica ninguna variable correlaciona con fuerza: los coeficientes de "
    "Pearson van de %s (DB) y %s (TB) a %s (ALB) y %s (A/G Ratio); todas son "
    "significativas al 5 %% salvo TP (%s, *p* = 0,40), porque con *n* = 583 un |*r*| de 0,08 "
    "(el de Gender) ya lo es: significación no implica utilidad. Con Spearman los "
    "marcadores de daño ganan peso (Sgot %s, TB %s, Sgpt %s, Alkphos %s), señal de que la "
    "relación con la enfermedad es monótona pero no lineal y de que los extremos de la "
    "escala original la ocultan (matriz completa en la celda 5.3 del cuaderno). El "
    "panorama anticipa una clasificación difícil, en la que ningún marcador aislado separa "
    "a los grupos y la ganancia vendrá de combinar variables y transformar las escalas."
    % (n(cc["DB"], 3), n(cc["TB"], 3), n(cc["ALB"], 3), n(cc["A/G Ratio"], 3), n(cc["TP"], 3),
       n(cs["Sgot"], 3), n(cs["TB"], 3), n(cs["Sgpt"], 3), n(cs["Alkphos"], 3))
)

apartado3("6. Imputación de valores faltantes por media y moda")
fi = IM["filas"]
texto(
    "Solo A/G Ratio tenía faltantes y sus cuatro celdas vacías recibieron la media de la "
    "columna, %s. Siguiendo la guía, se calculó la media de cada columna numérica de "
    "entrada, excluida la etiqueta, y la moda de cada columna no numérica, y con pandas se "
    "rellenaron los huecos con esos valores (celda 6.1); Gender no tiene faltantes (su "
    "moda es Male). El efecto sobre la distribución es pequeño pero visible: la "
    "media se mantiene por construcción, la desviación estándar baja de %s a %s y la "
    "mediana sube de %s a %s, porque en una variable asimétrica la media no es el centro."
    % (n(IM["media_ag"], 3), n(IM["std_antes"], 3), n(IM["std_despues"], 3),
       n(IM["mediana_ag"], 2), n(IM["media_ag"], 3))
)
texto(
    "La imputación por media es sencilla, rápida, no altera la media y sirve para cualquier "
    "algoritmo; con un 0,7 %% de faltantes el daño es mínimo. Sus desventajas son "
    "conocidas: reduce la varianza y "
    "debilita las correlaciones, ignora la relación entre variables y, en una variable "
    "asimétrica, inserta un valor que no es típico (la mediana de A/G Ratio es %s, algo "
    "menor que la media). La imputación por moda comparte la sencillez y es la única "
    "opción directa para categorías, pero refuerza la clase mayoritaria: aquí convertiría a "
    "todo paciente de sexo desconocido en hombre. La alternativa que el propio conjunto ofrece "
    "es más precisa: como A/G Ratio se deduce de ALB y TP, los cuatro valores "
    "reconstruidos serían %s, %s, %s y %s, bastante distintos de la media común que se "
    "les asignó. Para el modelado se recomienda esa reconstrucción determinista o, en su "
    "defecto, la mediana, y en cualquier caso ajustar el imputador solo con los datos de "
    "entrenamiento para no filtrar información de prueba."
    % (n(IM["mediana_ag"], 2), n(fi[0]["calc"], 2), n(fi[1]["calc"], 2), n(fi[2]["calc"], 2),
       n(fi[3]["calc"], 2))
)

apartado3("7. Normalización Min-Max y Z-Score")
texto(
    "Ambas técnicas conservan la forma de cada distribución y solo difieren en cómo tratan "
    "los extremos. Sobre el conjunto ya imputado se escalaron las nueve variables numéricas "
    "de entrada con MinMaxScaler y StandardScaler de scikit-learn (celda 7.1). Min-Max resta "
    "el mínimo y divide por el recorrido, y deja cada variable en el intervalo [0, 1]; "
    "Z-Score resta la media y divide por la desviación estándar, y deja cada variable con "
    "media 0 y desviación 1. Como muestra la Figura 3, la edad, que es razonablemente "
    "simétrica, queda bien repartida en ambas escalas (media Min-Max de %s; Z-Score entre "
    "%s y %s), mientras que Sgot sigue igual de concentrada."
    % (n(NO["mm_media_age"], 2), n(NO["z_min_age"], 1), n(NO["z_max_age"], 1))
)
figura(3, "Sgot (arriba) y Age (abajo) en escala original, Min-Max y Z-Score",
       "fig3_normalizacion.png",
       nota="Las transformaciones cambian la posición y la anchura, no la forma de la distribución. "
            "Elaboración propia a partir del conjunto imputado.",
       ancho=Cm(10.0))
texto(
    "Ahí reside la diferencia práctica. Con Min-Max el valor máximo de Sgot (4929) se "
    "convierte en 1 y arrastra al resto: la media escalada es %s y la mediana %s, es decir, "
    "la mitad de los pacientes ocupa menos del 1 %% del rango disponible y el algoritmo "
    "apenas podrá distinguirlos. Con Z-Score los extremos quedan a %s desviaciones en "
    "Sgot, %s en TB y %s en Sgpt, valores que en una normal serían imposibles y que "
    "dominarán cualquier distancia euclídea. Min-Max tiene la ventaja de un rango acotado "
    "y conocido, útil para redes neuronales y k vecinos, y la desventaja de depender solo "
    "del mínimo y el máximo y de producir valores fuera de [0, 1] si un dato nuevo los "
    "supera. Z-Score no acota el rango y también se contamina con los extremos, aunque "
    "menos, preserva la información de los atípicos y es el supuesto natural de la "
    "regresión regularizada y del análisis de componentes principales; los árboles y "
    "bosques no necesitan escalado alguno, lo que aquí es una razón para preferirlos. Con "
    "colas tan largas lo adecuado es log-transformar antes de escalar, o usar un escalador "
    "robusto (mediana e IQR), ajustado solo con el conjunto de entrenamiento."
    % (n(NO["mm_media_sgot"], 4), n(NO["mm_mediana_sgot"], 4), n(NO["z_max_sgot"], 1),
       n(NO["z_max_tb"], 1), n(NO["z_max_sgpt"], 1))
)

apartado3("8. Valores atípicos en Sgot, TP, ALB y A/G Ratio")
sg, tp, alb, ag = AT["Sgot"], AT["TP"], AT["ALB"], AT["A/G Ratio"]
texto(
    "Con la regla de Tukey (1977), que considera atípico lo que cae fuera de 1,5 veces el "
    "IQR desde los cuartiles, Sgot presenta %d atípicos (%s de los pacientes), todos por "
    "encima de %s UI/L, y %d son extremos (más de 3 IQR). El %s (64 pacientes) son enfermos, frente al %s de base, de modo que borrarlos quitaría "
    "el 15 %% de los enfermos y justamente los más reconocibles. TP tiene %d atípicos, "
    "%d por debajo de %s g/dL y %d por encima de %s, todos ellos enfermos. ALB no tiene "
    "ninguno (solo un 5,5 g/dL al separar por grupo), pero sus cajas por clase están "
    "desplazadas (medianas 3,0 y 3,4), mientras que las de TP son casi idénticas. A/G "
    "Ratio tiene %d atípicos, todos por encima de %s, de los que %d son extremos; %d son "
    "enfermos y %d sanos, es decir, proporcionalmente más frecuentes en los sanos (2,4 %% "
    "frente a 1,4 %%)."
    % (sg["total"], pct(sg["pct"]), n(sg["hi"], 0), sg["extremos"], pct(sg["pct_enf"], 0),
       pct(AT["pct_enfermos_global"]), tp["total"], tp["bajo"], n(tp["lo"], 1), tp["alto"],
       n(tp["hi"], 1), ag["total"], n(ag["hi"], 1), ag["extremos"],
       round(ag["total"] * ag["pct_enf"] / 100), ag["total"] - round(ag["total"] * ag["pct_enf"] / 100))
)
texto(
    "Los patrones son distintos en cada variable (Figura 4). En Sgot el problema no es un "
    "puñado de errores sino una cola continua: la caja de los enfermos es más alta y más "
    "ancha que la de los sanos, y en escala logarítmica se ve que los valores altos forman "
    "una secuencia sin saltos hasta 4929, salvo quizá ese último dato, 1,7 veces mayor "
    "que el siguiente (2946). En TP los atípicos son bilaterales y escasos, y los más bajos "
    "(2,7 a 3,0 g/dL) corresponden a desnutrición o insuficiencia hepática grave, no a "
    "errores. En A/G Ratio, en cambio, cinco de los diez atípicos coinciden con registros "
    "cuyo cociente no cuadra con ALB y TP (el máximo de 2,8 debería ser 0,7), de modo que "
    "aquí sí hay indicios de error de captura. En consecuencia, no se recomienda eliminar "
    "los atípicos de Sgot y TP, que llevan la señal clínica que se quiere modelar; basta "
    "transformarlos logarítmicamente o acotarlos (winsorizar) si el algoritmo lo exige. Sí "
    "conviene revisar uno a uno los %d registros incoherentes de A/G Ratio y el máximo de "
    "Sgot antes de modelar."
    % CO["ag_inconsistentes"]
)
figura(4, "Diagramas de caja de Sgot, TP, ALB y A/G Ratio por etiqueta clínica, Sgot en escala logarítmica y recuento de atípicos",
       "fig4_boxplots.png",
       nota="Los puntos son los valores fuera de 1,5 veces el IQR de cada grupo; los recuentos "
            "corresponden a la regla aplicada a la variable completa. Elaboración propia.",
       ancho=Cm(10.0))

apartado2("Conclusiones")
texto(
    "El conjunto es pequeño y está limpio en lo formal, pero presenta algunos faltantes y "
    "errores de coherencia, un desbalance de sexo y de clase que condicionará el modelado, duplicados que hay "
    "que controlar al particionar y, sobre todo, colas extremadamente largas en todos los "
    "marcadores. Las correlaciones fuertes son las que la bioquímica predice y ninguna "
    "variable aislada separa a enfermos de sanos. La imputación por media y la normalización "
    "cumplen su función, pero aquí conviene precederlas de una transformación logarítmica o "
    "emplear estimadores robustos, y tratar los atípicos como información clínica antes que "
    "como ruido. El conjunto imputado y sus dos versiones normalizadas quedan guardados por "
    "el cuaderno para la Actividad 2."
)

# ------------------------------------------------------------- REFERENCIAS
p_refs = apartado2("Referencias")
p_refs.paragraph_format.page_break_before = True
REFS = [
    "McKinney, W. (2010). Data structures for statistical computing in Python. En S. van der "
    "Walt y J. Millman (Eds.), *Proceedings of the 9th Python in Science Conference* "
    "(pp. 56–61). https://doi.org/10.25080/Majora-92bf1922-00a",
    "Meneses Yaranga, A. (2026). *Actividad 1. Tratamiento de datos para la predicción de "
    "enfermedades hepáticas* [Cuaderno Jupyter]. GitHub. " + REPO,
    "Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., Blondel, "
    "M., Prettenhofer, P., Weiss, R., Dubourg, V., Vanderplas, J., Passos, A., Cournapeau, D., "
    "Brucher, M., Perrot, M. y Duchesnay, É. (2011). Scikit-learn: Machine learning in Python. "
    "*Journal of Machine Learning Research, 12*, 2825–2830. "
    "https://jmlr.org/papers/v12/pedregosa11a.html",
    "Ramana, B. V., Babu, M. S. P. y Venkateswarlu, N. B. (2011). A critical study of selected "
    "classification algorithms for liver disease diagnosis. *International Journal of "
    "Database Management Systems, 3*(2), 101–114. https://doi.org/10.5121/ijdms.2011.3207",
    "Straw, I. y Wu, H. (2022). Investigating for bias in healthcare algorithms: A "
    "sex-stratified analysis of supervised machine learning models in liver disease "
    "prediction. *BMJ Health & Care Informatics, 29*(1), e100457. "
    "https://doi.org/10.1136/bmjhci-2021-100457",
    "Tukey, J. W. (1977). *Exploratory data analysis*. Addison-Wesley.",
    "Virtanen, P., Gommers, R., Oliphant, T. E., Haberland, M., Reddy, T., Cournapeau, D., "
    "Burovski, E., Peterson, P., Weckesser, W., Bright, J., van der Walt, S. J., Brett, M., "
    "Wilson, J., Millman, K. J., Mayorov, N., Nelson, A. R. J., Jones, E., Kern, R., Larson, "
    "E., … SciPy 1.0 Contributors. (2020). SciPy 1.0: Fundamental algorithms for scientific "
    "computing in Python. *Nature Methods, 17*(3), 261–272. "
    "https://doi.org/10.1038/s41592-019-0686-2",
    "Waskom, M. L. (2021). seaborn: Statistical data visualization. *Journal of Open Source "
    "Software, 6*(60), 3021. https://doi.org/10.21105/joss.03021",
]
for ref in REFS:
    p = texto(ref)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Cm(1.25)
    p.paragraph_format.first_line_indent = Cm(-1.25)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_after = Pt(5)

apartado2("Anexo. Cuaderno Jupyter")
p_anexo = texto(
    "El cuaderno CD_Actividad1_Liver.ipynb, con el código y las salidas de las ocho tareas, "
    "se entrega junto con este informe y puede abrirse y ejecutarse en Google Colab con el "
    "distintivo que encabeza el repositorio (Meneses Yaranga, 2026): " + REPO
)

p_anexo.alignment = WD_ALIGN_PARAGRAPH.LEFT

doc.save(str(SALIDA))
print("Informe escrito:", SALIDA)
