# Actividad 1. Tratamiento de datos para la predicción de enfermedades hepáticas

[![Abrir en Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/amenesesy/actividad-cd-liver/blob/main/CD_Actividad1_Liver.ipynb)

Certificado de Educación Continua en Ciencia de Datos (código C0024), MIU City University Miami.
Autor: Abel Meneses Yaranga. Octubre de 2026.

Análisis descriptivo y exploratorio, gestión de valores faltantes, normalización e
identificación de valores atípicos sobre `act_liver_disease.csv`, un derivado del
Indian Liver Patient Dataset usado por Straw y Wu (2022) en su estudio de sesgo por
sexo en modelos de predicción de enfermedad hepática.

## Contenido

| Archivo | Descripción |
|---|---|
| `CD_Actividad1_Liver.ipynb` | Cuaderno ejecutado con las ocho tareas de la guía (abre en Colab con el distintivo de arriba). |
| `act_liver_disease.csv` | Conjunto de datos de la actividad (583 pacientes, 11 variables). |
| `Actividad1_CienciaDatos_Meneses.pdf` | Informe de la actividad (10 páginas). |
| `figuras/` | Figuras generadas por el cuaderno. |
| `resultados.json` | Cifras que cita el informe, escritas por el cuaderno. |
| `liver_imputado.csv`, `liver_minmax.csv`, `liver_zscore.csv` | Conjunto tratado y sus dos versiones normalizadas. |
| `nb_base.py`, `nb_celdas.py`, `build_notebook.py` | Generan el cuaderno por código. |
| `build_informe.py`, `exportar_pdf.ps1` | Construyen el informe sobre la plantilla del curso y lo exportan a PDF. |

## Reproducir

```bash
python build_notebook.py
jupyter nbconvert --to notebook --execute --inplace CD_Actividad1_Liver.ipynb
python build_informe.py
```

El cuaderno detecta si corre en Google Colab, instala lo que falte y descarga el CSV
desde este repositorio si no está en el directorio de trabajo.

## Referencia del conjunto de datos

Straw, I. y Wu, H. (2022). Investigating for bias in healthcare algorithms: A sex-stratified
analysis of supervised machine learning models in liver disease prediction.
*BMJ Health & Care Informatics, 29*(1), e100457. https://doi.org/10.1136/bmjhci-2021-100457
