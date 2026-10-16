# -*- coding: utf-8 -*-
"""Infraestructura mínima para construir el notebook celda a celda.

El cuaderno se genera por código (y no editando el .ipynb a mano) para que sea
reproducible y fácil de versionar: cada celda se declara como una cadena de
texto en `nb_celdas.py` y este módulo las ensambla en el formato nbformat v4.
"""

import json

CELDAS = []


def md(texto):
    """Añade una celda Markdown al cuaderno."""
    CELDAS.append({
        "cell_type": "markdown",
        "metadata": {},
        "source": texto.strip("\n").splitlines(keepends=True),
    })


def code(texto):
    """Añade una celda de código Python al cuaderno."""
    CELDAS.append({
        "cell_type": "code",
        "execution_count": None,
        "metadata": {},
        "outputs": [],
        "source": texto.strip("\n").splitlines(keepends=True),
    })


def escribir(ruta):
    """Serializa las celdas acumuladas como un notebook nbformat v4."""
    cuaderno = {
        "cells": CELDAS,
        "metadata": {
            "kernelspec": {
                "display_name": "Python 3",
                "language": "python",
                "name": "python3",
            },
            "language_info": {"name": "python", "version": "3.14.0"},
        },
        "nbformat": 4,
        "nbformat_minor": 5,
    }
    with open(ruta, "w", encoding="utf-8") as fh:
        json.dump(cuaderno, fh, ensure_ascii=False, indent=1)
    return len(CELDAS)
