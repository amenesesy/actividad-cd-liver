# -*- coding: utf-8 -*-
"""Ensambla el notebook de la actividad a partir de nb_celdas.py.

Uso:
    python build_notebook.py
"""

import nb_base
import nb_celdas

SALIDA = "CD_Actividad1_Liver.ipynb"

if __name__ == "__main__":
    nb_celdas.construir()
    n = nb_base.escribir(SALIDA)
    n_codigo = sum(1 for c in nb_base.CELDAS if c["cell_type"] == "code")
    print("Notebook generado: %s" % SALIDA)
    print("  Celdas totales   : %d" % n)
    print("  Celdas de código : %d" % n_codigo)
    print("  Celdas Markdown  : %d" % (n - n_codigo))
