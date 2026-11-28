# -*- coding: utf-8 -*-
"""
aplicar_pqr_pie.py
Engancha en app.py los dos módulos nuevos (adjuntos + PIN en PQR, y pie de página legal).

  python aplicar_pqr_pie.py              -> aplica (hace respaldo app.py.antes_pqr_pie.bak la primera vez)
  python aplicar_pqr_pie.py --revertir   -> quita los bloques de app.py

Es seguro correrlo varias veces: no duplica los bloques. Conserva los saltos de línea del archivo.
"""
import os
import re
import shutil
import sys

APP = "app.py"
BAK = "app.py.antes_pqr_pie.bak"
BLOQUES = [
    ("procsis-pqr-seguro-v1", "pqr_adjuntos_pin", "_pqs1",
     "Adjuntos + verificación por PIN en PQR", "PQR SEGURO"),
    ("procsis-pie-legal-v1", "pie_pagina", "_pie1",
     "Pie de página legal + /terminos y /mapa-del-sitio", "PIE LEGAL"),
]


def bloque(marca, modulo, alias, desc, etiqueta, nl):
    lineas = [
        "# [%s] %s. Para quitarlo: borra este bloque o usa aplicar_pqr_pie.py --revertir" % (marca, desc),
        "try:",
        "    import %s as %s" % (modulo, alias),
        "    %s.instalar(globals())" % alias,
        "except Exception as %se:" % alias,
        "    print('%s no se pudo instalar (la app sigue igual):', repr(%se))" % (etiqueta, alias),
    ]
    return (nl.join(lineas) + nl + nl + nl).encode("utf-8")


def main():
    revertir = "--revertir" in sys.argv
    if not os.path.exists(APP):
        sys.exit("No encuentro app.py en esta carpeta. Ejecuta el script desde la raíz del proyecto.")
    datos = open(APP, "rb").read()
    nl = "\r\n" if b"\r\n" in datos else "\n"

    if revertir:
        n = 0
        for marca, modulo, alias, desc, etiqueta in BLOQUES:
            patron = re.compile(rb"# \[" + marca.encode() + rb"\].*?no se pudo instalar[^\r\n]*\r?\n(?:\r?\n){0,2}", re.S)
            datos, k = patron.subn(b"", datos)
            n += k
        open(APP, "wb").write(datos)
        print("Revertido: %d bloque(s) quitado(s) de app.py." % n)
        return

    if not os.path.exists(BAK):
        shutil.copyfile(APP, BAK)
        print("Respaldo creado:", BAK)
    ancla = b'if __name__ == "__main__":'
    pos = datos.rfind(ancla)
    if pos < 0:
        sys.exit('No encontré la línea if __name__ == "__main__": en app.py; no se modificó nada.')
    nuevos = b""
    for marca, modulo, alias, desc, etiqueta in BLOQUES:
        if ("[" + marca + "]").encode() in datos:
            print("Ya estaba aplicado:", marca)
            continue
        if not os.path.exists(modulo + ".py"):
            sys.exit("Falta el archivo %s.py en esta carpeta." % modulo)
        nuevos += bloque(marca, modulo, alias, desc, etiqueta, nl)
        print("Aplicado:", marca)
    if nuevos:
        datos = datos[:pos] + nuevos + datos[pos:]
        open(APP, "wb").write(datos)
    print("Listo. Reinicia la app y revisa /pqr/crear y el pie de página.")


if __name__ == "__main__":
    main()
