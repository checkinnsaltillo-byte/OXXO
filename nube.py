#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Corrida del scraper en Google Cloud (Cloud Run Job).

Baja la base del bucket, ejecuta la extracción publicando el avance en estado.json
y sube los resultados (directorio_rme.json, directorio_rme.xlsx, ultima_corrida.json).
"""
import datetime as dt
import json
import os
import time
import traceback

TRABAJO = "/tmp/rme"
os.makedirs(TRABAJO, exist_ok=True)
os.environ["RME_DIR"] = TRABAJO

from google.cloud import storage  # noqa: E402

import scraper_rme as s  # noqa: E402

BUCKET = os.environ["BUCKET"]
ARCHIVOS = ["directorio_rme.json", "serp_google.json"]
SALIDAS = ["directorio_rme.json", "directorio_rme.xlsx", "ultima_corrida.json"]


def ahora():
    return dt.datetime.utcnow().isoformat(timespec="seconds") + "Z"


def main():
    cubeta = storage.Client().bucket(BUCKET)
    estado = {"estado": "corriendo", "inicio": ahora(), "etapa": "Preparando", "hecho": 0, "total": 0}
    ultimo = [0.0]

    def publicar(forzar=False):
        if forzar or time.time() - ultimo[0] > 3:
            ultimo[0] = time.time()
            estado["actualizado"] = ahora()
            b = cubeta.blob("estado.json")
            b.cache_control = "no-store"
            b.upload_from_string(json.dumps(estado, ensure_ascii=False), content_type="application/json")

    def avisar(etapa, hecho=0, total=0):
        cambio = etapa != estado["etapa"]
        estado.update({"etapa": etapa, "hecho": hecho, "total": total})
        print("%s %s/%s" % (etapa, hecho, total), flush=True)
        publicar(cambio or hecho == total)

    s.AVISAR = avisar
    publicar(True)
    try:
        for nombre in ARCHIVOS:
            b = cubeta.blob(nombre)
            if b.exists():
                b.download_to_filename(os.path.join(TRABAJO, nombre))
        resumen = s.ejecutar("auto")
        for nombre in SALIDAS:
            b = cubeta.blob(nombre)
            b.cache_control = "no-store"
            b.upload_from_filename(os.path.join(TRABAJO, nombre))
        estado.update({"estado": "terminado", "fin": ahora(), "etapa": "Listo", "empresas": resumen["empresas"],
                       "por_estatus": resumen["por_estatus"], "motor": resumen["motor"]})
        publicar(True)
    except BaseException as e:  # incluye sys.exit por falta de claves
        traceback.print_exc()
        estado.update({"estado": "error", "fin": ahora(), "etapa": "Error", "error": str(e)[:300]})
        publicar(True)
        raise


if __name__ == "__main__":
    main()
