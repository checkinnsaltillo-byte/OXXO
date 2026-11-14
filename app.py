#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dashboard del directorio RME (servicio de Cloud Run).

  GET  /               dashboard
  GET  /api/datos      empresas + resumen de la última corrida + estado
  GET  /api/estado     avance de la extracción
  POST /api/ejecutar   lanza el Cloud Run Job que hace la extracción
  GET  /descargar.xlsx Excel de la última corrida
"""
import datetime as dt
import json
import os

import google.auth
from flask import Flask, Response, jsonify, send_from_directory
from google.auth.transport.requests import AuthorizedSession
from google.cloud import storage

BUCKET = os.environ["BUCKET"]
PROYECTO = os.environ["PROYECTO"]
REGION = os.environ.get("REGION", "us-central1")
JOB = os.environ.get("JOB", "oxxo-scraper")
MAX_MIN_CORRIDA = 60  # si una corrida lleva más de esto, se considera atorada

app = Flask(__name__, static_folder="static")
_cubeta = None


def cubeta():
    global _cubeta
    if _cubeta is None:
        _cubeta = storage.Client().bucket(BUCKET)
    return _cubeta


def leer(nombre, defecto):
    b = cubeta().blob(nombre)
    try:
        return json.loads(b.download_as_bytes())
    except Exception:
        return defecto


def ahora():
    return dt.datetime.utcnow()


def estado_actual():
    e = leer("estado.json", {"estado": "sin corridas"})
    if e.get("estado") in ("corriendo", "en cola") and e.get("inicio"):
        ini = dt.datetime.strptime(e["inicio"], "%Y-%m-%dT%H:%M:%SZ")
        if (ahora() - ini).total_seconds() > MAX_MIN_CORRIDA * 60:
            e["estado"] = "error"
            e["error"] = "La corrida no terminó (más de %d min)." % MAX_MIN_CORRIDA
    return e


@app.get("/")
def inicio():
    r = send_from_directory(app.static_folder, "index.html")
    r.headers["Cache-Control"] = "no-cache"
    return r


@app.get("/api/datos")
def datos():
    base = leer("directorio_rme.json", {})
    corrida = leer("ultima_corrida.json", {})
    empresas = []
    for k, r in base.items():
        if r.get("es_empresa"):
            r = dict(r)
            r["clave"] = k
            for c in ("respondio", "paginas", "url_final", "es_empresa", "menciones_sector"):
                r.pop(c, None)
            empresas.append(r)
    r = jsonify({"empresas": empresas, "corrida": corrida, "estado": estado_actual()})
    r.headers["Cache-Control"] = "no-store"
    return r


@app.get("/api/estado")
def estado():
    r = jsonify(estado_actual())
    r.headers["Cache-Control"] = "no-store"
    return r


@app.post("/api/ejecutar")
def ejecutar():
    e = estado_actual()
    if e.get("estado") in ("corriendo", "en cola"):
        return jsonify({"ok": False, "motivo": "Ya hay una extracción en curso.", "estado": e}), 409
    nuevo = {"estado": "en cola", "inicio": ahora().isoformat(timespec="seconds") + "Z", "etapa": "Iniciando", "hecho": 0, "total": 0}
    b = cubeta().blob("estado.json")
    b.cache_control = "no-store"
    b.upload_from_string(json.dumps(nuevo), content_type="application/json")
    cred, _ = google.auth.default(scopes=["https://www.googleapis.com/auth/cloud-platform"])
    resp = AuthorizedSession(cred).post(
        "https://run.googleapis.com/v2/projects/%s/locations/%s/jobs/%s:run" % (PROYECTO, REGION, JOB), json={}, timeout=30)
    if resp.status_code >= 300:
        nuevo.update({"estado": "error", "error": "No se pudo iniciar: %s" % resp.text[:200]})
        b.upload_from_string(json.dumps(nuevo), content_type="application/json")
        return jsonify({"ok": False, "motivo": nuevo["error"], "estado": nuevo}), 502
    return jsonify({"ok": True, "estado": nuevo})


@app.get("/descargar.xlsx")
def descargar():
    b = cubeta().blob("directorio_rme.xlsx")
    if not b.exists():
        return "Aún no hay Excel", 404
    return Response(b.download_as_bytes(), headers={
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="directorio_rme.xlsx"', "Cache-Control": "no-store"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
