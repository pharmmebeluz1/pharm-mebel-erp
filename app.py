# -*- coding: utf-8 -*-
"""Mebel360° — Flask/PWA starter web application.

The uploaded package contained an Excel-module installer as app.py rather than a
Flask application. That installer is preserved as excel_module_installer.py.
This file provides a working web/PWA shell for Render, phone and desktop use.
"""
from __future__ import annotations

import os
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, jsonify, make_response, render_template, send_from_directory

BASE_DIR = Path(__file__).resolve().parent

app = Flask(__name__, static_folder="static", template_folder="templates")
app.config.update(
    SECRET_KEY=os.environ.get("MEBEL360_SECRET", "change-this-in-production"),
    JSON_AS_ASCII=False,
)


@app.get("/")
def home():
    return render_template("index.html", year=datetime.now().year)


@app.get("/api/health")
def health():
    return jsonify(
        status="ok",
        app="Mebel360°",
        time=datetime.now(timezone.utc).isoformat(),
    )


@app.get("/manifest.webmanifest")
def manifest():
    response = make_response(send_from_directory(app.static_folder, "manifest.webmanifest"))
    response.headers["Content-Type"] = "application/manifest+json"
    response.headers["Cache-Control"] = "no-cache"
    return response


@app.get("/service-worker.js")
def service_worker():
    response = make_response(send_from_directory(app.static_folder, "service-worker.js"))
    response.headers["Content-Type"] = "application/javascript; charset=utf-8"
    response.headers["Service-Worker-Allowed"] = "/"
    response.headers["Cache-Control"] = "no-cache"
    return response


@app.get("/offline.html")
def offline():
    return send_from_directory(app.static_folder, "offline.html")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "10000"))
    app.run(host="0.0.0.0", port=port, debug=os.environ.get("FLASK_DEBUG") == "1")
