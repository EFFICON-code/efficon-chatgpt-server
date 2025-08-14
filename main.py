import os
import json
import requests
from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=False)

@app.get("/healthz")
def healthz():
    return jsonify(status="ok"), 200

@app.post("/chatgpt")
def chatgpt():
    data = request.get_json(silent=True) or {}
    prompt = data.get("prompt", "")
    entidad = data.get("entidad", "")
    token = data.get("token", "")

    # Seguridad simple por variables (opcional)
    allowed_entidad = os.environ.get("EFFICON_ENTIDAD", "")
    allowed_token = os.environ.get("EFFICON_TOKEN", "")
    if (allowed_entidad and entidad != allowed_entidad) or (allowed_token and token != allowed_token):
        return jsonify(error="Auth failed"), 403

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return jsonify(error="OPENAI_API_KEY not configured"), 500

    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
payload = {
    "model": "gpt-4o-mini",   # <--- CAMBIA a este
    "messages": [{"role": "user", "content": prompt}],
    "temperature": 0.2
}

    try:
        r = requests.post("https://api.openai.com/v1/chat/completions",
                          headers=headers, json=payload, timeout=120)
        r.raise_for_status()
    except requests.exceptions.RequestException as e:
        status = getattr(e.response, "status_code", 500)
        return jsonify(error="OpenAI request failed",
                       details=str(e),
                       body=getattr(e.response, "text", "")), status

    out = r.json()
    text = ""
    if out.get("choices"):
        msg = out["choices"][0].get("message", {})
        text = msg.get("content", "")

    return jsonify(ok=True, entidad=entidad, answer=text), 200

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
import os, json
from flask import Flask, request, jsonify, send_from_directory, abort
from flask_cors import CORS
import requests

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

ENTITY_NAME = os.environ.get("EFFICON_ENTITY_NAME", "ENTIDAD-NO-SET")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")

MODULE_DIR = os.path.join(os.getcwd(), "modules", "pack")
MANIFEST_PATH = os.path.join(MODULE_DIR, "manifest.json")

@app.get("/healthz")
def healthz():
    return jsonify(status="ok", entity=ENTITY_NAME, manifest=os.path.isfile(MANIFEST_PATH)), 200

@app.get("/modules/manifest")
def get_manifest():
    if not os.path.isfile(MANIFEST_PATH):
        return jsonify(error="manifest not found"), 404
    with open(MANIFEST_PATH, "r", encoding="utf-8") as f:
        mf = json.load(f)
    mf["entity"] = ENTITY_NAME
    return jsonify(mf), 200

@app.get("/modules/download/<path:fname>")
def download_bas(fname):
    if not fname.lower().endswith(".bas"):
        abort(400)
    fullpath = os.path.join(MODULE_DIR, fname)
    if not os.path.isfile(fullpath):
        abort(404)
    return send_from_directory(MODULE_DIR, fname, as_attachment=True)
# --- RUTA /chatgpt con OpenAI ---
import os, requests
from flask import request, jsonify

@app.route("/chatgpt", methods=["POST"])
def chatgpt():
    data = request.get_json(silent=True) or {}
    prompt = data.get("prompt", "")
    if not prompt:
        return jsonify(error="prompt requerido"), 400

    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return jsonify(error="OPENAI_API_KEY not configured"), 500

    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    payload = {
        "model": "gpt-4.1-mini",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.2
    }

    try:
        r = requests.post("https://api.openai.com/v1/chat/completions",
                          headers=headers, json=payload, timeout=120)
        r.raise_for_status()
        out = r.json()
        text = out.get("choices", [{}])[0].get("message", {}).get("content", "").strip()
        return jsonify(ok=True,
                       entity=os.environ.get("EFFICON_ENTITY_NAME", ""),
                       answer=text), 200
    except requests.exceptions.RequestException as e:
        status = getattr(e.response, "status_code", 502)
        return jsonify(error="OpenAI request failed",
                       details=str(e),
                       body=getattr(e.response, "text", "")), status
