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
        "model": "gpt-4.1-mini",  # TODO: cambia si quieres otro
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
