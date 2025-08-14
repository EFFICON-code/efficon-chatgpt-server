from flask import Flask, request, jsonify
import requests, os, traceback

app = Flask(__name__)

API_KEY     = os.getenv("OPENAI_API_KEY", "")
PROJECT_ID  = os.getenv("OPENAI_PROJECT_ID", "")
OPENAI_URL  = "https://api.openai.com/v1/chat/completions"
MODEL       = os.getenv("MODEL", "gpt-4o")

def openai_call(messages, max_tokens=2000, temperature=0.3, timeout_s=180):
    if not API_KEY:
        return {
            "choices": [],
            "error": "API_KEY no configurada",
            "code": 10001,
            "ok": False,
            "text": ""
        }

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    if PROJECT_ID:
        headers["OpenAI-Project"] = PROJECT_ID

    payload = {
        "model": MODEL,
        "messages": messages,
        "max_tokens": max_tokens,
        "temperature": temperature,
        "stream": False
    }

    try:
        resp = requests.post(OPENAI_URL, headers=headers, json=payload, timeout=timeout_s)
        if resp.status_code != 200:
            return {
                "choices": [],
                "error": f"OpenAI error HTTP {resp.status_code}",
                "code": resp.status_code,
                "ok": False,
                "text": resp.text[:800]
            }
        return resp.json()
    except Exception as e:
        return {
            "choices": [],
            "error": str(e),
            "code": 10002,
            "ok": False,
            "text": "",
            "_trace": traceback.format_exc()[:800]
        }

@app.route("/")
def home():
    return "✅ Servidor de ChatGPT activo en Railway."

@app.route("/chatgpt", methods=["POST"])
def chatgpt():
    try:
        data = request.get_json(force=True, silent=True) or {}
        userPrompt = (data.get("prompt") or "").strip()

        if not userPrompt:
            return jsonify({
                "choices": [],
                "error": "Prompt vacío",
                "code": 10003,
                "ok": False,
                "text": ""
            }), 200

        messages = [{"role": "user", "content": userPrompt}]
        result = openai_call(messages)
        if "choices" not in result:
            result["choices"] = []
        if "ok" not in result:
            result["ok"] = False
        if "code" not in result:
            result["code"] = 10004

        return jsonify(result), 200
    except Exception as e:
        return jsonify({
            "choices": [],
            "error": "Error inesperado",
            "code": 10005,
            "ok": False,
            "text": "",
            "_trace": traceback.format_exc()[:800]
        }), 200
