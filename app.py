"""
PedidosAuto - Bling Web
Servidor Flask com upload de planilha, execução do bot em thread e SSE para logs em tempo real.
"""

import io
import os
import queue
import threading
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from bot import BlingBot

# ── Diretórios ──────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
PRINTS_DIR = BASE_DIR / "prints"

for d in (UPLOAD_DIR, PRINTS_DIR):
    d.mkdir(exist_ok=True)

# ── Flask ───────────────────────────────────────────────────
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB

# ── Estado global (single-user) ─────────────────────────────
_log_queue: queue.Queue = queue.Queue()
_bot_instance: BlingBot | None = None
_bot_thread: threading.Thread | None = None
_resultado_filename: str | None = None
_resultado_bytes: bytes | None = None  # planilha de resultado em memória (sem persistir)


def _log_callback(msg: str):
    """Callback usado pelo bot — enfileira mensagens para o SSE."""
    _log_queue.put(msg)


# ── Rotas ───────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    global _bot_instance, _bot_thread, _resultado_filename, _resultado_bytes

    if _bot_thread and _bot_thread.is_alive():
        return jsonify({"error": "Bot já está em execução. Aguarde a conclusão."}), 409

    file = request.files.get("planilha")
    if not file or not file.filename or not file.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Envie um arquivo Excel (.xlsx)."}), 400

    # Salva upload (sanitiza nome para evitar path traversal e caracteres inválidos)
    nome_seguro = secure_filename(file.filename) or "planilha.xlsx"
    caminho = str(UPLOAD_DIR / nome_seguro)
    file.save(caminho)

    # Limpa fila de logs anteriores
    while not _log_queue.empty():
        try:
            _log_queue.get_nowait()
        except queue.Empty:
            break

    _resultado_filename = None
    _resultado_bytes = None

    # Campos opcionais (override do .env / padrões)
    email_form = (request.form.get("email") or "").strip() or None
    senha_form = request.form.get("senha") or None
    cliente_form = (request.form.get("cliente") or "").strip() or None
    loja_form = (request.form.get("loja") or "").strip() or None
    frete_form = (request.form.get("frete") or "").strip() or None

    try:
        _bot_instance = BlingBot(
            caminho_planilha=caminho,
            resultado_dir="",  # não usado — resultado é gerado em memória
            prints_dir=str(PRINTS_DIR),
            log=_log_callback,
            email=email_form,
            senha=senha_form,
            cliente=cliente_form,
            loja=loja_form,
            frete=frete_form,
        )
    except ValueError as e:
        return jsonify({"error": str(e)}), 400

    def _run():
        global _resultado_filename, _resultado_bytes
        _bot_instance.executar()
        if _bot_instance.resultado_bytes and _bot_instance.resultado_filename:
            _resultado_bytes = _bot_instance.resultado_bytes
            _resultado_filename = _bot_instance.resultado_filename

    _bot_thread = threading.Thread(target=_run, daemon=True)
    _bot_thread.start()

    return jsonify({"ok": True, "message": "Bot iniciado com sucesso."})


@app.route("/stream")
def stream():
    """SSE — envia logs do bot em tempo real."""

    def event_stream():
        while True:
            try:
                msg = _log_queue.get(timeout=30)
            except queue.Empty:
                yield "event: ping\ndata: keep-alive\n\n"
                continue

            # Envia a mensagem
            safe = msg.replace("\n", "\\n")
            yield f"data: {safe}\n\n"

            # Se o bot terminou, envia o nome do resultado e encerra
            if msg.strip() == "DONE":
                if _resultado_filename:
                    yield f"event: result\ndata: {_resultado_filename}\n\n"
                break

    return Response(event_stream(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})


@app.route("/cancelar", methods=["POST"])
def cancelar():
    if _bot_instance:
        _bot_instance.cancelar()
        return jsonify({"ok": True, "message": "Cancelamento solicitado."})
    return jsonify({"error": "Nenhum bot em execução."}), 404


@app.route("/download/<filename>")
def download(filename):
    """Entrega a planilha de resultado direto da memória — não persiste em disco."""
    if not _resultado_bytes or not _resultado_filename:
        return jsonify({"error": "Nenhum resultado disponível."}), 404

    nome_seguro = secure_filename(filename) or _resultado_filename
    return send_file(
        io.BytesIO(_resultado_bytes),
        as_attachment=True,
        download_name=nome_seguro,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Main ────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000, threaded=True)
