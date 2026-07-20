/* ============================================================
   Cámara → OneDrive
   - Autenticación: MSAL.js (OAuth2 / Microsoft Identity Platform)
   - Cámara: getUserMedia
   - Subida: Microsoft Graph API (PUT .../content)
   - Cola offline: localStorage (reintenta solas al recuperar conexión)
   ============================================================ */

const cfg = window.APP_CONFIG;

const els = {
  video: document.getElementById("video"),
  canvas: document.getElementById("canvas"),
  hint: document.getElementById("hint"),
  viewfinder: document.getElementById("viewfinder"),
  previewOverlay: document.getElementById("previewOverlay"),
  previewImg: document.getElementById("previewImg"),
  shutterBtn: document.getElementById("shutterBtn"),
  switchCamBtn: document.getElementById("switchCamBtn"),
  logToggleBtn: document.getElementById("logToggleBtn"),
  captureControls: document.getElementById("captureControls"),
  previewActions: document.getElementById("previewActions"),
  discardBtn: document.getElementById("discardBtn"),
  uploadBtn: document.getElementById("uploadBtn"),
  statusMsg: document.getElementById("statusMsg"),
  statusCount: document.getElementById("statusCount"),
  logPanel: document.getElementById("logPanel"),
  authPanel: document.getElementById("authPanel"),
  loginBtn: document.getElementById("loginBtn"),
  ledAuth: document.getElementById("led-auth"),
  ledCam: document.getElementById("led-cam"),
  ledSync: document.getElementById("led-sync"),
};

let currentStream = null;
let facingMode = "environment";
let capturedBlob = null;
let msalInstance = null;
let account = null;
const QUEUE_KEY = "camera-onedrive-queue-v1";

// ---------- utilidades de estado visual ----------
function setLed(el, state) {
  el.classList.remove("ok", "warn");
  if (state === "ok") el.classList.add("ok");
  else if (state === "warn") el.classList.add("warn");
}
function setStatus(msg) {
  els.statusMsg.textContent = msg;
}
function refreshCount() {
  const q = getQueue();
  els.statusCount.textContent = q.length ? `${q.length} pendiente(s)` : "";
}

// ---------- cola local (para cuando no hay red o falla la subida) ----------
function getQueue() {
  try { return JSON.parse(localStorage.getItem(QUEUE_KEY) || "[]"); }
  catch { return []; }
}
function saveQueue(q) {
  localStorage.setItem(QUEUE_KEY, JSON.stringify(q));
  refreshCount();
}
function addLogRow(name, state) {
  const row = document.createElement("div");
  row.className = `log-row ${state}`;
  row.id = `log-${name}`;
  row.innerHTML = `<span class="dot"></span><span class="name">${name}</span>`;
  els.logPanel.prepend(row);
}
function updateLogRow(name, state) {
  const row = document.getElementById(`log-${name}`);
  if (row) row.className = `log-row ${state}`;
}

// ---------- MSAL / autenticación ----------
function initAuth() {
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: cfg.authority,
      redirectUri: cfg.redirectUri,
    },
    cache: { cacheLocation: "localStorage" },
  });

  msalInstance.handleRedirectPromise().then((resp) => {
    if (resp && resp.account) account = resp.account;
    else {
      const accs = msalInstance.getAllAccounts();
      if (accs.length) account = accs[0];
    }
    onAuthChanged();
  }).catch((e) => {
    console.error(e);
    setStatus("Error de autenticación");
  });

  els.loginBtn.addEventListener("click", () => {
    if (cfg.clientId.includes("PEGA_AQUI")) {
      alert("Falta configurar el Client ID en config.js. Revisa el README.");
      return;
    }
    msalInstance.loginRedirect({ scopes: cfg.scopes });
  });
}

function onAuthChanged() {
  if (account) {
    setLed(els.ledAuth, "ok");
    els.authPanel.classList.remove("show");
    flushQueue();
  } else {
    setLed(els.ledAuth, "");
    els.authPanel.classList.add("show");
  }
}

async function getAccessToken() {
  const req = { scopes: cfg.scopes, account };
  try {
    const res = await msalInstance.acquireTokenSilent(req);
    return res.accessToken;
  } catch (e) {
    const res = await msalInstance.acquireTokenPopup(req);
    return res.accessToken;
  }
}

// ---------- cámara ----------
async function startCamera() {
  stopCamera();
  els.hint.textContent = "Iniciando cámara…";
  try {
    currentStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: facingMode } },
      audio: false,
    });
    els.video.srcObject = currentStream;
    els.hint.style.display = "none";
    els.shutterBtn.disabled = false;
    setLed(els.ledCam, "ok");
  } catch (e) {
    console.error(e);
    els.hint.textContent = "No se pudo acceder a la cámara";
    setLed(els.ledCam, "");
    setStatus("Permiso de cámara denegado o no disponible");
  }
}
function stopCamera() {
  if (currentStream) {
    currentStream.getTracks().forEach((t) => t.stop());
    currentStream = null;
  }
}

function capturePhoto() {
  const video = els.video;
  const canvas = els.canvas;
  canvas.width = video.videoWidth;
  canvas.height = video.videoHeight;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
  canvas.toBlob((blob) => {
    capturedBlob = blob;
    els.previewImg.src = URL.createObjectURL(blob);
    showPreview(true);
  }, "image/jpeg", 0.92);
}

function showPreview(show) {
  els.previewOverlay.classList.toggle("show", show);
  els.previewActions.classList.toggle("show", show);
  els.captureControls.style.display = show ? "none" : "flex";
}

// ---------- subida a OneDrive (Microsoft Graph) ----------
function filenameFor() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `foto_${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}.jpg`;
}

async function uploadBlob(blob, filename) {
  const token = await getAccessToken();
  const path = `${cfg.folderPath}/${filename}`;
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(path).replace(/%2F/g, "/")}:/content`;
  const res = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "image/jpeg",
    },
    body: blob,
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph API ${res.status}: ${text}`);
  }
  return res.json();
}

// Convierte un Blob a base64 para poder guardarlo en localStorage mientras no hay red.
function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = reject;
    r.readAsDataURL(blob);
  });
}
function base64ToBlob(base64) {
  return fetch(base64).then((r) => r.blob());
}

async function queuePhoto(blob) {
  const filename = filenameFor();
  const b64 = await blobToBase64(blob);
  const q = getQueue();
  q.push({ filename, b64 });
  saveQueue(q);
  addLogRow(filename, "pending");
  flushQueue();
}

let flushing = false;
async function flushQueue() {
  if (flushing || !account) return;
  flushing = true;
  setLed(els.ledSync, "warn");
  const q = getQueue();
  for (const item of [...q]) {
    try {
      setStatus(`Subiendo ${item.filename}…`);
      const blob = await base64ToBlob(item.b64);
      await uploadBlob(blob, item.filename);
      updateLogRow(item.filename, "done");
      const rest = getQueue().filter((x) => x.filename !== item.filename);
      saveQueue(rest);
      setStatus(`${item.filename} subida ✓`);
    } catch (e) {
      console.error(e);
      updateLogRow(item.filename, "error");
      setStatus(`Error al subir ${item.filename} — se reintentará`);
      break; // dejamos el resto en cola, reintenta más tarde
    }
  }
  setLed(els.ledSync, getQueue().length ? "warn" : "ok");
  flushing = false;
}

// ---------- eventos UI ----------
els.shutterBtn.addEventListener("click", capturePhoto);
els.discardBtn.addEventListener("click", () => {
  capturedBlob = null;
  showPreview(false);
});
els.uploadBtn.addEventListener("click", async () => {
  if (!capturedBlob) return;
  const blob = capturedBlob;
  showPreview(false);
  await queuePhoto(blob);
});
els.switchCamBtn.addEventListener("click", () => {
  facingMode = facingMode === "environment" ? "user" : "environment";
  startCamera();
});
els.logToggleBtn.addEventListener("click", () => {
  els.logPanel.classList.toggle("open");
});

window.addEventListener("online", flushQueue);

// ---------- arranque ----------
function boot() {
  initAuth();
  startCamera();
  refreshCount();
  getQueue().forEach((item) => addLogRow(item.filename, "pending"));

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch((e) => console.error("SW error", e));
  }
}
boot();
