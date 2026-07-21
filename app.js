/* ============================================================
   Cámara → OneDrive
   - Autenticación: MSAL.js (OAuth2 / Microsoft Identity Platform)
   - Cámara: getUserMedia (fotos con canvas, videos con MediaRecorder)
   - Subida: Microsoft Graph API (PUT .../content)
   - Cola offline: IndexedDB (soporta archivos grandes como video)
   - Carpeta destino: configurable desde la app, se crea si no existe
   ============================================================ */

const cfg = window.APP_CONFIG;

const els = {
  video: document.getElementById("video"),
  canvas: document.getElementById("canvas"),
  hint: document.getElementById("hint"),
  viewfinder: document.getElementById("viewfinder"),
  previewOverlay: document.getElementById("previewOverlay"),
  previewImg: document.getElementById("previewImg"),
  previewVideo: document.getElementById("previewVideo"),
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
  recBadge: document.getElementById("recBadge"),
  recTime: document.getElementById("recTime"),
  modePhotoBtn: document.getElementById("modePhotoBtn"),
  modeVideoBtn: document.getElementById("modeVideoBtn"),
  folderLabel: document.getElementById("folderLabel"),
  folderBtnOpen: document.getElementById("folderBtnOpen"),
  folderPanel: document.getElementById("folderPanel"),
  folderInput: document.getElementById("folderInput"),
  folderError: document.getElementById("folderError"),
  folderCancelBtn: document.getElementById("folderCancelBtn"),
  folderConfirmBtn: document.getElementById("folderConfirmBtn"),
};

let currentStream = null;
let facingMode = "environment";
let capturedBlob = null;   // {blob, kind: 'photo'|'video', mimeType, ext}
let msalInstance = null;
let account = null;
let mode = "photo";        // 'photo' | 'video'
let mediaRecorder = null;
let recordedChunks = [];
let recTimer = null;
let recSeconds = 0;

const FOLDER_KEY = "camera-onedrive-folder-v1";

// ---------- utilidades de estado visual ----------
function setLed(el, state) {
  el.classList.remove("ok", "warn");
  if (state === "ok") el.classList.add("ok");
  else if (state === "warn") el.classList.add("warn");
}
function setStatus(msg) {
  els.statusMsg.textContent = msg;
}

// ============================================================
// Carpeta destino
// ============================================================
function getActiveFolder() {
  return localStorage.getItem(FOLDER_KEY) || cfg.folderPath;
}
function updateFolderLabel() {
  els.folderLabel.textContent = getActiveFolder();
}
function openFolderPanel() {
  els.folderInput.value = getActiveFolder();
  els.folderError.style.display = "none";
  els.folderPanel.classList.add("show");
  els.folderInput.focus();
}
function closeFolderPanel() {
  els.folderPanel.classList.remove("show");
}
function showFolderError(msg) {
  els.folderError.textContent = msg;
  els.folderError.style.display = "block";
}

async function createOrUseFolder(name) {
  const token = await getAccessToken();
  const res = await fetch("https://graph.microsoft.com/v1.0/me/drive/root/children", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      name,
      folder: {},
      "@microsoft.graph.conflictBehavior": "fail",
    }),
  });
  if (res.ok) return { created: true };
  if (res.status === 409) return { created: false, existed: true };
  const text = await res.text().catch(() => "");
  throw new Error(`Graph API ${res.status}: ${text}`);
}

els.folderBtnOpen.addEventListener("click", openFolderPanel);
els.folderCancelBtn.addEventListener("click", closeFolderPanel);
els.folderConfirmBtn.addEventListener("click", async () => {
  const name = els.folderInput.value.trim();
  if (!name) { showFolderError("Escribe un nombre de carpeta."); return; }
  if (name.includes("/") || name.includes("\\")) {
    showFolderError("Solo un nombre simple, sin / ni \\ (una carpeta de nivel único).");
    return;
  }
  if (!account) { showFolderError("Primero inicia sesión con tu cuenta Microsoft."); return; }

  els.folderConfirmBtn.disabled = true;
  els.folderConfirmBtn.textContent = "Verificando…";
  try {
    const result = await createOrUseFolder(name);
    localStorage.setItem(FOLDER_KEY, name);
    updateFolderLabel();
    closeFolderPanel();
    setStatus(result.created ? `Carpeta "${name}" creada` : `Usando carpeta "${name}" (ya existía)`);
  } catch (e) {
    console.error(e);
    showFolderError("No se pudo crear/verificar la carpeta: " + e.message);
  } finally {
    els.folderConfirmBtn.disabled = false;
    els.folderConfirmBtn.textContent = "Usar / Crear";
  }
});

// ============================================================
// Cola local en IndexedDB (fotos y videos pendientes de subir)
// ============================================================
const DB_NAME = "cam-onedrive-db";
const STORE = "queue";

function openDb() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = () => {
      req.result.createObjectStore(STORE, { keyPath: "id" });
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function dbPut(item) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).put(item);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}
async function dbGetAll() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, "readonly");
    const req = tx.objectStore(STORE).getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function dbDelete(id) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    tx.objectStore(STORE).delete(id);
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
  });
}

function addLogRow(id, name, state) {
  const row = document.createElement("div");
  row.className = `log-row ${state}`;
  row.id = `log-${id}`;
  row.innerHTML = `<span class="dot"></span><span class="name">${name}</span>`;
  els.logPanel.prepend(row);
}
function updateLogRow(id, state) {
  const row = document.getElementById(`log-${id}`);
  if (row) row.className = `log-row ${state}`;
}
async function refreshCount() {
  const q = await dbGetAll();
  els.statusCount.textContent = q.length ? `${q.length} pendiente(s)` : "";
  setLed(els.ledSync, q.length ? "warn" : "ok");
}

// ---------- MSAL / autenticación ----------
function showAuthError(msg) {
  const el = document.getElementById("authError");
  el.textContent = msg;
  el.style.display = "block";
  els.authPanel.classList.add("show");
}

function initAuth() {
  if (!cfg.clientId || cfg.clientId.includes("PEGA_AQUI")) {
    showAuthError("Falta configurar el Client ID en config.js (ver README paso 1-2).");
    return;
  }

  try {
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: cfg.clientId,
        authority: cfg.authority,
        redirectUri: cfg.redirectUri,
      },
      cache: { cacheLocation: "localStorage" },
    });
  } catch (e) {
    console.error(e);
    showAuthError("No se pudo inicializar el login: " + e.message);
    return;
  }

  msalInstance.handleRedirectPromise().then((resp) => {
    if (resp && resp.account) account = resp.account;
    else {
      const accs = msalInstance.getAllAccounts();
      if (accs.length) account = accs[0];
    }
    onAuthChanged();
  }).catch((e) => {
    console.error(e);
    showAuthError("Error al iniciar sesión: " + e.message);
  });

  els.loginBtn.addEventListener("click", () => {
    msalInstance.loginRedirect({ scopes: cfg.scopes }).catch((e) => {
      console.error(e);
      showAuthError("No se pudo abrir el login: " + e.message);
    });
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

// ============================================================
// Cámara: foto y video
// ============================================================
async function startCamera() {
  stopCamera();
  els.hint.textContent = "Iniciando cámara…";
  els.hint.style.display = "block";
  try {
    currentStream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: facingMode } },
      audio: true, // para poder grabar audio en los videos
    });
    els.video.srcObject = currentStream;
    els.hint.style.display = "none";
    els.shutterBtn.disabled = false;
    setLed(els.ledCam, "ok");
  } catch (e) {
    console.error(e);
    els.hint.textContent = "No se pudo acceder a la cámara/micrófono";
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

function filenameFor(ext) {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const stamp = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
  return `${ext === "jpg" ? "foto" : "video"}_${stamp}.${ext}`;
}

// ---- foto ----
function capturePhoto() {
  const video = els.video;
  const canvas = els.canvas;
  canvas.width = video.videoWidth;
  canvas.height = video.videoHeight;
  const ctx = canvas.getContext("2d");
  ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
  canvas.toBlob((blob) => {
    capturedBlob = { blob, kind: "photo", mimeType: "image/jpeg", ext: "jpg" };
    els.previewImg.src = URL.createObjectURL(blob);
    els.previewImg.style.display = "block";
    els.previewVideo.style.display = "none";
    showPreview(true);
  }, "image/jpeg", 0.92);
}

// ---- video ----
function pickVideoMime() {
  const candidates = [
    "video/webm;codecs=vp9,opus",
    "video/webm;codecs=vp8,opus",
    "video/webm",
    "video/mp4",
  ];
  return candidates.find((m) => window.MediaRecorder && MediaRecorder.isTypeSupported(m)) || "";
}

function startRecording() {
  if (!currentStream) return;
  const mimeType = pickVideoMime();
  if (!mimeType) {
    setStatus("Este navegador no permite grabar video");
    return;
  }
  recordedChunks = [];
  try {
    mediaRecorder = new MediaRecorder(currentStream, { mimeType });
  } catch (e) {
    console.error(e);
    setStatus("No se pudo iniciar la grabación: " + e.message);
    return;
  }
  mediaRecorder.ondataavailable = (e) => {
    if (e.data && e.data.size > 0) recordedChunks.push(e.data);
  };
  mediaRecorder.onstop = () => {
    const blob = new Blob(recordedChunks, { type: mimeType });
    const ext = mimeType.includes("mp4") ? "mp4" : "webm";
    capturedBlob = { blob, kind: "video", mimeType: mimeType.split(";")[0], ext };
    els.previewVideo.src = URL.createObjectURL(blob);
    els.previewVideo.style.display = "block";
    els.previewImg.style.display = "none";
    showPreview(true);
  };
  mediaRecorder.start();
  recSeconds = 0;
  updateRecBadge();
  els.recBadge.classList.add("show");
  recTimer = setInterval(() => {
    recSeconds++;
    updateRecBadge();
  }, 1000);
  els.shutterBtn.classList.add("recording");
}

function stopRecording() {
  if (mediaRecorder && mediaRecorder.state !== "inactive") mediaRecorder.stop();
  clearInterval(recTimer);
  els.recBadge.classList.remove("show");
  els.shutterBtn.classList.remove("recording");
}

function updateRecBadge() {
  const m = String(Math.floor(recSeconds / 60)).padStart(2, "0");
  const s = String(recSeconds % 60).padStart(2, "0");
  els.recTime.textContent = `${m}:${s}`;
}

let isRecording = false;
function handleShutter() {
  if (mode === "photo") {
    capturePhoto();
    return;
  }
  // modo video: toca para empezar, toca de nuevo para terminar
  if (!isRecording) {
    isRecording = true;
    startRecording();
  } else {
    isRecording = false;
    stopRecording();
  }
}

function setMode(newMode) {
  if (isRecording) return; // no cambiar de modo a mitad de una grabación
  mode = newMode;
  els.modePhotoBtn.classList.toggle("active", mode === "photo");
  els.modeVideoBtn.classList.toggle("active", mode === "video");
  els.shutterBtn.classList.toggle("mode-video", mode === "video");
  els.shutterBtn.title = mode === "photo" ? "Tomar foto" : "Grabar video";
}

function showPreview(show) {
  els.previewOverlay.classList.toggle("show", show);
  els.previewActions.classList.toggle("show", show);
  els.captureControls.style.display = show ? "none" : "flex";
}

// ============================================================
// Subida a OneDrive (Microsoft Graph)
// ============================================================
async function uploadBlob(blob, filename, mimeType) {
  const token = await getAccessToken();
  const folder = getActiveFolder();
  const path = `${folder}/${filename}`;
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(path).replace(/%2F/g, "/")}:/content`;
  const res = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": mimeType,
    },
    body: blob,
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    throw new Error(`Graph API ${res.status}: ${text}`);
  }
  return res.json();
}

async function queueMedia(item) {
  const filename = filenameFor(item.ext);
  const id = `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
  await dbPut({ id, filename, blob: item.blob, mimeType: item.mimeType });
  addLogRow(id, filename, "pending");
  refreshCount();
  flushQueue();
}

let flushing = false;
async function flushQueue() {
  if (flushing || !account) return;
  flushing = true;
  setLed(els.ledSync, "warn");
  const q = await dbGetAll();
  for (const item of q) {
    try {
      setStatus(`Subiendo ${item.filename}…`);
      await uploadBlob(item.blob, item.filename, item.mimeType);
      updateLogRow(item.id, "done");
      await dbDelete(item.id);
      setStatus(`${item.filename} subida ✓`);
    } catch (e) {
      console.error(e);
      updateLogRow(item.id, "error");
      setStatus(`Error al subir ${item.filename} — se reintentará`);
      break; // dejamos el resto en cola, reintenta más tarde
    }
  }
  await refreshCount();
  flushing = false;
}

// ---------- eventos UI ----------
els.shutterBtn.addEventListener("click", handleShutter);
els.discardBtn.addEventListener("click", () => {
  capturedBlob = null;
  showPreview(false);
});
els.uploadBtn.addEventListener("click", async () => {
  if (!capturedBlob) return;
  const item = capturedBlob;
  capturedBlob = null;
  showPreview(false);
  await queueMedia(item);
});
els.switchCamBtn.addEventListener("click", () => {
  facingMode = facingMode === "environment" ? "user" : "environment";
  startCamera();
});
els.logToggleBtn.addEventListener("click", () => {
  els.logPanel.classList.toggle("open");
});
els.modePhotoBtn.addEventListener("click", () => setMode("photo"));
els.modeVideoBtn.addEventListener("click", () => setMode("video"));

window.addEventListener("online", flushQueue);

// ---------- arranque ----------
async function boot() {
  initAuth();
  startCamera();
  updateFolderLabel();
  await refreshCount();
  const q = await dbGetAll();
  q.forEach((item) => addLogRow(item.id, item.filename, "pending"));

  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch((e) => console.error("SW error", e));
  }
}
boot();
