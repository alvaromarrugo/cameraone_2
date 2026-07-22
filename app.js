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
  folderPanel: document.getElementById("folderPanel"),
  folderInput: document.getElementById("folderInput"),
  folderError: document.getElementById("folderError"),
  folderCancelBtn: document.getElementById("folderCancelBtn"),
  folderConfirmBtn: document.getElementById("folderConfirmBtn"),
  qualityPanel: document.getElementById("qualityPanel"),
  qualityOptionList: document.getElementById("qualityOptionList"),
  qualityCancelBtn: document.getElementById("qualityCancelBtn"),
  cameraPanel: document.getElementById("cameraPanel"),
  cameraOptionList: document.getElementById("cameraOptionList"),
  cameraError: document.getElementById("cameraError"),
  cameraCancelBtn: document.getElementById("cameraCancelBtn"),
  closeAppBtn: document.getElementById("closeAppBtn"),
  settingsSummary: document.getElementById("settingsSummary"),
  settingsBtnOpen: document.getElementById("settingsBtnOpen"),
  settingsPanel: document.getElementById("settingsPanel"),
  settingsFolderVal: document.getElementById("settingsFolderVal"),
  settingsQualityVal: document.getElementById("settingsQualityVal"),
  settingsCameraVal: document.getElementById("settingsCameraVal"),
  settingsOpenFolder: document.getElementById("settingsOpenFolder"),
  settingsOpenQuality: document.getElementById("settingsOpenQuality"),
  settingsOpenCamera: document.getElementById("settingsOpenCamera"),
  settingsCloseBtn: document.getElementById("settingsCloseBtn"),
  flashBtn: document.getElementById("flashBtn"),
  convertingOverlay: document.getElementById("convertingOverlay"),
  convertingProgress: document.getElementById("convertingProgress"),
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
let torchOn = false;

const FOLDER_KEY = "camera-onedrive-folder-v1";
const RES_KEY = "camera-onedrive-resolution-v1";
const CAMERA_KEY = "camera-onedrive-device-v1";
const RES_PRESETS = {
  sd:  { label: "Estándar (480p)", width: 640,  height: 480,  crf: 24 },
  hd:  { label: "HD (720p)",       width: 1280, height: 720,  crf: 23 },
  fhd: { label: "Full HD (1080p)", width: 1920, height: 1080, crf: 21 },
  uhd: { label: "4K (máxima disponible)", width: 3840, height: 2160, crf: 20 },
};

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
  els.settingsFolderVal.textContent = getActiveFolder();
  updateSettingsSummary();
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
// Calidad / resolución de cámara
// ============================================================
function getActiveResKey() {
  const k = localStorage.getItem(RES_KEY);
  return RES_PRESETS[k] ? k : "fhd";
}
function updateQualityLabel(actualW, actualH) {
  const preset = RES_PRESETS[getActiveResKey()];
  els.settingsQualityVal.textContent = actualW
    ? `${preset.label} — real: ${actualW}×${actualH}`
    : preset.label;
  updateSettingsSummary();
}
function renderQualityOptions() {
  const activeKey = getActiveResKey();
  els.qualityOptionList.innerHTML = "";
  Object.entries(RES_PRESETS).forEach(([key, preset]) => {
    const btn = document.createElement("button");
    btn.className = "option-btn" + (key === activeKey ? " active" : "");
    btn.innerHTML = `<span>${preset.label}</span><span class="sub">${preset.width}×${preset.height}</span>`;
    btn.addEventListener("click", async () => {
      localStorage.setItem(RES_KEY, key);
      closeQualityPanel();
      updateQualityLabel();
      setStatus("Aplicando nueva calidad…");
      await startCamera();
    });
    els.qualityOptionList.appendChild(btn);
  });
}
function openQualityPanel() {
  renderQualityOptions();
  els.qualityPanel.classList.add("show");
}
function closeQualityPanel() {
  els.qualityPanel.classList.remove("show");
}
els.qualityCancelBtn.addEventListener("click", closeQualityPanel);

// ============================================================
// Cámara física específica (para celulares con varios lentes traseros)
// ============================================================
function getActiveDeviceId() {
  return localStorage.getItem(CAMERA_KEY) || null; // null = automático (usa facingMode)
}
function updateCameraLabel(labelText) {
  if (labelText) { els.settingsCameraVal.textContent = labelText; updateSettingsSummary(); return; }
  const id = getActiveDeviceId();
  els.settingsCameraVal.textContent = id ? "Cámara seleccionada" : "Automática";
  updateSettingsSummary();
}

async function listCameras() {
  const devices = await navigator.mediaDevices.enumerateDevices();
  return devices.filter((d) => d.kind === "videoinput");
}

async function renderCameraOptions() {
  els.cameraError.style.display = "none";
  els.cameraOptionList.innerHTML = "<p style='color:var(--muted);font-family:var(--mono);font-size:11px;'>Buscando cámaras…</p>";
  let cams = [];
  try {
    cams = await listCameras();
  } catch (e) {
    console.error(e);
    els.cameraOptionList.innerHTML = "";
    els.cameraError.textContent = "No se pudo obtener la lista de cámaras: " + e.message;
    els.cameraError.style.display = "block";
    return;
  }
  els.cameraOptionList.innerHTML = "";
  const activeId = getActiveDeviceId();

  // opción "Automático"
  const autoBtn = document.createElement("button");
  autoBtn.className = "option-btn" + (!activeId ? " active" : "");
  autoBtn.innerHTML = `<span>Automática</span><span class="sub">trasera / frontal</span>`;
  autoBtn.addEventListener("click", async () => {
    localStorage.removeItem(CAMERA_KEY);
    updateCameraLabel();
    closeCameraPanel();
    setStatus("Aplicando cámara automática…");
    await startCamera();
  });
  els.cameraOptionList.appendChild(autoBtn);

  if (!cams.length) {
    const p = document.createElement("p");
    p.style.cssText = "color:var(--muted);font-family:var(--mono);font-size:11px;";
    p.textContent = "No se detectaron cámaras adicionales (o el navegador aún no dio permiso para verlas).";
    els.cameraOptionList.appendChild(p);
    return;
  }

  cams.forEach((cam, i) => {
    const label = cam.label || `Cámara ${i + 1}`;
    const btn = document.createElement("button");
    btn.className = "option-btn" + (cam.deviceId === activeId ? " active" : "");
    btn.innerHTML = `<span>${label}</span><span class="sub">tocar para usar</span>`;
    btn.addEventListener("click", async () => {
      localStorage.setItem(CAMERA_KEY, cam.deviceId);
      updateCameraLabel(label);
      closeCameraPanel();
      setStatus("Aplicando cámara seleccionada…");
      await startCamera();
    });
    els.cameraOptionList.appendChild(btn);
  });
}

function openCameraPanel() {
  els.cameraPanel.classList.add("show");
  renderCameraOptions();
}
function closeCameraPanel() {
  els.cameraPanel.classList.remove("show");
}
els.cameraCancelBtn.addEventListener("click", closeCameraPanel);

// ============================================================
// Panel de Ajustes (resumen + acceso a carpeta/calidad/cámara)
// ============================================================
function updateSettingsSummary() {
  const folder = els.settingsFolderVal ? els.settingsFolderVal.textContent : "";
  const quality = RES_PRESETS[getActiveResKey()].label;
  const camera = getActiveDeviceId() ? "Personalizada" : "Automática";
  els.settingsSummary.textContent = `📁 ${folder} · 🎚️ ${quality} · 📷 ${camera}`;
}
function openSettingsPanel() {
  els.settingsPanel.classList.add("show");
}
function closeSettingsPanel() {
  els.settingsPanel.classList.remove("show");
}
els.settingsBtnOpen.addEventListener("click", openSettingsPanel);
els.settingsCloseBtn.addEventListener("click", closeSettingsPanel);
els.settingsOpenFolder.addEventListener("click", () => {
  closeSettingsPanel();
  openFolderPanel();
});
els.settingsOpenQuality.addEventListener("click", () => {
  closeSettingsPanel();
  openQualityPanel();
});
els.settingsOpenCamera.addEventListener("click", () => {
  closeSettingsPanel();
  openCameraPanel();
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
  const preset = RES_PRESETS[getActiveResKey()];
  const deviceId = getActiveDeviceId();
  const videoConstraints = deviceId
    ? { deviceId: { exact: deviceId }, width: { ideal: preset.width }, height: { ideal: preset.height } }
    : { facingMode: { ideal: facingMode }, width: { ideal: preset.width }, height: { ideal: preset.height } };
  try {
    currentStream = await navigator.mediaDevices.getUserMedia({
      video: videoConstraints,
      audio: true, // para poder grabar audio en los videos
    });
  } catch (e) {
    console.error(e);
    if (deviceId) {
      // la cámara guardada ya no existe o falló (celular distinto, lente removido, etc.)
      console.warn("Fallback a cámara automática:", e.message);
      localStorage.removeItem(CAMERA_KEY);
      updateCameraLabel();
      try {
        currentStream = await navigator.mediaDevices.getUserMedia({
          video: { facingMode: { ideal: facingMode }, width: { ideal: preset.width }, height: { ideal: preset.height } },
          audio: true,
        });
      } catch (e2) {
        console.error(e2);
        els.hint.textContent = "No se pudo acceder a la cámara/micrófono";
        setLed(els.ledCam, "");
        setStatus("Permiso de cámara denegado o no disponible");
        return;
      }
    } else {
      els.hint.textContent = "No se pudo acceder a la cámara/micrófono";
      setLed(els.ledCam, "");
      setStatus("Permiso de cámara denegado o no disponible");
      return;
    }
  }
  els.video.srcObject = currentStream;
  els.hint.style.display = "none";
  els.shutterBtn.disabled = false;
  setLed(els.ledCam, "ok");
  const track = currentStream.getVideoTracks()[0];
  const settings = track ? track.getSettings() : {};
  updateQualityLabel(settings.width, settings.height);
  setupTorch(track);
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

// ============================================================
// Conversión de video a MP4 (ffmpeg.wasm, autohospedado, sin CDN)
// ============================================================
let ffmpegInstance = null;
let ffmpegLoading = null;

function showConverting(show, text) {
  els.convertingOverlay.classList.toggle("show", show);
  if (text) els.convertingProgress.textContent = text;
}

async function ensureFFmpeg() {
  if (ffmpegInstance) return ffmpegInstance;
  if (ffmpegLoading) return ffmpegLoading;
  ffmpegLoading = (async () => {
    showConverting(true, "Preparando conversor de video (solo la primera vez)…");
    const ffmpeg = new FFmpegWASM.FFmpeg();
    ffmpeg.on("progress", ({ progress }) => {
      const pct = Math.min(100, Math.max(0, Math.round(progress * 100)));
      els.convertingProgress.textContent = `Convirtiendo… ${pct}%`;
    });
    await ffmpeg.load({
      coreURL: "ffmpeg-vendor/ffmpeg-core.js",
      wasmURL: "ffmpeg-vendor/ffmpeg-core.wasm",
    });
    ffmpegInstance = ffmpeg;
    return ffmpeg;
  })();
  try {
    return await ffmpegLoading;
  } finally {
    ffmpegLoading = null;
  }
}

async function convertWebmToMp4(webmBlob) {
  showConverting(true, "Preparando…");
  const preset = RES_PRESETS[getActiveResKey()];
  const inputName = "input.webm";
  const outputName = "output.mp4";
  const ffmpeg = await ensureFFmpeg();

  const inputData = new Uint8Array(await webmBlob.arrayBuffer());
  await ffmpeg.writeFile(inputName, inputData);

  showConverting(true, "Convirtiendo… 0%");
  await ffmpeg.exec([
    "-i", inputName,
    "-c:v", "libx264",
    "-preset", "veryfast",
    "-crf", String(preset.crf),
    "-pix_fmt", "yuv420p",
    "-c:a", "aac",
    "-b:a", "128k",
    "-movflags", "+faststart",
    outputName,
  ]);

  const outputData = await ffmpeg.readFile(outputName);
  const mp4Blob = new Blob([outputData.buffer], { type: "video/mp4" });

  await ffmpeg.deleteFile(inputName).catch(() => {});
  await ffmpeg.deleteFile(outputName).catch(() => {});

  return mp4Blob;
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
  const preset = RES_PRESETS[getActiveResKey()];
  const videoBitsPerSecond = Math.round(preset.width * preset.height * 0.12); // ~escala con la resolución
  try {
    mediaRecorder = new MediaRecorder(currentStream, { mimeType, videoBitsPerSecond });
  } catch (e) {
    console.error(e);
    setStatus("No se pudo iniciar la grabación: " + e.message);
    return;
  }
  mediaRecorder.ondataavailable = (e) => {
    if (e.data && e.data.size > 0) recordedChunks.push(e.data);
  };
  mediaRecorder.onstop = async () => {
    const webmBlob = new Blob(recordedChunks, { type: mimeType });
    try {
      const mp4Blob = await convertWebmToMp4(webmBlob);
      capturedBlob = { blob: mp4Blob, kind: "video", mimeType: "video/mp4", ext: "mp4" };
      els.previewVideo.src = URL.createObjectURL(mp4Blob);
    } catch (e) {
      console.error(e);
      setStatus("No se pudo convertir a MP4, se conserva el video original (" + (mimeType.includes("mp4") ? "mp4" : "webm") + ")");
      const ext = mimeType.includes("mp4") ? "mp4" : "webm";
      capturedBlob = { blob: webmBlob, kind: "video", mimeType: mimeType.split(";")[0], ext };
      els.previewVideo.src = URL.createObjectURL(webmBlob);
    }
    showConverting(false);
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

// ============================================================
// Flash / linterna (torch de la cámara trasera)
// ============================================================
function setupTorch(track) {
  torchOn = false;
  els.flashBtn.classList.remove("on");
  const caps = track && track.getCapabilities ? track.getCapabilities() : {};
  const supported = !!(caps && caps.torch);
  els.flashBtn.style.display = supported ? "flex" : "none";
}

els.flashBtn.addEventListener("click", async () => {
  if (!currentStream) return;
  const track = currentStream.getVideoTracks()[0];
  if (!track) return;
  const next = !torchOn;
  try {
    await track.applyConstraints({ advanced: [{ torch: next }] });
    torchOn = next;
    els.flashBtn.classList.toggle("on", torchOn);
  } catch (e) {
    console.error(e);
    setStatus("No se pudo controlar el flash en este celular/navegador");
  }
});

// ============================================================
// Cerrar la app
// ============================================================
els.closeAppBtn.addEventListener("click", () => {
  try { window.close(); } catch (e) { /* ignorado */ }
  // Android no permite que una PWA se cierre sola por seguridad;
  // si sigue visible medio segundo después, avisamos cómo cerrarla a mano.
  setTimeout(() => {
    setStatus("Para cerrarla del todo, deslízala en 'Apps recientes' de tu celular.");
  }, 400);
});

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
  localStorage.removeItem(CAMERA_KEY);
  updateCameraLabel();
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
  updateQualityLabel();
  updateCameraLabel();
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
