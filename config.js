// ============================================================
// CONFIGURACIÓN — edita estos valores antes de publicar la app
// ============================================================
window.APP_CONFIG = {
  // ID de la aplicación (Application/Client ID) que obtienes al
  // registrar la app en Azure Portal > Entra ID > Registros de app.
  // Ver README.md paso 1.
  clientId: "3642ff38-ccf8-41d4-be59-057f00365367",

  // "common" sirve para cuentas personales (outlook.com, hotmail, etc.)
  // y también para cuentas de trabajo/escuela. Normalmente no hace
  // falta cambiarlo.
  authority: "https://login.microsoftonline.com/common",

  // Debe coincidir EXACTAMENTE con el "Redirect URI" configurado en
  // Azure Portal (tipo SPA). Ejemplo: "https://tuusuario.github.io/onedrive-camera-pwa/"
  redirectUri: window.location.origin + window.location.pathname,

  // Carpeta dentro de "Mis archivos" de OneDrive donde se subirán
  // las fotos. Se crea sola la primera vez si no existe.
  folderPath: "CameraUploads",

  // Permisos que la app pide a Microsoft Graph.
  scopes: ["User.Read", "Files.ReadWrite"]
};
