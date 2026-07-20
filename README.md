# Cámara → OneDrive (PWA)

App web progresiva (PWA) para Android: abre la cámara del celular, tomas la
foto y la subes con un botón a una carpeta de tu OneDrive (por defecto
`CameraUploads`). Se puede "instalar" desde Chrome como si fuera una app
normal, y queda con ícono en el escritorio.

## Qué incluye

- `index.html`, `app.js` — interfaz y lógica (cámara + subida)
- `config.js` — **aquí pones tus datos de Azure AD**
- `manifest.json`, `sw.js`, `icons/` — lo necesario para que sea instalable
- Autenticación con `msal-browser` (librería oficial de Microsoft)
- Sube los archivos a OneDrive con Microsoft Graph API
- Cola local: si falla la subida o no hay internet, la foto queda guardada
  en el celular y se reintenta sola al recuperar conexión

## Paso 1 — Registrar la app en Azure (obligatorio, una sola vez)

Para poder llamar a la API de OneDrive necesitas un "Client ID" gratuito:

1. Entra a https://portal.azure.com con tu cuenta Microsoft (puede ser la
   misma cuenta personal donde tienes el OneDrive).
2. Ve a **Microsoft Entra ID** → **Registros de aplicaciones** → **Nuevo
   registro**.
3. Nombre: el que quieras (ej. "Camara OneDrive"). En "Tipos de cuenta
   admitidos" elige **Cuentas en cualquier organización y cuentas
   personales de Microsoft**.
4. En **Redireccionar URI** elige tipo **SPA (Single-page application)** y
   pon la URL donde vas a publicar la app, por ejemplo:
   `https://tuusuario.github.io/onedrive-camera-pwa/`
   (si todavía no la tienes, puedes agregarla después desde
   **Autenticación**).
5. Clic en **Registrar**. En la página que se abre, copia el **Id. de
   aplicación (cliente)** — es un GUID largo.
6. Ve a **Autenticación** y confirma que el Redirect URI quedó guardado
   como tipo **SPA** (no "Web").
7. No hace falta crear ningún "secreto" — esta app corre en el navegador
   y no lo necesita.

## Paso 2 — Configurar la app

Abre `config.js` y reemplaza:

```js
clientId: "PEGA_AQUI_TU_CLIENT_ID",
```

con el Id. de aplicación que copiaste. Si quieres cambiar el nombre de la
carpeta de destino en OneDrive, edita también:

```js
folderPath: "CameraUploads",
```

## Paso 3 — Publicar la app (necesita HTTPS)

El navegador solo permite usar la cámara y hacer login por HTTPS. Formas
gratuitas y simples de publicarla:

**Opción A — GitHub Pages**
1. Crea un repositorio nuevo y sube todos estos archivos.
2. En **Settings → Pages**, activa Pages sobre la rama `main`.
3. La URL quedará como `https://tuusuario.github.io/nombre-repo/`.
4. Vuelve al registro de Azure (Paso 1.4) y confirma que el Redirect URI
   coincide EXACTO con esa URL (con la barra `/` final incluida).

**Opción B — Netlify / Vercel (arrastrar y soltar)**
Sube la carpeta completa desde su panel web; ambos dan HTTPS automático.
Actualiza el Redirect URI en Azure con la URL que te asignen.

## Paso 4 — Usarla desde el celular

1. Abre la URL publicada en Chrome (Android).
2. Menú (⋮) → **Instalar app** (o "Agregar a pantalla de inicio").
3. Ábrela, toca **Iniciar sesión** y entra con tu cuenta Microsoft.
4. Da permiso de cámara cuando lo pida el navegador.
5. Toca el botón central (obturador) para tomar la foto, revisa la
   vista previa y toca **Subir a OneDrive**.

Las fotos aparecerán en OneDrive dentro de la carpeta configurada
(`CameraUploads` por defecto), que se crea sola en la primera subida.

## Notas técnicas

- El indicador **AUTH** (arriba) se pone verde cuando hay sesión activa;
  **CAM** cuando la cámara está lista; **SYNC** en ámbar mientras hay
  fotos pendientes de subir y verde cuando la cola está vacía.
- El ícono "≡" abre un registro con el estado de cada foto (pendiente /
  subida / error).
- Los permisos de Graph pedidos son `User.Read` y `Files.ReadWrite`
  (acceso solo a los archivos del propio usuario, no a todo OneDrive de
  la organización).
- Si cambias el dominio de publicación, recuerda actualizar el Redirect
  URI en Azure Portal, o el login fallará con error `AADSTS50011`.
