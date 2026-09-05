# Guía Práctica: Integración de Autenticación con Microsoft Entra ID en React (MSAL)

**Asignatura:** Cloud Native / Desarrollo Web

**Tecnologías:** React, Vite, TypeScript, `@azure/msal-react`, Microsoft Entra ID

---

### Contexto e Introducción

Microsoft Authentication Library (MSAL) es la librería oficial de Microsoft para gestionar el flujo OAuth2 / OpenID Connect (OIDC). En aplicaciones de una sola página (SPA) escritas en React, nos permite:

* Autenticar usuarios con sus cuentas institucionales o corporativas.


* Obtener tokens de identidad y de acceso (`ID Token` y `Access Token`).


* Adjuntar dichos tokens a peticiones HTTP para consumir APIs protegidas.



---

### Requisitos Previos

1. Node.js (v18+) e `npm` instalados.
2. Una cuenta de Microsoft activa (de la institución o de pruebas de Azure Pass).
3. Acceso a **Azure Portal**.

---

### Paso 1: Configurar la Aplicación en Microsoft Entra ID

Antes de escribir código, debemos registrar la aplicación en el directorio de identidades de Azure para definir los parámetros de autenticación.

1. Ingresa a **[portal.azure.com](https://www.google.com/search?q=https://portal.azure.com)**.
2. En la barra de búsqueda superior, escribe **Registros de aplicaciones** (*App registrations*) y selecciona la opción con el icono del folder y la llave.


3. Haz clic en **+ Nuevo registro** (*+ New registration*).


4. Completa la configuración básica:
* **Nombre:** `MiAppReact` (o el nombre de tu proyecto).
* **Tipos de cuenta compatibles:** Selecciona *Solo las cuentas de este directorio organizativo (Inquilino único / Single tenant)*.
* **URI de redirección:** Selecciona la plataforma **SPA (Aplicación de una sola página)** e ingresa `http://localhost:5173`.


5. Haz clic en **Registrar**.


6. En la pantalla de **Información general** (*Overview*), copia y guarda los siguientes valores:


* **ID de la aplicación (cliente)** (*Application/Client ID*)


* **ID del directorio (inquilino)** (*Directory/Tenant ID*)





---

### Paso 2: Crear el Proyecto e Instalar Librerías

1. Crea la estructura base del proyecto usando Vite y TypeScript:
```bash
npm create vite@latest mi-app-msal -- --template react-ts
cd mi-app-msal

```


2. Instala las dependencias oficiales de Microsoft para navegadores y React:
```bash
npm install @azure/msal-browser @azure/msal-react
npm install

```



---

### Paso 3: Configurar Variables de Entorno

Crea un archivo `.env` en la raíz del proyecto para resguardar las credenciales obtenidas en Azure:

```env
VITE_AZURE_CLIENT_ID=TU_CLIENT_ID_COPIADO
VITE_AZURE_TENANT_ID=TU_TENANT_ID_COPIADO
VITE_AZURE_REDIRECT_URI=http://localhost:5173

```

---

### Paso 4: Definir la Configuración de MSAL (`src/authConfig.ts`)

Crea el archivo `src/authConfig.ts` para mapear los parámetros de autenticación y los permisos (*scopes*) solicitados:

```typescript
import type { Configuration } from '@azure/msal-browser';
import { LogLevel } from '@azure/msal-browser';

export const msalConfig: Configuration = {
  auth: {
    clientId: import.meta.env.VITE_AZURE_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AZURE_TENANT_ID}`,
    redirectUri: import.meta.env.VITE_AZURE_REDIRECT_URI,
  },
  cache: {
    cacheLocation: 'localStorage',
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
        if (containsPii) return;
        if (level === LogLevel.Error) console.error(message);
      },
      logLevel: LogLevel.Error,
    },
  },
};

export const loginRequest = {
  scopes: ['User.Read'],
};

```

---

### Paso 5: Inicializar la Instancia Global y el Proveedor (`src/main.tsx`)

Instancia `PublicClientApplication` y envuelve la aplicación dentro de `<MsalProvider>`:

```tsx
import React from 'react';
import ReactDOM from 'react-dom/client';
import { PublicClientApplication, EventType } from '@azure/msal-browser';
import type { EventMessage, AuthenticationResult } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import { msalConfig } from './authConfig';
import App from './App';

const msalInstance = new PublicClientApplication(msalConfig);

// Activar la cuenta si ya existe una sesión previa
if (!msalInstance.getActiveAccount() && msalInstance.getAllAccounts().length > 0) {
  msalInstance.setActiveAccount(msalInstance.getAllAccounts()[0]);
}

msalInstance.addEventCallback((event: EventMessage) => {
  if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
    const payload = event.payload as AuthenticationResult;
    msalInstance.setActiveAccount(payload.account);
  }
});

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <MsalProvider instance={msalInstance}>
      <App />
    </MsalProvider>
  </React.StrictMode>
);

```

---

### Paso 6: Crear el Hook para Llamadas HTTP Protegidas (`src/useApi.ts`)

Para adjuntar tokens Bearer a nuestras peticiones (equivalente a un *HTTP Interceptor*), creamos un Hook que obtiene el token en segundo plano usando `acquireTokenSilent`:

```typescript
import { useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig';

export function useApi() {
  const { instance, accounts } = useMsal();

  const fetchWithToken = async (url: string) => {
    const account = accounts[0] || instance.getActiveAccount();
    if (!account) throw new Error('No hay una cuenta activa');

    // Solicitar token silenciosamente
    const response = await instance.acquireTokenSilent({
      ...loginRequest,
      account,
    });

    // Adjuntar token en el header HTTP
    return fetch(url, {
      headers: {
        Authorization: `Bearer ${response.accessToken}`,
      },
    });
  };

  return { fetchWithToken };
}

```

---

### Paso 7: Proteger Vistas y Componentes (`src/ProtectedData.tsx`)

Utilizamos los componentes `<AuthenticatedTemplate>` y `<UnauthenticatedTemplate>` para controlar la visibilidad según el estado de sesión (equivalente a los *Guards* de rutas):

```tsx
import { useState } from 'react';
import { AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { useApi } from './useApi';

export function ProtectedData() {
  const { fetchWithToken } = useApi();
  const [data, setData] = useState<any>(null);
  const [loading, setLoading] = useState(false);

  const handleFetchData = async () => {
    setLoading(true);
    try {
      const res = await fetchWithToken('https://graph.microsoft.com/v1.0/me');
      const json = await res.json();
      setData(json);
    } catch (err) {
      console.error(err);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div>
      <AuthenticatedTemplate>
        <button onClick={handleFetchData} disabled={loading}>
          {loading ? 'Consultando...' : 'Obtener Datos del Usuario vía API'}
        </button>
        {data && <pre>{JSON.stringify(data, null, 2)}</pre>}
      </AuthenticatedTemplate>

      <UnauthenticatedTemplate>
        <p>⚠️ Acceso denegado. Debes iniciar sesión para consultar este recurso.</p>
      </UnauthenticatedTemplate>
    </div>
  );
}

```

---

### Paso 8: Construir la Interfaz Principal (`src/App.tsx`)

Integra el flujo de login/logout junto a las vistas protegidas:

```tsx
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { InteractionStatus } from '@azure/msal-browser';
import { loginRequest } from './authConfig';
import { ProtectedData } from './ProtectedData';

export default function App() {
  const { instance, accounts, inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();
  const currentUser = accounts[0];

  const handleLogin = () => {
    if (inProgress === InteractionStatus.None) {
      instance.loginRedirect(loginRequest).catch((e) => console.error(e));
    }
  };

  const handleLogout = () => {
    if (inProgress === InteractionStatus.None) {
      instance.logoutRedirect({ postLogoutRedirectUri: '/' }).catch((e) => console.error(e));
    }
  };

  return (
    <div style={{ padding: '2rem', fontFamily: 'sans-serif' }}>
      <h1>Portal de Autenticación con Microsoft Entra ID</h1>

      {isAuthenticated ? (
        <div>
          <p>Bienvenido, <strong>{currentUser?.name || currentUser?.username}</strong></p>
          <button onClick={handleLogout}>Cerrar Sesión</button>
          <hr style={{ margin: '1.5rem 0' }} />
          <ProtectedData />
        </div>
      ) : (
        <div>
          <p>Debes iniciar sesión con tu cuenta institucional para continuar.</p>
          <button onClick={handleLogin} disabled={inProgress !== InteractionStatus.None}>
            {inProgress !== InteractionStatus.None ? 'Cargando...' : 'Iniciar Sesión'}
          </button>
        </div>
      )}
    </div>
  );
}

```

---

### Paso 9: Ejecución y Pruebas

1. Inicia el servidor de desarrollo:
```bash
npm run dev

```


2. Navega a `http://localhost:5173`.
3. Haz clic en **Iniciar Sesión**. Serás redirigido al portal oficial de Microsoft.
4. Tras validar credenciales, la app desplegará la vista protegida con la opción de consultar la API con Bearer Token.