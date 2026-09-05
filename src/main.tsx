import React from "react";
import ReactDOM from "react-dom/client";
import { PublicClientApplication, EventType } from "@azure/msal-browser";
import type { EventMessage, AuthenticationResult } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./authConfig";
import App from "./App";
import { AuthProvider } from "react-oidc-context";
import "./index.css";

// Configuración de Cognito
const cognitoAuthConfig = {
  authority: "https://cognito-idp.us-east-1.amazonaws.com/us-east-1_sm3KXrDjs",
  client_id: "3jbqjca1hj0idgm52ti3rdcu43",
  redirect_uri: "http://localhost:5173/",
  response_type: "code",
  scope: "phone openid email",
};

// 1. Crear la instancia global
const msalInstance = new PublicClientApplication(msalConfig);

// 2. INICIALIZAR MSAL ANTES DE RENDERIZAR (Requisito para MSAL v3+)
msalInstance
  .initialize()
  .then(() => {
    // Manejar la cuenta activa en la carga inicial
    if (
      !msalInstance.getActiveAccount() &&
      msalInstance.getAllAccounts().length > 0
    ) {
      msalInstance.setActiveAccount(msalInstance.getAllAccounts()[0]);
    }

    // Escuchar eventos de login
    msalInstance.addEventCallback((event: EventMessage) => {
      if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
        const payload = event.payload as AuthenticationResult;
        msalInstance.setActiveAccount(payload.account);
      }
    });

    // 3. Renderizar la aplicación SOLAMENTE después de que MSAL esté listo
    ReactDOM.createRoot(document.getElementById("root")!).render(
      <React.StrictMode>
        <AuthProvider {...cognitoAuthConfig}>
          <MsalProvider instance={msalInstance}>
            <App />
          </MsalProvider>
        </AuthProvider>
      </React.StrictMode>,
    );
  })
  .catch((e) => {
    console.error("Error al inicializar MSAL:", e);
  });
