// src/App.tsx
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";
import { useAuth } from "react-oidc-context";
import { loginRequest } from "./authConfig";
import { ProtectedData } from "./ProtectedData";
import "./App.css";

export default function App() {
  // --- Estado de Microsoft MSAL ---
  const { instance, accounts, inProgress } = useMsal();
  const isMicrosoftAuth = useIsAuthenticated();
  const currentUserMs = accounts[0];

  // --- Estado de AWS Cognito ---
  const auth = useAuth();
  const isCognitoAuth = auth.isAuthenticated;

  // --- Estado Global Combinado ---
  const isAnyUserAuthenticated = isMicrosoftAuth || isCognitoAuth;

  // --- Handlers de Login ---
  const handleLoginMicrosoft = () => {
    if (inProgress === InteractionStatus.None) {
      instance.loginRedirect(loginRequest).catch((e) => console.error(e));
    }
  };

  const handleLoginCognito = () => {
    auth.signinRedirect();
  };

  // --- Handler de Logout ---
  const handleLogout = () => {
    if (isMicrosoftAuth && inProgress === InteractionStatus.None) {
      instance
        .logoutRedirect({ postLogoutRedirectUri: "/" })
        .catch((e) => console.error(e));
    }

    if (isCognitoAuth) {
      auth.removeUser();
      const clientId = "3jbqjca1hj0idgm52ti3rdcu43"; // Reemplaza con tu Client ID
      const logoutUri = "http://localhost:5173/"; // Reemplaza con tu URL de redirección
      const cognitoDomain = "https://<user pool domain>"; // Reemplaza con tu dominio Cognito
      window.location.href = `${cognitoDomain}/logout?client_id=${clientId}&logout_uri=${encodeURIComponent(logoutUri)}`;
    }
  };

  // Manejo de estados de carga y error para Cognito
  if (auth.isLoading) return <div className="layout">Cargando Cognito...</div>;
  if (auth.error)
    return <div className="layout">Error en Cognito: {auth.error.message}</div>;

  return (
    <div className="layout">
      <header className="navbar">
        <div className="logo">
          ⚡ <span>Portal MiApp</span>
        </div>
        <div>
          {isAnyUserAuthenticated ? (
            <button
              className="btn btn-logout"
              onClick={handleLogout}
              disabled={inProgress !== InteractionStatus.None}
            >
              Cerrar Sesión
            </button>
          ) : (
            <span style={{ fontSize: "0.9rem", color: "#666" }}>
              No autenticado
            </span>
          )}
        </div>
      </header>

      <main className="container">
        {isAnyUserAuthenticated ? (
          <div className="card">
            <div className="avatar">
              {isMicrosoftAuth
                ? currentUserMs?.name
                  ? currentUserMs.name.charAt(0).toUpperCase()
                  : "M"
                : auth.user?.profile?.email
                  ? auth.user.profile.email.charAt(0).toUpperCase()
                  : "C"}
            </div>

            <h2>
              ¡Bienvenido,{" "}
              {isMicrosoftAuth
                ? currentUserMs?.name
                : auth.user?.profile?.email}
              !
            </h2>
            <p className="subtitle">
              Autenticado con{" "}
              {isMicrosoftAuth ? "Microsoft Entra ID" : "AWS Cognito"}
            </p>

            <div className="user-details">
              {/* Detalles si es Microsoft */}
              {isMicrosoftAuth && (
                <>
                  <div className="detail-item">
                    <strong>Correo / Usuario:</strong>
                    <span>{currentUserMs?.username}</span>
                  </div>
                  <div className="detail-item">
                    <strong>Tenant ID:</strong>
                    <code>{currentUserMs?.tenantId}</code>
                  </div>
                </>
              )}

              {/* Detalles si es Cognito */}
              {isCognitoAuth && (
                <>
                  <div className="detail-item">
                    <strong>ID Token:</strong>
                    <code style={{ wordBreak: "break-all" }}>
                      {auth.user?.id_token?.substring(0, 20)}...
                    </code>
                  </div>
                  <div className="detail-item">
                    <strong>Access Token:</strong>
                    <code style={{ wordBreak: "break-all" }}>
                      {auth.user?.access_token?.substring(0, 20)}...
                    </code>
                  </div>
                </>
              )}
            </div>

            <hr style={{ margin: "1.5rem 0", borderColor: "#eee" }} />
            <ProtectedData />
          </div>
        ) : (
          <div className="login-container">
            {/* Tarjeta Microsoft */}
            <div className="card text-center">
              <img
                src="../public/images/microsoft-logo.webp"
                alt="Microsoft Logo"
                height={90}
              />
              <h2>Acceso Requerido</h2>
              <p className="subtitle">
                Para ingresar al sistema debes validar tus credenciales
                corporativas o institucionales.
              </p>
              <button
                className="btn btn-login btn-lg"
                onClick={handleLoginMicrosoft}
                disabled={inProgress !== InteractionStatus.None}
              >
                {inProgress !== InteractionStatus.None
                  ? "Cargando..."
                  : "Iniciar Sesión con Microsoft"}
              </button>
            </div>

            {/* Tarjeta Cognito */}
            <div className="card text-center">
              <img
                src="../public/images/cognito-logo.png"
                alt="Cognito Logo"
                height={90}
              />
              <h2>Acceso Requerido</h2>
              <p className="subtitle">
                Para ingresar al sistema debes validar tus credenciales
                corporativas o institucionales.
              </p>
              <button
                className="btn btn-login btn-lg"
                onClick={handleLoginCognito}
                disabled={inProgress !== InteractionStatus.None}
              >
                Iniciar Sesión con Cognito
              </button>
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
