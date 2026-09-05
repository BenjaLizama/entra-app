// src/App.tsx
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";
import { loginRequest } from "./authConfig";
import { ProtectedData } from "./ProtectedData";
import "./App.css";

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
      instance
        .logoutRedirect({ postLogoutRedirectUri: "/" })
        .catch((e) => console.error(e));
    }
  };

  return (
    <div className="layout">
      <header className="navbar">
        <div className="logo">
          ⚡ <span>Portal MiApp</span>
        </div>
        <div>
          {isAuthenticated ? (
            <button
              className="btn btn-logout"
              onClick={handleLogout}
              disabled={inProgress !== InteractionStatus.None}
            >
              Cerrar Sesión
            </button>
          ) : (
            <button
              className="btn btn-login"
              onClick={handleLogin}
              disabled={inProgress !== InteractionStatus.None}
            >
              Iniciar Sesión
            </button>
          )}
        </div>
      </header>

      <main className="container">
        {isAuthenticated ? (
          <div className="card">
            <div className="avatar">
              {currentUser?.name
                ? currentUser.name.charAt(0).toUpperCase()
                : "U"}
            </div>
            <h2>¡Bienvenido, {currentUser?.name || "Usuario"}!</h2>
            <p className="subtitle">Autenticado con Microsoft Entra ID</p>

            <div className="user-details">
              <div className="detail-item">
                <strong>Correo / Usuario:</strong>
                <span>{currentUser?.username}</span>
              </div>
              <div className="detail-item">
                <strong>Tenant ID:</strong>
                <code>{currentUser?.tenantId}</code>
              </div>
            </div>

            {/* Integración del componente protegido + Interceptor/API */}
            <hr style={{ margin: "1.5rem 0", borderColor: "#eee" }} />
            <ProtectedData />
          </div>
        ) : (
          <div className="card text-center">
            <h2>Acceso Requerido</h2>
            <p className="subtitle">
              Para ingresar al sistema debes validar tus credenciales
              corporativas o institucionales.
            </p>
            <button
              className="btn btn-login btn-lg"
              onClick={handleLogin}
              disabled={inProgress !== InteractionStatus.None}
            >
              {inProgress !== InteractionStatus.None
                ? "Cargando..."
                : "Iniciar Sesión con Microsoft"}
            </button>
          </div>
        )}
      </main>
    </div>
  );
}
