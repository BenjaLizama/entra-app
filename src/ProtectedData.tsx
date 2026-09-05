// src/ProtectedData.tsx
import { useState } from "react";
import { useIsAuthenticated } from "@azure/msal-react";
import { useAuth } from "react-oidc-context";
import { useApi } from "./useApi";

export function ProtectedData() {
  const { fetchWithToken } = useApi();

  // Evaluamos ambas sesiones
  const isMicrosoftAuth = useIsAuthenticated();
  const auth = useAuth();
  const isCognitoAuth = auth.isAuthenticated;

  const isAnyUserAuthenticated = isMicrosoftAuth || isCognitoAuth;

  const [data, setData] = useState<any>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleFetchData = async () => {
    setLoading(true);
    setError(null);
    try {
      if (isMicrosoftAuth) {
        // Petición exclusiva para Microsoft usando tu hook
        const res = await fetchWithToken("https://graph.microsoft.com/v1.0/me");
        if (!res.ok)
          throw new Error(`Error en la API de MS: ${res.statusText}`);
        const json = await res.json();
        setData(json);
      } else if (isCognitoAuth) {
        // Lógica para Cognito: Puedes hacer un fetch normal a tu backend aquí
        // O simplemente mostrar la data que ya tienes en el token como ejemplo:
        setData({
          origen: "AWS Cognito",
          mensaje:
            "Aquí puedes hacer un fetch() a tu API enviando el token de Cognito",
          perfil_usuario: auth.user?.profile,
          token_usado: auth.user?.access_token?.substring(0, 30) + "...",
        });
      }
    } catch (err: any) {
      setError(err.message || "Error al obtener datos");
    } finally {
      setLoading(false);
    }
  };

  // Si no hay ninguna sesión activa, mostramos el mensaje de error directamente
  if (!isAnyUserAuthenticated) {
    return (
      <p style={{ color: "#d9534f", marginTop: "1rem" }}>
        ⚠️ Acceso denegado. Debes iniciar sesión para consultar este recurso
        protegido.
      </p>
    );
  }

  // Si hay alguna sesión (Microsoft o Cognito), mostramos la interfaz
  return (
    <div style={{ marginTop: "1.5rem", textAlign: "left" }}>
      <h3>Consulta Protegida de Datos</h3>
      <p style={{ fontSize: "0.85rem", color: "#666", marginBottom: "1rem" }}>
        Se obtendrá información dependiendo de si usaste Microsoft o Cognito.
      </p>

      <button
        className="btn btn-login"
        onClick={handleFetchData}
        disabled={loading}
      >
        {loading ? "Consultando..." : "Obtener Datos del Usuario"}
      </button>

      {error && <p style={{ color: "red", marginTop: "1rem" }}>{error}</p>}

      {data && (
        <pre
          style={{
            backgroundColor: "#2d2d2d",
            color: "#67cdaa",
            padding: "1rem",
            borderRadius: "6px",
            marginTop: "1rem",
            fontSize: "0.8rem",
            overflowX: "auto",
          }}
        >
          {JSON.stringify(data, null, 2)}
        </pre>
      )}
    </div>
  );
}
