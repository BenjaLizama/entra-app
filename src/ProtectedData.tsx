// src/ProtectedData.tsx
import { useState } from 'react';
import { AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react';
import { useApi } from './useApi';

export function ProtectedData() {
  const { fetchWithToken } = useApi();
  const [data, setData] = useState<any>(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleFetchData = async () => {
    setLoading(true);
    setError(null);
    try {
      // Ejemplo: Llamar a Microsoft Graph API o a tu Backend protegido
      const res = await fetchWithToken('https://graph.microsoft.com/v1.0/me');
      if (!res.ok) throw new Error(`Error en la API: ${res.statusText}`);
      
      const json = await res.json();
      setData(json);
    } catch (err: any) {
      setError(err.message || 'Error al obtener datos');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div>
      {/* Equivalente a MsalGuard cuando el usuario está logueado */}
      <AuthenticatedTemplate>
        <div style={{ marginTop: '1.5rem', textAlign: 'left' }}>
          <h3>Consulta Protegida a API (Graph / Backend)</h3>
          <p style={{ fontSize: '0.85rem', color: '#666', marginBottom: '1rem' }}>
            Esta sección prueba la obtención e inyección del Bearer Token.
          </p>
          <button className="btn btn-login" onClick={handleFetchData} disabled={loading}>
            {loading ? 'Consultando...' : 'Obtener Datos del Usuario vía API'}
          </button>

          {error && <p style={{ color: 'red', marginTop: '1rem' }}>{error}</p>}

          {data && (
            <pre style={{
              backgroundColor: '#2d2d2d',
              color: '#67cdaa',
              padding: '1rem',
              borderRadius: '6px',
              marginTop: '1rem',
              fontSize: '0.8rem',
              overflowX: 'auto'
            }}>
              {JSON.stringify(data, null, 2)}
            </pre>
          )}
        </div>
      </AuthenticatedTemplate>

      {/* Equivalente a MsalGuard cuando NO está logueado */}
      <UnauthenticatedTemplate>
        <p style={{ color: '#d9534f', marginTop: '1rem' }}>
          ⚠️ Acceso denegado. Debes iniciar sesión para consultar este recurso protegido.
        </p>
      </UnauthenticatedTemplate>
    </div>
  );
}