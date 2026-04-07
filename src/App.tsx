import { useState, useEffect, useCallback, useRef } from 'react'

// ============================================================
// TYPES
// ============================================================
interface EvidenciaFile {
  id: string; name: string; folio: string; size: number;
  modified: string; mimeType: string; driveId: string;
  downloadUrl: string | null; webUrl: string | null;
}

interface ScanProgress { total: number; current: number; current_folio?: string; }
interface BatchProgress { total: number; current: number; }
interface AnalysisResult {
  legible: boolean; tipo_documento: string; fecha: string | null;
  monto: number | null; referencia: string | null; cliente_documento: string | null;
  banco_emisor: string | null; folio_presente: boolean | null;
  observaciones: string; semaforo: 'verde' | 'amarillo' | 'rojo';
}
interface User { displayName: string; mail: string; }

// ============================================================
// CONFIG
// ============================================================
const CONFIG = {
  clientId: 'b271f29f-65f7-476e-a272-63669bdfd85e',
  tenantId: '746b050c-a1ff-45b9-9858-e142490982b7',
  siteHostname: 'cisurft.sharepoint.com',
  sitePath: '/sites/PlaneacionFinanciera',
  redirectUri: window.location.origin,
  scopes: ['Files.Read', 'Files.Read.All', 'offline_access', 'User.Read'],
  // Lee de Vercel. Si pruebas localmente, asegúrate de tener un archivo .env
  geminiApiKey: import.meta.env.VITE_GEMINI_API_KEY || '' 
}

// ============================================================
// AUTH & HELPERS
// ============================================================
async function pkceLogin() {
  const array = crypto.getRandomValues(new Uint8Array(32));
  const verifier = Array.from(array).map(b => b.toString(16).padStart(2, '0')).join('');
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
  const challenge = btoa(String.fromCharCode(...new Uint8Array(digest))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
  sessionStorage.setItem('pkce_verifier', verifier);
  sessionStorage.setItem('oauth_state', Math.random().toString(36).slice(2));
  const params = new URLSearchParams({
    client_id: CONFIG.clientId, response_type: 'code', redirect_uri: CONFIG.redirectUri,
    scope: CONFIG.scopes.join(' '), state: sessionStorage.getItem('oauth_state') ?? '',
    code_challenge: challenge, code_challenge_method: 'S256', response_mode: 'query',
  });
  window.location.href = `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/authorize?${params}`;
}

async function pkceExchange(code: string): Promise<string> {
  const body = new URLSearchParams({
    client_id: CONFIG.clientId, grant_type: 'authorization_code', code,
    redirect_uri: CONFIG.redirectUri, code_verifier: sessionStorage.getItem('pkce_verifier') ?? '',
    scope: CONFIG.scopes.join(' '),
  });
  const res = await fetch(`https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`, { method: 'POST', body });
  const data = await res.json();
  if (!data.access_token) throw new Error(data.error_description ?? 'Auth failed');
  sessionStorage.setItem('ms_token', JSON.stringify({ token: data.access_token, expires: Date.now() + (data.expires_in - 60) * 1000 }));
  return data.access_token;
}

const getStoredToken = () => {
  const raw = sessionStorage.getItem('ms_token');
  if (!raw) return null;
  const { token, expires } = JSON.parse(raw);
  return Date.now() < expires ? token : null;
};

const fmtMXN = (n: number | null) => n ? new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' }).format(n) : '$0.00';
const SEM: Record<string, any> = {
  verde: { bg: '#d1fae5', color: '#064e3b', label: '✓ OK' },
  amarillo: { bg: '#fef3c7', color: '#78350f', label: '⚠ Revisar' },
  rojo: { bg: '#fee2e2', color: '#7f1d1d', label: '✗ Alerta' },
};

// ============================================================
// MAIN COMPONENT
// ============================================================
export default function App() {
  const [token, setToken] = useState<string | null>(null);
  const [user, setUser] = useState<User | null>(null);
  const [driveId, setDriveId] = useState<string | null>(null);
  const [files, setFiles] = useState<EvidenciaFile[]>([]);
  const [analyses, setAnalyses] = useState<Record<string, AnalysisResult>>({});
  const [scanning, setScanning] = useState(false);
  const [scanProgress, setScanProgress] = useState<ScanProgress | null>(null);
  const [batchProgress, setBatchProgress] = useState<BatchProgress | null>(null);
  const [analyzingIds, setAnalyzingIds] = useState<Set<string>>(new Set());
  const [error, setError] = useState<string | null>(null);
  const [basePath, setBasePath] = useState('Ventas');
  const [search, setSearch] = useState('');
  const [preview, setPreview] = useState<EvidenciaFile | null>(null);
  const stopRef = useRef(false);

  useEffect(() => {
    const code = new URLSearchParams(window.location.search).get('code');
    if (code) {
      window.history.replaceState({}, '', window.location.pathname);
      pkceExchange(code).then(setToken).catch(e => setError(String(e)));
    } else {
      const t = getStoredToken(); if (t) setToken(t);
    }
  }, []);

  useEffect(() => {
    if (!token) return;
    const g = (url: string) => fetch(url, { headers: { Authorization: `Bearer ${token}` } }).then(r => r.json());
    g('https://graph.microsoft.com/v1.0/me').then(setUser).catch(() => {});
    g(`https://graph.microsoft.com/v1.0/sites/${CONFIG.siteHostname}:${CONFIG.sitePath}`)
      .then(site => g(`https://graph.microsoft.com/v1.0/sites/${site.id}/drives`))
      .then(drives => {
        const ev = drives.value.find((d: any) => d.name === 'Evidencias' || d.webUrl.includes('/Evidencias'));
        setDriveId(ev?.id || drives.value[0]?.id);
      }).catch(e => setError('Error SharePoint: ' + e.message));
  }, [token]);

  const scan = async () => {
    if (!token || !driveId) return;
    setScanning(true); setFiles([]); setAnalyses({});
    try {
      const g = (url: string) => fetch(url, { headers: { Authorization: `Bearer ${token}` } }).then(r => r.json());
      const folders = await g(`https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${basePath}:/children`);
      const folioFolders = folders.value.filter((f: any) => f.folder);
      
      const found: EvidenciaFile[] = [];
      setScanProgress({ total: folioFolders.length, current: 0 });

      for (let i = 0; i < folioFolders.length; i++) {
        setScanProgress({ total: folioFolders.length, current: i + 1, current_folio: folioFolders[i].name });
        const items = await g(`https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${basePath}/${folioFolders[i].name}:/children`);
        items.value.forEach((item: any) => {
          if (item.file && /\.(pdf|jpg|jpeg|png)$/i.test(item.name)) {
            found.push({
              id: item.id, name: item.name, folio: folioFolders[i].name, size: item.size,
              modified: item.lastModifiedDateTime, mimeType: item.file.mimeType, driveId: driveId!,
              downloadUrl: item['@microsoft.graph.downloadUrl'], webUrl: item.webUrl
            });
          }
        });
      }
      setFiles(found);
    } catch (e: any) { setError(e.message); }
    setScanning(false); setScanProgress(null);
  };

  const analyzeOne = async (file: EvidenciaFile) => {
    if (!CONFIG.geminiApiKey) {
      setError("Falta API Key de Gemini en Vercel (VITE_GEMINI_API_KEY)");
      return;
    }
    setAnalyzingIds(prev => new Set(prev).add(file.id));
    try {
      const proxyUrl = `/api/file?${new URLSearchParams({ driveId: file.driveId, fileId: file.id, token: token! })}`;
      const blob = await fetch(proxyUrl).then(r => { if(!r.ok) throw new Error("Error Proxy"); return r.blob(); });
      const base64 = await new Promise<string>((res) => {
        const r = new FileReader(); r.onloadend = () => res((r.result as string).split(',')[1]); r.readAsDataURL(blob);
      });

      const geminiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${CONFIG.geminiApiKey}`;
      const prompt = `Eres auditor financiero. Analiza esta evidencia del folio "${file.folio}". Extrae y responde SOLO con JSON válido:
      {"legible": true, "tipo_documento": "transferencia|ticket_caja|factura|remision|comprobante_pago", "fecha": "DD/MM/YYYY", "monto": 0.0, "referencia": "string", "cliente_documento": "string", "banco_emisor": "string", "folio_presente": true, "observaciones": "string", "semaforo": "verde|amarillo|rojo"}`;

      const res = await fetch(geminiUrl, {
        method: 'POST',
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: file.mimeType, data: base64 } }] }],
          generationConfig: { response_mime_type: "application/json", temperature: 0.1 }
        })
      });
      const data = await res.json();
      const result = JSON.parse(data.candidates[0].content.parts[0].text.trim());
      setAnalyses(prev => ({ ...prev, [file.id]: result }));
    } catch (e: any) {
      setAnalyses(prev => ({ ...prev, [file.id]: { semaforo: 'rojo', observaciones: e.message } as any }));
    }
    setAnalyzingIds(prev => { const s = new Set(prev); s.delete(file.id); return s; });
  };

  const analyzeAll = async () => {
    stopRef.current = false;
    const pending = files.filter(f => !analyses[f.id]);
    setBatchProgress({ total: pending.length, current: 0 });
    for (let i = 0; i < pending.length; i++) {
      if (stopRef.current) break;
      await analyzeOne(pending[i]);
      setBatchProgress({ total: pending.length, current: i + 1 });
      if (i % 5 === 4) await new Promise(r => setTimeout(r, 1000));
    }
    setBatchProgress(null);
  };

  const stats = {
    total: files.length,
    analizados: Object.keys(analyses).length,
    verde: Object.values(analyses).filter(a => a.semaforo === 'verde').length,
    amarillo: Object.values(analyses).filter(a => a.semaforo === 'amarillo').length,
    rojo: Object.values(analyses).filter(a => a.semaforo === 'rojo').length,
    monto: Object.values(analyses).reduce((s, a) => s + (a.monto || 0), 0)
  };

  if (!token) return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#0f172a' }}>
      <button onClick={pkceLogin} style={{ padding: '15px 30px', background: '#0078d4', color: '#fff', border: 'none', borderRadius: 10, fontWeight: 700, cursor: 'pointer' }}>🔐 Conectar SharePoint IQ</button>
    </div>
  );

  return (
    <div style={{ background: '#f1f5f9', minHeight: '100vh', fontFamily: 'sans-serif' }}>
      <div style={{ background: '#0f172a', padding: '15px 25px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', color: '#fff' }}>
        <div><strong style={{ fontSize: 18 }}>EvidenciasIQ</strong> <span style={{ opacity: 0.6, fontSize: 12 }}>| {user?.displayName}</span></div>
        <div style={{ display: 'flex', gap: 10 }}>
          {files.length > 0 && <button onClick={analyzeAll} style={{ background: '#7c3aed', color: '#fff', border: 'none', padding: '8px 15px', borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>🤖 Analizar todo ({stats.total - stats.analizados})</button>}
          <button onClick={() => { sessionStorage.clear(); window.location.href = window.location.origin; }} style={{ background: 'rgba(255,255,255,0.1)', color: '#fff', border: '1px solid #334155', padding: '8px 15px', borderRadius: 8, fontSize: 12, cursor: 'pointer' }}>Salir</button>
        </div>
      </div>

      <div style={{ padding: 25, maxWidth: 1300, margin: '0 auto' }}>
        {error && <div style={{ background: '#fee2e2', color: '#b91c1c', padding: 15, borderRadius: 10, marginBottom: 20, fontSize: 13 }}>⚠️ {error}</div>}

        <div style={{ background: '#fff', padding: 20, borderRadius: 15, border: '1px solid #e2e8f0', marginBottom: 20 }}>
          <div style={{ display: 'flex', gap: 12, alignItems: 'flex-end' }}>
            <div style={{ flex: 1 }}>
              <label style={{ fontSize: 11, fontWeight: 800, color: '#64748b' }}>Carpeta en SharePoint</label>
              <input value={basePath} onChange={e => setBasePath(e.target.value)} style={{ width: '100%', padding: '10px', borderRadius: 8, border: '1px solid #e2e8f0', marginTop: 5 }} />
            </div>
            <button onClick={scan} disabled={scanning} style={{ background: '#1e40af', color: '#fff', border: 'none', padding: '12px 25px', borderRadius: 10, fontWeight: 700, cursor: 'pointer' }}>{scanning ? 'Escaneando...' : '📂 Escanear'}</button>
          </div>
          {(scanProgress || batchProgress) && (
            <div style={{ marginTop: 15 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: '#64748b', marginBottom: 5 }}>
                <span>{scanProgress ? `Escaneando: ${scanProgress.current_folio}` : 'Analizando con Gemini IA...'}</span>
                <span>{(scanProgress || batchProgress)?.current}/{(scanProgress || batchProgress)?.total}</span>
              </div>
              <div style={{ background: '#e2e8f0', height: 8, borderRadius: 10 }}><div style={{ background: scanProgress ? '#3b82f6' : '#7c3aed', height: '100%', borderRadius: 10, width: `${((scanProgress || batchProgress)!.current / (scanProgress || batchProgress)!.total) * 100}%`, transition: '0.3s' }} /></div>
            </div>
          )}
        </div>

        {files.length > 0 && (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(6, 1fr)', gap: 15, marginBottom: 20 }}>
            {[ ['Total', stats.total, '#1e40af'], ['Analizados', stats.analizados, '#0f766e'], ['✓ OK', stats.verde, '#064e3b'], ['⚠ Revisar', stats.amarillo, '#92400e'], ['✗ Alertas', stats.rojo, '#991b1b'], ['Monto MXN', fmtMXN(stats.monto), '#4c1d95'] ]
              .map(([lbl, val, col]) => (
                <div key={lbl as string} style={{ background: '#fff', padding: 15, borderRadius: 15, border: '1px solid #e2e8f0', textAlign: 'center' }}>
                  <div style={{ fontSize: 10, fontWeight: 800, color: '#94a3b8' }}>{lbl}</div>
                  <div style={{ fontSize: 18, fontWeight: 900, color: col as string, marginTop: 5 }}>{val}</div>
                </div>
              ))}
          </div>
        )}

        <div style={{ display: 'grid', gridTemplateColumns: preview ? '1fr 400px' : '1fr', gap: 20 }}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <input placeholder="🔍 Filtrar por nombre o folio..." value={search} onChange={e => setSearch(e.target.value)} style={{ padding: '12px', borderRadius: 10, border: '1px solid #e2e8f0', marginBottom: 10 }} />
            {files.filter(f => f.folio.toLowerCase().includes(search.toLowerCase()) || f.name.toLowerCase().includes(search.toLowerCase())).map(f => {
              const a = analyses[f.id]; const isAn = analyzingIds.has(f.id);
              const s = a ? SEM[a.semaforo] : null;
              return (
                <div key={f.id} onClick={() => setPreview(f)} style={{ background: '#fff', padding: 15, borderRadius: 12, border: `2px solid ${preview?.id === f.id ? '#3b82f6' : '#e2e8f0'}`, cursor: 'pointer' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontWeight: 800, fontSize: 14 }}>{f.folio} <span style={{ fontWeight: 400, color: '#94a3b8', fontSize: 12 }}>{f.name}</span></div>
                      {a && (
                        <div style={{ marginTop: 10, padding: 10, background: '#f8fafc', borderRadius: 10, fontSize: 12 }}>
                          <div style={{ display: 'flex', gap: 10, marginBottom: 8 }}>
                            <span style={{ background: s.bg, color: s.color, padding: '2px 8px', borderRadius: 20, fontWeight: 700 }}>{s.label}</span>
                            <span style={{ background: '#e0e7ff', color: '#3730a3', padding: '2px 8px', borderRadius: 20, fontWeight: 700 }}>{a.tipo_documento}</span>
                          </div>
                          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 5 }}>
                            {a.monto && <div>Monto: <strong>{fmtMXN(a.monto)}</strong></div>}
                            {a.fecha && <div>Fecha: <strong>{a.fecha}</strong></div>}
                          </div>
                          {a.observaciones && <div style={{ marginTop: 5, fontStyle: 'italic', color: '#64748b' }}>{a.observaciones}</div>}
                        </div>
                      )}
                    </div>
                    <button onClick={e => { e.stopPropagation(); analyzeOne(f); }} disabled={isAn} style={{ background: a ? '#f1f5f9' : '#7c3aed', color: a ? '#64748b' : '#fff', border: 'none', padding: '8px 12px', borderRadius: 8, cursor: 'pointer' }}>{isAn ? '...' : a ? '↻' : '🤖'}</button>
                  </div>
                </div>
              );
            })}
          </div>

          {preview && (
            <div style={{ background: '#fff', borderRadius: 15, border: '1px solid #e2e8f0', position: 'sticky', top: 20, height: '85vh', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              <div style={{ padding: 15, borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <strong style={{ fontSize: 14 }}>{preview.folio}</strong>
                <button onClick={() => setPreview(null)} style={{ border: 'none', background: 'none', fontSize: 20, cursor: 'pointer' }}>✕</button>
              </div>
              <div style={{ flex: 1, background: '#f8fafc' }}>
                <iframe src={`/api/file?${new URLSearchParams({ driveId: preview.driveId, fileId: preview.id, token: token! })}`} style={{ width: '100%', height: '100%', border: 'none' }} />
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
