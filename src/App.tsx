import { useState, useEffect, useCallback, useRef } from 'react'

// ============================================================
// TYPES
// ============================================================
interface EvidenciaFile {
  id: string
  name: string
  folio: string
  size: number
  modified: string
  mimeType: string
  driveId: string
  downloadUrl: string | null
  webUrl: string | null
}

interface ScanProgress {
  total: number
  current: number
  current_folio?: string
}

interface BatchProgress {
  total: number
  current: number
}

interface AnalysisResult {
  legible: boolean
  tipo_documento: string
  fecha: string | null
  monto: number | null
  referencia: string | null
  cliente_documento: string | null
  banco_emisor: string | null
  folio_presente: boolean | null
  observaciones: string
  semaforo: 'verde' | 'amarillo' | 'rojo'
}

interface User {
  displayName: string
  mail: string
}

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
  // Si la variable de Vercel no carga, puedes pegarla aquí temporalmente entre comillas
  geminiApiKey: import.meta.env.VITE_GEMINI_API_KEY || '' 
}

// ============================================================
// AUTH — OAuth2 PKCE
// ============================================================
async function pkceLogin(): Promise<void> {
  const array = crypto.getRandomValues(new Uint8Array(32))
  const verifier = Array.from(array).map(b => b.toString(16).padStart(2, '0')).join('')
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier))
  const challenge = btoa(String.fromCharCode(...new Uint8Array(digest)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '')
  sessionStorage.setItem('pkce_verifier', verifier)
  sessionStorage.setItem('oauth_state', Math.random().toString(36).slice(2))
  const params = new URLSearchParams({
    client_id: CONFIG.clientId,
    response_type: 'code',
    redirect_uri: CONFIG.redirectUri,
    scope: CONFIG.scopes.join(' '),
    state: sessionStorage.getItem('oauth_state') ?? '',
    code_challenge: challenge,
    code_challenge_method: 'S256',
    response_mode: 'query',
  })
  window.location.href = `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/authorize?${params}`
}

async function pkceExchange(code: string): Promise<string> {
  const body = new URLSearchParams({
    client_id: CONFIG.clientId,
    grant_type: 'authorization_code',
    code,
    redirect_uri: CONFIG.redirectUri,
    code_verifier: sessionStorage.getItem('pkce_verifier') ?? '',
    scope: CONFIG.scopes.join(' '),
  })
  const res = await fetch(
    `https://login.microsoftonline.com/${CONFIG.tenantId}/oauth2/v2.0/token`,
    { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body }
  )
  const data = await res.json()
  if (!data.access_token) throw new Error(data.error_description ?? 'Auth failed')
  sessionStorage.setItem('ms_token', JSON.stringify({
    token: data.access_token,
    expires: Date.now() + (data.expires_in - 60) * 1000,
  }))
  return data.access_token
}

function getStoredToken(): string | null {
  const raw = sessionStorage.getItem('ms_token')
  if (!raw) return null
  const { token, expires } = JSON.parse(raw)
  return Date.now() < expires ? token : null
}

function clearAuth(): void {
  sessionStorage.removeItem('ms_token')
  sessionStorage.removeItem('pkce_verifier')
  sessionStorage.removeItem('oauth_state')
}

// ============================================================
// GRAPH API
// ============================================================
async function graphGet(url: string, token: string): Promise<any> {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } })
  if (!res.ok) throw new Error(`Graph ${res.status}: ${url}`)
  return res.json()
}

async function getSiteId(token: string): Promise<string> {
  const data = await graphGet(
    `https://graph.microsoft.com/v1.0/sites/${CONFIG.siteHostname}:${CONFIG.sitePath}`,
    token
  )
  return data.id
}

async function getEvidenciasDriveId(token: string, siteId: string): Promise<string> {
  const data = await graphGet(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
    token
  )
  const drives: any[] = data.value ?? []
  const ev = drives.find(
    (d: any) => d.name === 'Evidencias' || d.webUrl?.toLowerCase().includes('/evidencias')
  )
  return ev?.id ?? drives[0]?.id ?? ''
}

async function listAllChildren(token: string, driveId: string, path: string): Promise<any[]> {
  const select = '$select=id,name,size,file,folder,lastModifiedDateTime,webUrl,@microsoft.graph.downloadUrl'
  const firstUrl = path
    ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(path)}:/children?$top=999&${select}`
    : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$top=999&${select}`
  const all: any[] = []
  let url: string | null = firstUrl
  while (url) {
    const data = await graphGet(url, token)
    all.push(...(data.value ?? []))
    url = data['@odata.nextLink'] ?? null
  }
  return all
}

async function scanFolios(
  token: string,
  driveId: string,
  basePath: string,
  onProgress: (p: ScanProgress) => void
): Promise<EvidenciaFile[]> {
  const folders = await listAllChildren(token, driveId, basePath)
  const folioFolders = folders.filter((f: any) => f.folder)
  onProgress({ total: folioFolders.length, current: 0 })
  const results: EvidenciaFile[] = []

  for (let i = 0; i < folioFolders.length; i++) {
    const folder = folioFolders[i]
    onProgress({ total: folioFolders.length, current: i + 1, current_folio: folder.name })
    try {
      const items = await listAllChildren(token, driveId, `${basePath}/${folder.name}`)
      for (const item of items) {
        if (item.file && /\.(pdf|jpg|jpeg|png|webp)$/i.test(item.name)) {
          results.push({
            id: item.id,
            name: item.name,
            folio: folder.name,
            size: item.size ?? 0,
            modified: item.lastModifiedDateTime ?? '',
            mimeType: item.file?.mimeType ?? '',
            driveId,
            downloadUrl: item['@microsoft.graph.downloadUrl'] ?? null,
            webUrl: item.webUrl ?? null,
          })
        }
      }
    } catch { /* skip inaccessible folder */ }
  }
  return results
}

function proxyUrl(token: string, driveId: string, fileId: string): string {
  const params = new URLSearchParams({ driveId, fileId, token })
  return `/api/file?${params}`
}

async function getFileBase64(token: string, driveId: string, fileId: string): Promise<string> {
  const res = await fetch(proxyUrl(token, driveId, fileId))
  if (!res.ok) throw new Error(`Proxy ${res.status}`)
  const blob = await res.blob()
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onloadend = () => resolve((reader.result as string).split(',')[1])
    reader.onerror = reject
    reader.readAsDataURL(blob)
  })
}

// ============================================================
// GEMINI API (Ajustado para evitar Error 413)
// ============================================================
async function analyzeWithGemini(
  base64: string,
  mimeType: string,
  folio: string
): Promise<AnalysisResult> {
  const apiKey = CONFIG.geminiApiKey;
  if (!apiKey || apiKey.length < 10) throw new Error("API Key no configurada");

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
  const prompt = `Eres auditor financiero de "Flor de Tabasco". Analiza esta evidencia del folio "${folio}".
Extrae y responde SOLO con JSON válido:
{
  "legible": true,
  "tipo_documento": "transferencia|ticket_caja|factura|remision|foto_entrega|comprobante_pago|otro",
  "fecha": "DD/MM/YYYY o null",
  "monto": 1234.56,
  "referencia": "numero de operacion o null",
  "cliente_documento": "nombre del cliente o null",
  "banco_emisor": "banco o null",
  "folio_presente": true,
  "observaciones": "una línea máximo",
  "semaforo": "verde|amarillo|rojo"
}`;

  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mimeType, data: base64 } }] }],
      generationConfig: { response_mime_type: "application/json", temperature: 0.1 }
    })
  });

  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error?.message || "Error en Gemini API");
  }

  const data = await res.json();
  const text = data.candidates[0].content.parts[0].text;
  return JSON.parse(text.trim());
}

// ============================================================
// HELPERS & UI
// ============================================================
const fmtMXN = (n: number | null) => n != null ? new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' }).format(n) : '$0.00';
const fmtKB = (b: number) => b > 1024 * 1024 ? `${(b / 1024 / 1024).toFixed(1)} MB` : `${(b / 1024).toFixed(0)} KB`;
const SEM: Record<string, any> = {
  verde:    { bg: '#d1fae5', color: '#064e3b', label: '✓ OK' },
  amarillo: { bg: '#fef3c7', color: '#78350f', label: '⚠ Revisar' },
  rojo:     { bg: '#fee2e2', color: '#7f1d1d', label: '✗ Alerta' },
};

function AnalysisCard({ r }: { r: AnalysisResult }) {
  const s = SEM[r.semaforo] ?? SEM.amarillo
  return (
    <div style={{ marginTop: 8, padding: 10, background: '#f8faff', borderRadius: 8, border: '1px solid #e2e8f0', fontSize: 12 }}>
      <div style={{ display: 'flex', gap: 6, marginBottom: 6 }}>
        <span style={{ background: s.bg, color: s.color, fontSize: 11, fontWeight: 700, padding: '2px 8px', borderRadius: 20 }}>{s.label}</span>
        <span style={{ background: '#dbeafe', color: '#1e3a8a', fontSize: 11, fontWeight: 600, padding: '2px 8px', borderRadius: 20 }}>{r.tipo_documento}</span>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '3px 12px', color: '#475569' }}>
        {[['Fecha', r.fecha], ['Monto', r.monto ? fmtMXN(r.monto) : null], ['Ref', r.referencia], ['Cliente', r.cliente_documento]].map(([k, v]) => 
          v && <div key={k}><span style={{ color: '#94a3b8' }}>{k}: </span><strong>{v}</strong></div>
        )}
      </div>
      {r.observaciones && <p style={{ margin: '6px 0 0', color: '#64748b', fontStyle: 'italic', fontSize: 11 }}>{r.observaciones}</p>}
    </div>
  )
}

// ============================================================
// MAIN APP
// ============================================================
export default function App() {
  const [token, setToken] = useState<string | null>(null)
  const [user, setUser] = useState<User | null>(null)
  const [driveId, setDriveId] = useState<string | null>(null)
  const [files, setFiles] = useState<EvidenciaFile[]>([])
  const [analyses, setAnalyses] = useState<Record<string, AnalysisResult>>({})
  const [scanning, setScanning] = useState(false)
  const [scanProgress, setScanProgress] = useState<ScanProgress | null>(null)
  const [batchProgress, setBatchProgress] = useState<BatchProgress | null>(null)
  const [analyzingIds, setAnalyzingIds] = useState<Set<string>>(new Set())
  const [error, setError] = useState<string | null>(null)
  const [basePath, setBasePath] = useState('Ventas')
  const [search, setSearch] = useState('')
  const [preview, setPreview] = useState<EvidenciaFile | null>(null)
  const stopRef = useRef(false)

  useEffect(() => {
    const code = new URLSearchParams(window.location.search).get('code')
    if (code) {
      window.history.replaceState({}, '', window.location.pathname)
      pkceExchange(code).then(setToken).catch(e => setError(String(e)))
    } else {
      const t = getStoredToken(); if (t) setToken(t)
    }
  }, [])

  useEffect(() => {
    if (!token) return
    graphGet('https://graph.microsoft.com/v1.0/me', token).then(setUser).catch(() => {})
    getSiteId(token).then(id => getEvidenciasDriveId(token, id)).then(setDriveId).catch(e => setError(String(e)))
  }, [token])

  const analyzeOne = async (file: EvidenciaFile) => {
    setAnalyzingIds(prev => new Set(prev).add(file.id))
    try {
      const base64 = await getFileBase64(token!, file.driveId, file.id)
      const res = await analyzeWithGemini(base64, file.mimeType || 'image/jpeg', file.folio)
      setAnalyses(prev => ({ ...prev, [file.id]: res }))
    } catch (e: any) {
      setAnalyses(prev => ({ ...prev, [file.id]: { semaforo: 'rojo', observaciones: e.message } as any }))
    }
    setAnalyzingIds(prev => { const s = new Set(prev); s.delete(file.id); return s })
  }

  const analyzeAll = async () => {
    stopRef.current = false
    const pending = files.filter(f => !analyses[f.id])
    setBatchProgress({ total: pending.length, current: 0 })
    for (let i = 0; i < pending.length; i++) {
      if (stopRef.current) break
      await analyzeOne(pending[i])
      setBatchProgress({ total: pending.length, current: i + 1 })
      if (i % 5 === 4) await new Promise(r => setTimeout(r, 1000))
    }
    setBatchProgress(null)
  }

  const stats = {
    total: files.length,
    analizados: Object.keys(analyses).length,
    verde: Object.values(analyses).filter(a => a.semaforo === 'verde').length,
    amarillo: Object.values(analyses).filter(a => a.semaforo === 'amarillo').length,
    rojo: Object.values(analyses).filter(a => a.semaforo === 'rojo').length,
    monto: Object.values(analyses).reduce((s, a) => s + (a.monto || 0), 0)
  }

  if (!token) return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#0f172a' }}>
      <button onClick={pkceLogin} style={{ padding: '15px 30px', background: '#0078d4', color: '#fff', border: 'none', borderRadius: 10, fontWeight: 700, cursor: 'pointer' }}>🔐 Conectar SharePoint IQ</button>
    </div>
  )

  return (
    <div style={{ background: '#f1f5f9', minHeight: '100vh', fontFamily: 'sans-serif' }}>
      {/* HEADER */}
      <div style={{ background: '#0f172a', padding: '15px 25px', display: 'flex', justifyContent: 'space-between', alignItems: 'center', color: '#fff' }}>
        <div><strong style={{ fontSize: 18 }}>EvidenciasIQ</strong> <span style={{ opacity: 0.6, fontSize: 12 }}>| {user?.displayName}</span></div>
        <div style={{ display: 'flex', gap: 10 }}>
          {files.length > 0 && <button onClick={analyzeAll} style={{ background: '#7c3aed', color: '#fff', border: 'none', padding: '8px 15px', borderRadius: 8, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}>🤖 Analizar todo ({stats.total - stats.analizados})</button>}
          <button onClick={() => { sessionStorage.clear(); window.location.href = window.location.origin; }} style={{ background: 'rgba(255,255,255,0.1)', color: '#fff', border: '1px solid #334155', padding: '8px 15px', borderRadius: 8, fontSize: 12, cursor: 'pointer' }}>Salir</button>
        </div>
      </div>

      <div style={{ padding: 25, maxWidth: 1300, margin: '0 auto' }}>
        {/* SCAN BAR */}
        <div style={{ background: '#fff', padding: 20, borderRadius: 15, border: '1px solid #e2e8f0', marginBottom: 20 }}>
          <div style={{ display: 'flex', gap: 12, alignItems: 'flex-end' }}>
            <div style={{ flex: 1 }}>
              <label style={{ fontSize: 11, fontWeight: 800, color: '#64748b', textTransform: 'uppercase' }}>Carpeta en SharePoint</label>
              <input value={basePath} onChange={e => setBasePath(e.target.value)} style={{ width: '100%', padding: '10px', borderRadius: 8, border: '1px solid #e2e8f0', marginTop: 5 }} />
            </div>
            <button onClick={async () => { setScanning(true); setFiles(await scanFolios(token, driveId!, basePath, setScanProgress)); setScanning(false); setScanProgress(null); }} style={{ background: '#1e40af', color: '#fff', border: 'none', padding: '12px 25px', borderRadius: 10, fontWeight: 700, cursor: 'pointer' }}>{scanning ? 'Escaneando...' : '📂 Escanear Carpetas'}</button>
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

        {/* KPIs CARDS */}
        {files.length > 0 && (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(6, 1fr)', gap: 15, marginBottom: 20 }}>
            {[ ['Total', stats.total, '#1e40af'], ['Analizados', stats.analizados, '#0f766e'], ['✓ OK', stats.verde, '#064e3b'], ['⚠ Revisar', stats.amarillo, '#92400e'], ['✗ Alertas', stats.rojo, '#991b1b'], ['Monto MXN', fmtMXN(stats.monto), '#4c1d95'] ]
              .map(([lbl, val, col]) => (
                <div key={lbl as string} style={{ background: '#fff', padding: 15, borderRadius: 15, border: '1px solid #e2e8f0', textAlign: 'center' }}>
                  <div style={{ fontSize: 10, fontWeight: 800, color: '#94a3b8', textTransform: 'uppercase' }}>{lbl}</div>
                  <div style={{ fontSize: 18, fontWeight: 900, color: col as string, marginTop: 5 }}>{val}</div>
                </div>
              ))}
          </div>
        )}

        {/* CONTENT & SEARCH */}
        <div style={{ display: 'grid', gridTemplateColumns: preview ? '1fr 400px' : '1fr', gap: 20 }}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <input placeholder="🔍 Buscar por folio o nombre..." value={search} onChange={e => setSearch(e.target.value)} style={{ padding: '12px', borderRadius: 10, border: '1px solid #e2e8f0', marginBottom: 10 }} />
            
            {files.filter(f => f.folio.toLowerCase().includes(search.toLowerCase()) || f.name.toLowerCase().includes(search.toLowerCase())).map(f => {
              const a = analyses[f.id]; const isAn = analyzingIds.has(f.id);
              return (
                <div key={f.id} onClick={() => setPreview(f)} style={{ background: '#fff', padding: 15, borderRadius: 12, border: `2px solid ${preview?.id === f.id ? '#3b82f6' : '#e2e8f0'}`, cursor: 'pointer' }}>
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                    <div style={{ flex: 1 }}>
                      <div style={{ fontWeight: 800, fontSize: 14 }}>{f.folio} <span style={{ fontWeight: 400, color: '#94a3b8', fontSize: 12 }}>{f.name}</span></div>
                      {a && <AnalysisCard r={a} />}
                    </div>
                    <button onClick={e => { e.stopPropagation(); analyzeOne(f); }} disabled={isAn} style={{ background: a ? '#f1f5f9' : '#7c3aed', color: a ? '#64748b' : '#fff', border: 'none', padding: '8px 12px', borderRadius: 8, cursor: 'pointer' }}>{isAn ? '...' : a ? '↻' : '🤖'}</button>
                  </div>
                </div>
              );
            })}
          </div>

          {/* PREVIEW PANEL */}
          {preview && (
            <div style={{ background: '#fff', borderRadius: 15, border: '1px solid #e2e8f0', position: 'sticky', top: 20, height: '85vh', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
              <div style={{ padding: 15, borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <strong style={{ fontSize: 14 }}>{preview.folio}</strong>
                <button onClick={() => setPreview(null)} style={{ border: 'none', background: 'none', fontSize: 20, cursor: 'pointer' }}>✕</button>
              </div>
              <div style={{ flex: 1, background: '#f8fafc' }}>
                <iframe src={proxyUrl(token!, preview.driveId, preview.id)} style={{ width: '100%', height: '100%', border: 'none' }} />
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
