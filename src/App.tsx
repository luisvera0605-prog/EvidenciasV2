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

async function graphBlob(url: string, token: string): Promise<Blob> {
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } })
  if (!res.ok) throw new Error(`Graph blob ${res.status}`)
  return res.blob()
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

const MEDIA_RE = /\.(jpg|jpeg|png|gif|webp|pdf|bmp|tiff|heic)$/i

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
        if (item.file && MEDIA_RE.test(item.name)) {
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
    } catch {
      // skip inaccessible folder
    }
  }
  return results
}

// Use Vercel edge proxy to avoid SharePoint CORS restrictions
function proxyUrl(token: string, driveId: string, fileId: string): string {
  const params = new URLSearchParams({ driveId, fileId, token })
  return `/api/file?${params}`
}

async function getFileBase64(token: string, driveId: string, fileId: string, _downloadUrl?: string | null): Promise<string> {
  const blob = await fetch(proxyUrl(token, driveId, fileId))
    .then(r => { if (!r.ok) throw new Error(`Proxy ${r.status}`); return r.blob(); })
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onloadend = () => {
      const result = reader.result as string
      resolve(result.split(',')[1])
    }
    reader.onerror = reject
    reader.readAsDataURL(blob)
  })
}

async function getFileBlob(token: string, driveId: string, fileId: string, _downloadUrl?: string | null): Promise<Blob> {
  return fetch(proxyUrl(token, driveId, fileId))
    .then(r => { if (!r.ok) throw new Error(`Proxy ${r.status}`); return r.blob(); })
}

// ============================================================
// CLAUDE VISION
// ============================================================
function getMimeType(file: EvidenciaFile): string {
  if (file.mimeType) return file.mimeType
  const ext = file.name.split('.').pop()?.toLowerCase() ?? ''
  const map: Record<string, string> = {
    jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png',
    gif: 'image/gif', webp: 'image/webp', pdf: 'application/pdf', bmp: 'image/bmp',
  }
  return map[ext] ?? 'image/jpeg'
}

async function analyzeWithClaude(
  base64: string,
  mimeType: string,
  folio: string
): Promise<AnalysisResult> {
  const isPdf = mimeType.includes('pdf')
  const prompt = `Eres auditor financiero. Analiza esta evidencia de pago del folio "${folio}".
Extrae y responde SOLO con JSON válido sin texto adicional:
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
}
semaforo: verde=todo correcto, amarillo=datos parciales, rojo=ilegible o inconsistente`

  const imageContent = isPdf
    ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: base64 } }
    : { type: 'image', source: { type: 'base64', media_type: mimeType, data: base64 } }

  const res = await fetch('/api/claude', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 500,
      messages: [{ role: 'user', content: [imageContent, { type: 'text', text: prompt }] }],
    }),
  })
  const data = await res.json()
  const text: string = data.content?.[0]?.text ?? '{}'
  try {
    return JSON.parse(text.replace(/```json|```/g, '').trim()) as AnalysisResult
  } catch {
    return {
      legible: false, tipo_documento: 'error_parse', fecha: null, monto: null,
      referencia: null, cliente_documento: null, banco_emisor: null,
      folio_presente: null, observaciones: text.slice(0, 120), semaforo: 'rojo',
    }
  }
}

// ============================================================
// HELPERS
// ============================================================
const fmtMXN = (n: number | null) =>
  n != null ? new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN' }).format(n) : '—'

const fmtKB = (b: number) =>
  b > 1024 * 1024 ? `${(b / 1024 / 1024).toFixed(1)} MB` : `${(b / 1024).toFixed(0)} KB`

const SEM: Record<string, { bg: string; color: string; label: string }> = {
  verde:    { bg: '#d1fae5', color: '#064e3b', label: '✓ OK' },
  amarillo: { bg: '#fef3c7', color: '#78350f', label: '⚠ Revisar' },
  rojo:     { bg: '#fee2e2', color: '#7f1d1d', label: '✗ Alerta' },
}

function exportCSV(files: EvidenciaFile[], analyses: Record<string, AnalysisResult>): void {
  const headers = ['Folio','Archivo','URL_SharePoint','Semaforo','Tipo','Fecha','Monto','Referencia','Cliente','Banco','Folio_Presente','Observaciones','Tamaño','Modificado']
  const rows = files.map(f => {
    const a = analyses[f.id]
    return [
      f.folio, f.name, f.webUrl ?? '', a?.semaforo ?? 'sin_analizar', a?.tipo_documento ?? '',
      a?.fecha ?? '', a?.monto ?? '', a?.referencia ?? '', a?.cliente_documento ?? '',
      a?.banco_emisor ?? '', String(a?.folio_presente ?? ''), a?.observaciones ?? '',
      fmtKB(f.size), new Date(f.modified).toLocaleDateString('es-MX'),
    ].map(v => `"${String(v).replace(/"/g, '""')}"`)
  })
  const csv = [headers, ...rows].map(r => r.join(',')).join('\n')
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `Evidencias_${new Date().toISOString().slice(0, 10)}.csv`
  a.click()
  URL.revokeObjectURL(url)
}

// ============================================================
// COMPONENTS
// ============================================================
function AnalysisCard({ r }: { r: AnalysisResult }) {
  const s = SEM[r.semaforo] ?? SEM.amarillo
  return (
    <div style={{ marginTop: 8, padding: 10, background: '#f8faff', borderRadius: 8, border: '1px solid #e2e8f0', fontSize: 12 }}>
      <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap', marginBottom: 6 }}>
        <span style={{ background: s.bg, color: s.color, fontSize: 11, fontWeight: 700, padding: '2px 8px', borderRadius: 20 }}>{s.label}</span>
        {r.tipo_documento && (
          <span style={{ background: '#dbeafe', color: '#1e3a8a', fontSize: 11, fontWeight: 600, padding: '2px 8px', borderRadius: 20 }}>{r.tipo_documento}</span>
        )}
        {!r.legible && (
          <span style={{ background: '#fee2e2', color: '#7f1d1d', fontSize: 11, fontWeight: 600, padding: '2px 8px', borderRadius: 20 }}>Ilegible</span>
        )}
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '3px 12px', color: '#475569' }}>
        {([['Fecha', r.fecha], ['Monto', r.monto != null ? fmtMXN(r.monto) : null],
          ['Referencia', r.referencia], ['Cliente', r.cliente_documento], ['Banco', r.banco_emisor],
        ] as [string, string | null][]).filter(([, v]) => v && v !== 'null').map(([k, v]) => (
          <div key={k}><span style={{ color: '#94a3b8' }}>{k}: </span><strong>{v}</strong></div>
        ))}
      </div>
      {r.observaciones && <p style={{ margin: '6px 0 0', color: '#64748b', fontStyle: 'italic' }}>{r.observaciones}</p>}
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
  const [filterSem, setFilterSem] = useState('todos')
  const [preview, setPreview] = useState<EvidenciaFile | null>(null)
  const [previewUrl, setPreviewUrl] = useState<string | null>(null)
  const stopRef = useRef(false)

  // Handle OAuth callback
  useEffect(() => {
    const params = new URLSearchParams(window.location.search)
    const code = params.get('code')
    if (code) {
      window.history.replaceState({}, '', window.location.pathname)
      pkceExchange(code).then(setToken).catch(e => setError(String(e)))
    } else {
      const t = getStoredToken()
      if (t) setToken(t)
    }
  }, [])

  // Load user + drive when token available
  useEffect(() => {
    if (!token) return
    graphGet('https://graph.microsoft.com/v1.0/me', token)
      .then(d => setUser(d as User))
      .catch(() => {})
    getSiteId(token)
      .then(siteId => getEvidenciasDriveId(token, siteId))
      .then(setDriveId)
      .catch(e => setError('Error conectando a SharePoint: ' + String(e)))
  }, [token])

  const scan = useCallback(async () => {
    if (!token || !driveId) return
    setScanning(true)
    setError(null)
    setFiles([])
    setAnalyses({})
    try {
      const found = await scanFolios(token, driveId, basePath, setScanProgress)
      setFiles(found)
    } catch (e) {
      setError('Error escaneando: ' + String(e))
    }
    setScanning(false)
    setScanProgress(null)
  }, [token, driveId, basePath])

  const analyzeOne = useCallback(async (file: EvidenciaFile) => {
    if (!token) return
    setAnalyzingIds(prev => new Set(prev).add(file.id))
    try {
      const base64 = await getFileBase64(token, file.driveId, file.id, file.downloadUrl)
      const mime = getMimeType(file)
      const result = await analyzeWithClaude(base64, mime, file.folio)
      setAnalyses(prev => ({ ...prev, [file.id]: result }))
    } catch (e) {
      setAnalyses(prev => ({
        ...prev,
        [file.id]: {
          legible: false, tipo_documento: 'error', fecha: null, monto: null,
          referencia: null, cliente_documento: null, banco_emisor: null,
          folio_presente: null, observaciones: String(e), semaforo: 'rojo',
        },
      }))
    }
    setAnalyzingIds(prev => { const s = new Set(prev); s.delete(file.id); return s })
  }, [token])

  const analyzeAll = useCallback(async () => {
    stopRef.current = false
    const pending = files.filter(f => !analyses[f.id])
    setBatchProgress({ current: 0, total: pending.length })
    for (let i = 0; i < pending.length; i++) {
      if (stopRef.current) break
      await analyzeOne(pending[i])
      setBatchProgress({ current: i + 1, total: pending.length })
      if (i % 10 === 9) await new Promise(r => setTimeout(r, 800))
    }
    setBatchProgress(null)
  }, [files, analyses, analyzeOne])

  const openPreview = useCallback(async (file: EvidenciaFile) => {
    if (!token) return
    setPreview(file)
    setPreviewUrl(null)
    try {
      const blob = await getFileBlob(token, file.driveId, file.id, file.downloadUrl)
      setPreviewUrl(URL.createObjectURL(blob))
    } catch {
      // preview failed silently
    }
  }, [token])

  const stats = {
    total: files.length,
    analizados: Object.keys(analyses).length,
    verde: Object.values(analyses).filter(a => a.semaforo === 'verde').length,
    amarillo: Object.values(analyses).filter(a => a.semaforo === 'amarillo').length,
    rojo: Object.values(analyses).filter(a => a.semaforo === 'rojo').length,
    monto: Object.values(analyses).reduce((s, a) => s + (a.monto ?? 0), 0),
  }

  const filtered = files.filter(f => {
    const a = analyses[f.id]
    const matchSearch = !search ||
      f.folio.toLowerCase().includes(search.toLowerCase()) ||
      f.name.toLowerCase().includes(search.toLowerCase())
    if (filterSem === 'verde') return matchSearch && a?.semaforo === 'verde'
    if (filterSem === 'amarillo') return matchSearch && a?.semaforo === 'amarillo'
    if (filterSem === 'rojo') return matchSearch && a?.semaforo === 'rojo'
    if (filterSem === 'pendiente') return matchSearch && !a
    return matchSearch
  })

  // ---- LOGIN SCREEN ----
  if (!token) {
    return (
      <div style={{ minHeight: '100vh', background: '#f1f5f9', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
        <div style={{ background: '#fff', borderRadius: 16, padding: 48, textAlign: 'center', maxWidth: 400, width: '90%', boxShadow: '0 4px 24px rgba(0,0,0,0.08)' }}>
          <div style={{ fontSize: 52, marginBottom: 12 }}>📋</div>
          <h1 style={{ fontSize: 22, fontWeight: 800, color: '#0f172a', margin: '0 0 8px' }}>EvidenciasIQ</h1>
          <p style={{ color: '#64748b', fontSize: 13, margin: '0 0 6px' }}>Verificación automática de evidencias con IA</p>
          <p style={{ color: '#94a3b8', fontSize: 12, margin: '0 0 28px' }}>SharePoint · Claude Vision · Flor de Tabasco</p>
          {error && <div style={{ background: '#fee2e2', color: '#7f1d1d', borderRadius: 8, padding: 10, marginBottom: 16, fontSize: 12 }}>{error}</div>}
          <button
            onClick={pkceLogin}
            style={{ background: '#0078d4', color: '#fff', border: 'none', borderRadius: 10, padding: '13px 0', fontSize: 14, fontWeight: 700, cursor: 'pointer', width: '100%' }}
          >
            🔐 Iniciar sesión con Microsoft
          </button>
        </div>
      </div>
    )
  }

  // ---- MAIN APP ----
  return (
    <div style={{ background: '#f1f5f9', minHeight: '100vh' }}>
      {/* HEADER */}
      <div style={{ background: 'linear-gradient(135deg,#0f172a,#1e3a5f)', padding: '14px 24px', display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexWrap: 'wrap', gap: 10 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <div style={{ width: 34, height: 34, background: '#0078d4', borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center', fontSize: 16 }}>📋</div>
          <div>
            <p style={{ color: '#fff', fontSize: 15, fontWeight: 700, margin: 0 }}>EvidenciasIQ</p>
            <p style={{ color: '#94a3b8', fontSize: 11, margin: 0 }}>{user?.displayName ?? '...'} · PlaneaciónFinanciera</p>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          {files.length > 0 && !batchProgress && (
            <button onClick={analyzeAll} style={{ background: '#7c3aed', color: '#fff', border: 'none', borderRadius: 8, padding: '8px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
              🤖 Analizar todo ({files.filter(f => !analyses[f.id]).length})
            </button>
          )}
          {batchProgress && (
            <button onClick={() => { stopRef.current = true }} style={{ background: '#ef4444', color: '#fff', border: 'none', borderRadius: 8, padding: '8px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
              ⏹ Detener
            </button>
          )}
          {files.length > 0 && stats.analizados > 0 && (
            <button onClick={() => exportCSV(files, analyses)} style={{ background: '#059669', color: '#fff', border: 'none', borderRadius: 8, padding: '8px 14px', fontSize: 12, fontWeight: 600, cursor: 'pointer' }}>
              ⬇ Exportar CSV
            </button>
          )}
          <button onClick={() => { clearAuth(); setToken(null) }} style={{ background: 'rgba(255,255,255,0.1)', color: '#fff', border: '1px solid rgba(255,255,255,0.2)', borderRadius: 8, padding: '8px 12px', fontSize: 12, cursor: 'pointer' }}>
            Salir
          </button>
        </div>
      </div>

      <div style={{ padding: 20, maxWidth: 1300, margin: '0 auto' }}>
        {error && (
          <div style={{ background: '#fee2e2', border: '1px solid #fecaca', borderRadius: 10, padding: 12, marginBottom: 14, color: '#7f1d1d', fontSize: 13 }}>
            ⚠️ {error}
          </div>
        )}

        {/* SCAN BAR */}
        <div style={{ background: '#fff', borderRadius: 12, padding: 18, border: '1px solid #e2e8f0', marginBottom: 16 }}>
          <div style={{ display: 'flex', gap: 10, alignItems: 'flex-end', flexWrap: 'wrap' }}>
            <div style={{ flex: 1, minWidth: 220 }}>
              <label style={{ fontSize: 11, fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.05em', display: 'block', marginBottom: 4 }}>
                Carpeta dentro de la biblioteca Evidencias
              </label>
              <input
                value={basePath}
                onChange={e => setBasePath(e.target.value)}
                style={{ width: '100%', border: '1px solid #e2e8f0', borderRadius: 8, padding: '8px 12px', fontSize: 13, boxSizing: 'border-box' as const }}
              />
            </div>
            <button
              onClick={scan}
              disabled={scanning || !driveId}
              style={{ background: '#1e40af', color: '#fff', border: 'none', borderRadius: 8, padding: '10px 20px', fontSize: 13, fontWeight: 600, cursor: scanning ? 'not-allowed' : 'pointer', opacity: scanning ? 0.7 : 1 }}
            >
              {scanning ? 'Escaneando...' : driveId ? '📂 Escanear carpetas' : 'Conectando...'}
            </button>
          </div>

          {scanProgress && (
            <div style={{ marginTop: 12 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: '#64748b', marginBottom: 4 }}>
                <span>{scanProgress.current_folio ?? 'Leyendo...'}</span>
                <span>{scanProgress.current}/{scanProgress.total}</span>
              </div>
              <div style={{ background: '#e2e8f0', borderRadius: 4, height: 6 }}>
                <div style={{ background: '#3b82f6', height: '100%', borderRadius: 4, width: `${scanProgress.total ? (scanProgress.current / scanProgress.total * 100) : 0}%`, transition: 'width 0.3s' }} />
              </div>
            </div>
          )}

          {batchProgress && (
            <div style={{ marginTop: 12 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, color: '#64748b', marginBottom: 4 }}>
                <span>🤖 Analizando con IA...</span>
                <span>{batchProgress.current}/{batchProgress.total}</span>
              </div>
              <div style={{ background: '#e2e8f0', borderRadius: 4, height: 6 }}>
                <div style={{ background: '#7c3aed', height: '100%', borderRadius: 4, width: `${batchProgress.total ? (batchProgress.current / batchProgress.total * 100) : 0}%`, transition: 'width 0.3s' }} />
              </div>
            </div>
          )}
        </div>

        {/* KPIs */}
        {files.length > 0 && (
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(6,1fr)', gap: 10, marginBottom: 16 }}>
            {([
              ['Archivos', stats.total, '#1e40af', '#dbeafe'],
              ['Analizados', stats.analizados, '#0f766e', '#ccfbf1'],
              ['✓ OK', stats.verde, '#064e3b', '#d1fae5'],
              ['⚠ Revisar', stats.amarillo, '#78350f', '#fef3c7'],
              ['✗ Alertas', stats.rojo, '#7f1d1d', '#fee2e2'],
              ['Monto total', fmtMXN(stats.monto), '#4c1d95', '#ede9fe'],
            ] as [string, string | number, string, string][]).map(([lbl, val, c, bg]) => (
              <div key={lbl} style={{ background: '#fff', borderRadius: 10, padding: 12, border: '1px solid #e2e8f0', textAlign: 'center' }}>
                <p style={{ fontSize: 10, color: '#94a3b8', margin: 0, fontWeight: 700, textTransform: 'uppercase' }}>{lbl}</p>
                <p style={{ fontSize: typeof val === 'string' && val.length > 8 ? 12 : 20, fontWeight: 800, color: c, margin: '4px 0 0' }}>{val}</p>
              </div>
            ))}
          </div>
        )}

        {/* FILTERS */}
        {files.length > 0 && (
          <div style={{ display: 'flex', gap: 8, marginBottom: 14, flexWrap: 'wrap' }}>
            <input
              value={search}
              onChange={e => setSearch(e.target.value)}
              placeholder="🔍 Buscar folio o archivo..."
              style={{ flex: 1, minWidth: 180, border: '1px solid #e2e8f0', borderRadius: 8, padding: '7px 12px', fontSize: 13, background: '#fff' }}
            />
            {(['todos', 'verde', 'amarillo', 'rojo', 'pendiente'] as const).map(f => (
              <button
                key={f}
                onClick={() => setFilterSem(f)}
                style={{
                  padding: '7px 14px', borderRadius: 20, fontSize: 12, fontWeight: 600, border: '2px solid',
                  borderColor: filterSem === f ? '#3b82f6' : '#e2e8f0',
                  background: filterSem === f ? '#eff6ff' : '#fff',
                  color: filterSem === f ? '#1e40af' : '#64748b', cursor: 'pointer',
                }}
              >
                {f === 'todos' ? `Todos (${files.length})` :
                 f === 'verde' ? `✓ OK (${stats.verde})` :
                 f === 'amarillo' ? `⚠ Revisar (${stats.amarillo})` :
                 f === 'rojo' ? `✗ Alertas (${stats.rojo})` :
                 `Pendiente (${files.length - stats.analizados})`}
              </button>
            ))}
          </div>
        )}

        {/* CONTENT */}
        {files.length === 0 && !scanning && (
          <div style={{ background: '#fff', borderRadius: 12, padding: 60, textAlign: 'center', border: '1px solid #e2e8f0' }}>
            <div style={{ fontSize: 44, marginBottom: 12 }}>📁</div>
            <p style={{ color: '#0f172a', fontWeight: 700, fontSize: 15 }}>Ruta configurada: <code style={{ background: '#f1f5f9', padding: '2px 8px', borderRadius: 4 }}>{basePath}</code></p>
            <p style={{ color: '#64748b', fontSize: 13, marginTop: 8 }}>Haz clic en "Escanear carpetas" para leer todos los folios</p>
          </div>
        )}

        <div style={{ display: 'grid', gridTemplateColumns: preview ? '1fr 400px' : '1fr', gap: 16 }}>
          {/* FILE LIST */}
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {filtered.slice(0, 200).map(file => {
              const a = analyses[file.id]
              const isAnalyzing = analyzingIds.has(file.id)
              const isPdf = file.name.toLowerCase().endsWith('.pdf')
              const sem = a?.semaforo
              const rowBg = sem === 'verde' ? '#f0fdf4' : sem === 'amarillo' ? '#fffbeb' : sem === 'rojo' ? '#fef2f2' : '#fff'

              return (
                <div
                  key={file.id}
                  onClick={() => openPreview(file)}
                  style={{ background: rowBg, borderRadius: 10, padding: 14, border: `1px solid ${preview?.id === file.id ? '#3b82f6' : '#e2e8f0'}`, cursor: 'pointer' }}
                >
                  <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 10 }}>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap', marginBottom: 3 }}>
                        <span style={{ fontSize: 16 }}>{isPdf ? '📄' : '🖼️'}</span>
                        <strong style={{ fontSize: 13, color: '#0f172a' }}>{file.folio}</strong>
                        <span style={{ fontSize: 11, color: '#94a3b8' }}>{file.name}</span>
                        {sem && SEM[sem] && (
                          <span style={{ background: SEM[sem].bg, color: SEM[sem].color, fontSize: 11, fontWeight: 700, padding: '1px 7px', borderRadius: 20 }}>
                            {SEM[sem].label}
                          </span>
                        )}
                      </div>
                      <p style={{ margin: 0, fontSize: 11, color: '#94a3b8' }}>
                        {fmtKB(file.size)} · {new Date(file.modified).toLocaleDateString('es-MX')}
                      </p>
                      {a && <AnalysisCard r={a} />}
                    </div>
                    <button
                      onClick={e => { e.stopPropagation(); analyzeOne(file) }}
                      disabled={isAnalyzing}
                      style={{ flexShrink: 0, background: isAnalyzing ? '#f1f5f9' : a ? '#f1f5f9' : '#7c3aed', color: isAnalyzing ? '#94a3b8' : a ? '#64748b' : '#fff', border: 'none', borderRadius: 8, padding: '6px 11px', fontSize: 11, fontWeight: 600, cursor: isAnalyzing ? 'not-allowed' : 'pointer', whiteSpace: 'nowrap' }}
                    >
                      {isAnalyzing ? '⏳' : a ? '↻' : '🤖'}
                    </button>
                  </div>
                </div>
              )
            })}
            {filtered.length > 200 && (
              <p style={{ textAlign: 'center', color: '#94a3b8', fontSize: 12, padding: 12 }}>
                Mostrando 200 de {filtered.length} — usa el buscador para filtrar
              </p>
            )}
          </div>

          {/* PREVIEW PANEL */}
          {preview && (
            <div style={{ background: '#fff', borderRadius: 12, border: '1px solid #e2e8f0', overflow: 'hidden', position: 'sticky', top: 16, maxHeight: '88vh', display: 'flex', flexDirection: 'column' }}>
              <div style={{ padding: '10px 14px', borderBottom: '1px solid #f1f5f9', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <div>
                  <p style={{ margin: 0, fontSize: 13, fontWeight: 700, color: '#0f172a' }}>{preview.folio}</p>
                  <p style={{ margin: 0, fontSize: 11, color: '#94a3b8' }}>{preview.name}</p>
                </div>
                <button onClick={() => { setPreview(null); setPreviewUrl(null) }} style={{ background: 'none', border: 'none', fontSize: 18, cursor: 'pointer', color: '#64748b' }}>✕</button>
              </div>
              <div style={{ flex: 1, overflow: 'auto', padding: 12, background: '#f8faff', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                {previewUrl
                  ? preview.name.toLowerCase().endsWith('.pdf')
                    ? <embed src={previewUrl} type="application/pdf" style={{ width: '100%', height: 480 }} />
                    : <img src={previewUrl} alt={preview.name} style={{ maxWidth: '100%', maxHeight: 480, borderRadius: 6, objectFit: 'contain' }} />
                  : <p style={{ color: '#94a3b8', fontSize: 13 }}>Cargando...</p>}
              </div>
              <div style={{ padding: 12, borderTop: '1px solid #f1f5f9' }}>
                {analyses[preview.id]
                  ? <AnalysisCard r={analyses[preview.id]} />
                  : (
                    <button
                      onClick={() => analyzeOne(preview)}
                      disabled={analyzingIds.has(preview.id)}
                      style={{ width: '100%', background: '#7c3aed', color: '#fff', border: 'none', borderRadius: 8, padding: 10, fontSize: 13, fontWeight: 600, cursor: 'pointer' }}
                    >
                      {analyzingIds.has(preview.id) ? '⏳ Analizando...' : '🤖 Analizar esta evidencia'}
                    </button>
                  )}
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}
