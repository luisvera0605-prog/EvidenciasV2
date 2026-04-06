export const config = { runtime: 'edge' }

export default async function handler(req) {
  // Only allow GET
  if (req.method !== 'GET') {
    return new Response('Method not allowed', { status: 405 })
  }

  const url = new URL(req.url)
  const driveId = url.searchParams.get('driveId')
  const fileId = url.searchParams.get('fileId')
  const token = url.searchParams.get('token')

  if (!driveId || !fileId || !token) {
    return new Response('Missing params', { status: 400 })
  }

  try {
    const graphUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`
    const res = await fetch(graphUrl, {
      headers: { Authorization: `Bearer ${token}` },
      redirect: 'follow',
    })

    if (!res.ok) {
      return new Response(`Graph error: ${res.status}`, { status: res.status })
    }

    const blob = await res.arrayBuffer()
    const contentType = res.headers.get('content-type') ?? 'application/octet-stream'

    return new Response(blob, {
      status: 200,
      headers: {
        'Content-Type': contentType,
        'Access-Control-Allow-Origin': '*',
        'Cache-Control': 'private, max-age=300',
      },
    })
  } catch (e) {
    return new Response(`Error: ${e.message}`, { status: 500 })
  }
}
