export default async function handler(req, res) {
  const { driveId, fileId, token } = req.query;

  if (!driveId || !fileId || !token) {
    return res.status(400).json({ error: 'Missing parameters' });
  }

  const url = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`;

  try {
    const response = await fetch(url, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!response.ok) throw new Error(`SharePoint error: ${response.status}`);

    const contentType = response.headers.get('content-type');
    const buffer = await response.arrayBuffer();

    res.setHeader('Content-Type', contentType);
    res.status(200).send(Buffer.from(buffer));
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
}
