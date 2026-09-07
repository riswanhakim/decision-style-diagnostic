export default async function handler(req, res) {
  // Hanya terima method POST
  if (req.method !== 'POST') {
    return res.status(405).json({ message: 'Method Not Allowed' });
  }

  const gasWebAppUrl = process.env.GAS_WEBAPP_URL;

  if (!gasWebAppUrl) {
    console.error('GAS_WEBAPP_URL Environment Variable missing');
    return res.status(500).json({ status: 'error', message: 'Server configuration error: GAS_WEBAPP_URL is not defined.' });
  }

  try {
    const payload = req.body;

    // Forward request ke Google Apps Script secara server-to-server
    const response = await fetch(gasWebAppUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const data = await response.json();
    return res.status(200).json({ status: 'success', data });
  } catch (error) {
    console.error('Error forwarding to GAS:', error);
    return res.status(500).json({ status: 'error', message: error.message });
  }
}
