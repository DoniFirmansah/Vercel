
export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbx2p1g2dBmQEkdmbCWB3jMNVoZ4vsdqLRVcdB-h4rCKaFGv-hlVu7f7c97J60fVgtdv/exec';

  try {
    const { action, ...restParams } = req.method === 'GET' ? req.query : req.body;

    if (!action) {
      return res.status(400).json({
        success: false,
        error: 'Action parameter is required'
      });
    }
    const url = new URL(APPS_SCRIPT_URL);
    url.searchParams.append('api', 'guruizin');
    url.searchParams.append('action', action);

    let response;
    
    if (req.method === 'GET') {
      Object.keys(restParams).forEach(key => {
        if (restParams[key] !== undefined && restParams[key] !== null) {
          url.searchParams.append(key, restParams[key]);
        }
      });

      response = await fetch(url.toString(), {
        method: 'GET',
        headers: {
          'User-Agent': 'Vercel-Proxy/1.0'
        },
        redirect: 'follow' 
      });
    } else if (req.method === 'POST') {
      const formData = new URLSearchParams();
      formData.append('api', 'guruizin');
      formData.append('action', action);
      
      Object.keys(restParams).forEach(key => {
        if (restParams[key] !== undefined && restParams[key] !== null) {
          const value = typeof restParams[key] === 'object' 
            ? JSON.stringify(restParams[key]) 
            : String(restParams[key]);
          formData.append(key, value);
        }
      });

      response = await fetch(url.toString(), {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
          'User-Agent': 'Vercel-Proxy/1.0'
        },
        body: formData.toString(),
        redirect: 'follow' 
      });
    } else {
      return res.status(405).json({
        success: false,
        error: 'Method not allowed'
      });
    }

    if (!response.ok) {
      const text = await response.text();
      console.error('Apps Script Error:', response.status, text.substring(0, 200));
      return res.status(response.status).json({
        success: false,
        error: `Apps Script returned status ${response.status}`,
        details: text.substring(0, 500)
      });
    }

    const contentType = response.headers.get('content-type');
    const responseText = await response.text();
    if (!contentType || !contentType.includes('application/json')) {
      console.error('Non-JSON Response:', contentType, responseText.substring(0, 200));
      if (responseText.includes('Sign in') || responseText.includes('sign in')) {
        return res.status(401).json({
          success: false,
          error: 'Apps Script requires authentication. Please check Web App deployment settings (should be "Anyone").'
        });
      }
      
      if (responseText.includes('<!DOCTYPE') || responseText.includes('<html')) {
        return res.status(500).json({
          success: false,
          error: 'Apps Script returned HTML instead of JSON. Check if URL is correct and Web App is deployed with "Anyone" access.',
          details: responseText.substring(0, 500)
        });
      }
      try {
        const data = JSON.parse(responseText);
        return res.status(200).json(data);
      } catch (parseError) {
        return res.status(500).json({
          success: false,
          error: 'Failed to parse response from Apps Script',
          details: responseText.substring(0, 500)
        });
      }
    }
    try {
      const data = JSON.parse(responseText);
      return res.status(200).json(data);
    } catch (parseError) {
      console.error('JSON Parse Error:', parseError, responseText.substring(0, 200));
      return res.status(500).json({
        success: false,
        error: 'Invalid JSON response from Apps Script',
        details: responseText.substring(0, 500)
      });
    }

  } catch (error) {
    console.error('Proxy Error:', error);
    return res.status(500).json({
      success: false,
      error: error.message || 'Internal server error',
      stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
    });
  }
}

