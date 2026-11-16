/**
 * MCO Maker Licensing & Device-Limiting Server
 * Node.js / Express API backed by Supabase PostgreSQL
 */

require('dotenv').config();
const crypto = require('crypto');
const express = require('express');
const cors = require('cors');
const { createClient } = require('@supabase/supabase-js');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const ADMIN_SECRET = process.env.ADMIN_SECRET || 'ChangeMeAdminSecret2026';

const TABLE_LICENSES = process.env.TABLE_LICENSES || 'mco_licenses';
const TABLE_DEVICES = process.env.TABLE_DEVICES || 'mco_license_devices';
const TABLE_UPDATES = process.env.TABLE_UPDATES || 'mco_app_updates';
const KEY_PREFIX = process.env.KEY_PREFIX || 'MCO';
const GITHUB_REPO = process.env.GITHUB_REPO || 'naqasht128/MCOMaker';
const GITHUB_TOKEN = process.env.GITHUB_TOKEN || '';

function cleanReleaseNotes(raw) {
  if (!raw || typeof raw !== 'string') return '';
  const lines = raw.replace(/\r\n/g, '\n').split('\n');
  const cleaned = [];

  for (let line of lines) {
    let trimmed = line.trim();
    if (!trimmed) continue;

    // Skip automated GitHub comparison footer lines (e.g. **Full Changelog**: ...)
    if (/full changelog/i.test(trimmed)) continue;

    // Strip GitHub repository links, comparison URLs, and PR references
    if (/github\.com/i.test(trimmed)) {
      trimmed = trimmed
        .replace(/\s*in\s+https?:\/\/github\.com\S+/gi, '')
        .replace(/\[([^\]]+)\]\(https?:\/\/github\.com[^\)]*\)/gi, '$1')
        .replace(/https?:\/\/github\.com\S*/gi, '')
        .replace(/\s*by\s+@[\w-]+/gi, '')
        .trim();
    }

    if (trimmed) {
      cleaned.push(trimmed);
    }
  }

  return cleaned.join('\n').trim();
}

let supabase = null;
if (SUPABASE_URL && SUPABASE_SERVICE_ROLE_KEY) {
  supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, {
    auth: { persistSession: false, autoRefreshToken: false }
  });
} else {
  console.warn('WARNING: Supabase credentials not set in environment.');
}

const PLAN_CONFIG = {
  '6_months': {
    duration_days: 180,
    max_devices: 1,
    label: '6 Months',
    price: parseInt(process.env.PRICE_6_MONTHS || '3000', 10)
  },
  '1_year': {
    duration_days: 365,
    max_devices: 1,
    label: '1 Year',
    price: parseInt(process.env.PRICE_1_YEAR || '5000', 10)
  }
};

function generateSerialKey() {
  const CHARSET = '23456789ABCDEFGHJKLMNPQRSTUVWXYZ';
  const getChunk = (len = 4) => {
    const bytes = crypto.randomBytes(len);
    let str = '';
    for (let i = 0; i < len; i++) {
      str += CHARSET[bytes[i] % CHARSET.length];
    }
    return str;
  };
  return `${KEY_PREFIX}-${getChunk(4)}-${getChunk(4)}-${getChunk(4)}`;
}

// Health check
app.get(['/', '/health'], (req, res) => {
  res.json({
    status: 'ok',
    app: 'MCO Maker License Server',
    tables: { licenses: TABLE_LICENSES, devices: TABLE_DEVICES, updates: TABLE_UPDATES },
    timestamp: new Date().toISOString()
  });
});

// Public: Dynamic subscription plans & pricing
app.get(['/api/plans', '/plans'], (req, res) => {
  const p6 = PLAN_CONFIG['6_months'].price;
  const p1y = PLAN_CONFIG['1_year'].price;
  const plans = [
    {
      id: '6_months',
      label: '6 Months',
      price: p6,
      currency: 'Rs.',
      duration_days: 180,
      max_devices: 1,
      line: `• 6 Months: Rs. ${p6.toLocaleString()} (180 Days, 1 PC)`
    },
    {
      id: '1_year',
      label: '1 Year',
      price: p1y,
      currency: 'Rs.',
      duration_days: 365,
      max_devices: 1,
      recommended: true,
      line: `• 1 Year: Rs. ${p1y.toLocaleString()} (365 Days, 1 PC) [Recommended]`
    }
  ];

  res.json({
    success: true,
    plans,
    summary: plans.map(p => p.line).join('\n')
  });
});

// Verification & Activation Handler
async function handleLicenseVerification(req, res, { isExplicitActivation = false } = {}) {
  try {
    if (!supabase) {
      return res.status(500).json({ valid: false, error: 'Database client not initialized.' });
    }

    const { serial_key, device_id, os_info } = req.body;
    if (!serial_key || !device_id) {
      return res.status(400).json({ valid: false, error: 'Missing serial_key or device_id.' });
    }

    const cleanKey = String(serial_key).trim().toUpperCase();
    const cleanDeviceId = String(device_id).trim().toUpperCase();

    // 1. Fetch license
    const { data: license, error: fetchErr } = await supabase
      .from(TABLE_LICENSES)
      .select('*')
      .eq('serial_key', cleanKey)
      .single();

    if (fetchErr || !license) {
      return res.status(404).json({ valid: false, error: 'Invalid serial key. Please verify and try again.' });
    }

    if (license.is_revoked) {
      return res.status(403).json({ valid: false, error: 'This serial key has been revoked. Contact support.' });
    }

    // 2. Fetch active registered devices
    const { data: devices, error: devErr } = await supabase
      .from(TABLE_DEVICES)
      .select('*')
      .eq('license_id', license.id);

    if (devErr) {
      return res.status(500).json({ valid: false, error: 'Failed to query registered devices.' });
    }

    const existingDevice = (devices || []).find(d => d.device_id === cleanDeviceId);
    const now = new Date();

    // 3. First-time activation
    if (!license.activated_at) {
      const expiresAt = new Date(now.getTime() + license.duration_days * 24 * 60 * 60 * 1000);
      
      const { error: updErr } = await supabase
        .from(TABLE_LICENSES)
        .update({ activated_at: now.toISOString(), expires_at: expiresAt.toISOString() })
        .eq('id', license.id);

      if (updErr) {
        return res.status(500).json({ valid: false, error: 'Failed to record activation date.' });
      }

      await supabase
        .from(TABLE_DEVICES)
        .insert({
          license_id: license.id,
          device_id: cleanDeviceId,
          os_info: os_info || 'Windows PC',
          first_seen: now.toISOString(),
          last_active: now.toISOString()
        });

      return res.json({
        valid: true,
        activated: true,
        serial_key: cleanKey,
        plan_type: license.plan_type,
        duration_days: license.duration_days,
        activated_at: now.toISOString(),
        expires_at: expiresAt.toISOString(),
        devices_used: 1,
        max_devices: license.max_devices
      });
    }

    // 4. Check expiration
    const expiryDate = new Date(license.expires_at);
    if (now > expiryDate) {
      return res.status(403).json({
        valid: false,
        error: `Your subscription expired on ${expiryDate.toLocaleDateString()}. Please renew your license.`
      });
    }

    // 5. Existing device
    if (existingDevice) {
      await supabase
        .from(TABLE_DEVICES)
        .update({ last_active: now.toISOString(), os_info: os_info || existingDevice.os_info })
        .eq('id', existingDevice.id);

      return res.json({
        valid: true,
        serial_key: cleanKey,
        plan_type: license.plan_type,
        duration_days: license.duration_days,
        activated_at: license.activated_at,
        expires_at: license.expires_at,
        devices_used: devices.length,
        max_devices: license.max_devices
      });
    }

    // 6. New device registration check
    if ((devices || []).length >= license.max_devices) {
      return res.status(403).json({
        valid: false,
        error: `Device limit reached (${license.max_devices} of ${license.max_devices} used). Contact support to transfer license.`
      });
    }

    await supabase
      .from(TABLE_DEVICES)
      .insert({
        license_id: license.id,
        device_id: cleanDeviceId,
        os_info: os_info || 'Windows PC',
        first_seen: now.toISOString(),
        last_active: now.toISOString()
      });

    return res.json({
      valid: true,
      serial_key: cleanKey,
      plan_type: license.plan_type,
      duration_days: license.duration_days,
      activated_at: license.activated_at,
      expires_at: license.expires_at,
      devices_used: devices.length + 1,
      max_devices: license.max_devices
    });

  } catch (err) {
    console.error('License verification unexpected error:', err);
    return res.status(500).json({ valid: false, error: 'Internal licensing error.' });
  }
}

app.post(['/api/license/activate', '/api/activate'], (req, res) => handleLicenseVerification(req, res, { isExplicitActivation: true }));
app.post(['/api/license/verify', '/api/verify'], (req, res) => handleLicenseVerification(req, res, { isExplicitActivation: false }));

// Admin Serial Key Generator
app.post(['/api/admin/generate', '/admin/generate'], async (req, res) => {
  try {
    const adminHeader = req.headers['x-admin-secret'] || req.body.admin_secret;
    if (!ADMIN_SECRET || adminHeader !== ADMIN_SECRET) {
      return res.status(401).json({ error: 'Unauthorized: Invalid admin secret.' });
    }

    const { plan_type = '1_year', max_devices, count = 1 } = req.body;
    const config = PLAN_CONFIG[plan_type] || PLAN_CONFIG['1_year'];
    const deviceCap = max_devices ? Number(max_devices) : config.max_devices;

    const generated = [];
    for (let i = 0; i < Math.min(count, 50); i++) {
      const serial = generateSerialKey();
      const { data, error } = await supabase
        .from(TABLE_LICENSES)
        .insert({
          serial_key: serial,
          plan_type: plan_type,
          duration_days: config.duration_days,
          max_devices: deviceCap
        })
        .select()
        .single();

      if (error) {
        console.error('Error inserting key:', error);
      } else {
        generated.push(data);
      }
    }

    return res.json({
      success: true,
      count: generated.length,
      keys: generated.map(g => ({
        serial_key: g.serial_key,
        plan_type: g.plan_type,
        duration_days: g.duration_days,
        max_devices: g.max_devices
      }))
    });
  } catch (err) {
    console.error('Key generation error:', err);
    return res.status(500).json({ error: 'Failed to generate keys.' });
  }
});

// Update checks
app.get(['/api/updates/latest', '/updates/latest'], async (req, res) => {
  try {
    // 1. Query GitHub dynamically (supports both authenticated and public access)
    try {
      const ghHeaders = {
        'User-Agent': 'MCOMaker-LicenseServer',
        'Accept': 'application/vnd.github.v3+json'
      };
      if (GITHUB_TOKEN) {
        ghHeaders['Authorization'] = `Bearer ${GITHUB_TOKEN}`;
      }

      const ghUrl = `https://api.github.com/repos/${GITHUB_REPO}/releases/latest`;
      let ghRes = await fetch(ghUrl, { headers: ghHeaders });

      // If authenticated fetch failed (e.g. expired/invalid token), retry anonymously
      if (!ghRes.ok && GITHUB_TOKEN) {
        delete ghHeaders['Authorization'];
        ghRes = await fetch(ghUrl, { headers: ghHeaders });
      }

      if (ghRes.ok) {
        const data = await ghRes.json();
        const tag = (data.tag_name || '').replace(/^v/, '');
        const assets = data.assets || [];

        const setupAsset = assets.find(a => a.name.toLowerCase().includes('setup') && a.name.toLowerCase().endsWith('.exe'))
          || assets.find(a => a.name.toLowerCase().endsWith('.exe'));

        if (setupAsset) {
          const host = req.get('host');
          const proto = req.headers['x-forwarded-proto'] || 'https';
          const dlUrl = setupAsset.browser_download_url || `${proto}://${host}/api/updates/download?asset_id=${setupAsset.id}`;
          return res.json({
            update_available: true,
            latest_version: tag,
            version: tag,
            file_name: setupAsset.name,
            file_size: setupAsset.size || 0,
            download_url: dlUrl,
            release_notes: cleanReleaseNotes(data.body) || 'New update available with performance enhancements.'
          });
        }
      }
    } catch (ghErr) {
      console.warn('[GH FETCH WARN] Failed to fetch from GitHub API:', ghErr.message);
    }

    // 2. Fallback: Check Supabase table TABLE_UPDATES if present
    try {
      if (supabase) {
        const { data: updateRecord } = await supabase
          .from(TABLE_UPDATES)
          .select('*')
          .order('created_at', { ascending: false })
          .limit(1)
          .maybeSingle();

        if (updateRecord && updateRecord.version) {
          const v = updateRecord.version.replace(/^v/, '');
          return res.json({
            update_available: true,
            latest_version: v,
            version: v,
            file_name: updateRecord.file_name || 'MCOMaker-Setup.exe',
            file_size: updateRecord.file_size || 0,
            download_url: updateRecord.download_url,
            release_notes: cleanReleaseNotes(updateRecord.release_notes) || 'New update available.'
          });
        }
      }
    } catch (dbErr) {
      // Table may not exist yet, continue to default fallback
    }

    // 3. Fallback to environment variables
    const fallbackVersion = (process.env.LATEST_VERSION || '1.0.6').replace(/^v/, '');
    const fallbackUrl = process.env.LATEST_DOWNLOAD_URL || `https://github.com/${GITHUB_REPO}/releases/latest/download/MCOMaker-Setup.exe`;
    return res.json({
      update_available: true,
      latest_version: fallbackVersion,
      version: fallbackVersion,
      file_name: 'MCOMaker-Setup.exe',
      download_url: fallbackUrl,
      release_notes: 'Latest version of MCO Maker.'
    });

  } catch (err) {
    console.error('[UPDATES ERROR]', err);
    return res.status(500).json({ error: 'Failed to check for updates' });
  }
});

/**
 * GET /api/updates/download
 * Proxies/streams the installer file from private GitHub Release securely
 */
app.get(['/api/updates/download', '/updates/download'], async (req, res) => {
  try {
    const { asset_id } = req.query;

    if (asset_id && GITHUB_TOKEN) {
      try {
        const assetUrl = `https://api.github.com/repos/${GITHUB_REPO}/releases/assets/${asset_id}`;
        const ghRes = await fetch(assetUrl, {
          headers: {
            'User-Agent': 'MCOMaker-LicenseServer',
            'Authorization': `Bearer ${GITHUB_TOKEN}`,
            'Accept': 'application/octet-stream'
          },
          redirect: 'follow'
        });

        if (ghRes.ok) {
          const contentType = ghRes.headers.get('content-type') || 'application/octet-stream';
          const arrayBuffer = await ghRes.arrayBuffer();
          const buffer = Buffer.from(arrayBuffer);

          res.setHeader('Content-Type', contentType);
          res.setHeader('Content-Length', buffer.length);
          res.setHeader('Cache-Control', 'no-transform, public, max-age=300');
          res.setHeader('Accept-Ranges', 'bytes');
          res.setHeader('Content-Disposition', 'attachment; filename="MCOMaker-Setup.exe"');

          return res.send(buffer);
        }
      } catch (proxyErr) {
        console.warn('[DOWNLOAD PROXY WARN]', proxyErr.message);
      }
    }

    // Direct redirect to GitHub Release download URL
    try {
      const ghUrl = `https://api.github.com/repos/${GITHUB_REPO}/releases/latest`;
      const ghRes = await fetch(ghUrl, {
        headers: { 'User-Agent': 'MCOMaker-LicenseServer', 'Accept': 'application/vnd.github.v3+json' }
      });
      if (ghRes.ok) {
        const data = await ghRes.json();
        const assets = data.assets || [];
        const setupAsset = (asset_id ? assets.find(a => String(a.id) === String(asset_id)) : null)
          || assets.find(a => a.name.toLowerCase().includes('setup') && a.name.toLowerCase().endsWith('.exe'))
          || assets.find(a => a.name.toLowerCase().endsWith('.exe'));
        if (setupAsset && setupAsset.browser_download_url) {
          return res.redirect(302, setupAsset.browser_download_url);
        }
      }
    } catch (redirectErr) {
      console.warn('[DOWNLOAD REDIRECT WARN]', redirectErr.message);
    }

    const directUrl = process.env.LATEST_DOWNLOAD_URL || `https://github.com/${GITHUB_REPO}/releases/latest/download/MCOMaker-Setup.exe`;
    return res.redirect(302, directUrl);
  } catch (err) {
    console.error('[DOWNLOAD PROXY ERROR]', err);
    return res.redirect(302, `https://github.com/${GITHUB_REPO}/releases/latest/download/MCOMaker-Setup.exe`);
  }
});

// Broadcast release update
app.post(['/api/admin/broadcast-update', '/admin/broadcast-update'], async (req, res) => {
  try {
    const adminHeader = req.headers['x-admin-secret'] || req.body.admin_secret;
    if (!ADMIN_SECRET || adminHeader !== ADMIN_SECRET) {
      return res.status(401).json({ error: 'Unauthorized.' });
    }

    const { version, download_url, file_name, release_notes } = req.body;
    if (!version || !download_url) {
      return res.status(400).json({ error: 'Missing version or download_url.' });
    }

    const { data, error } = await supabase
      .from(TABLE_UPDATES)
      .upsert({
        version,
        download_url,
        file_name: file_name || 'MCOMaker-Setup.exe',
        release_notes: release_notes || 'General improvements and bug fixes.',
        updated_at: new Date().toISOString()
      })
      .select()
      .single();

    if (error) {
      return res.status(500).json({ error: error.message });
    }

    return res.json({ success: true, update: data });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

if (process.env.NODE_ENV !== 'production' && !process.env.VERCEL) {
  app.listen(PORT, () => console.log(`MCO Maker license server running on port ${PORT}`));
}

module.exports = app;
