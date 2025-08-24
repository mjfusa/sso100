const express = require('express');
const axios = require('axios');
const session = require('express-session');
require('dotenv').config();
const msal = require('@azure/msal-node');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');

const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const TENANT_ID = process.env.TENANT_ID;
const PORT = process.env.PORT || 3000;
const REDIRECT_URI = process.env.REDIRECT_URI || `https://sx9lvcg1-3000.usw2.devtunnels.ms/auth/callback`;

// Scopes: separate code-flow (backend API) scopes and Graph scopes.
// CODE_SCOPES are used when initiating interactive auth (must be a single resource + openid/offline_access)
const BACKEND_SCOPE = process.env.BACKEND_SCOPE || `api://auth-71e5173f-60c6-41aa-816d-76bc7582752a/ab1814f4-c848-4582-9733-9be230a383ac/access_as_user`;
const CODE_SCOPES = [BACKEND_SCOPE, 'openid', 'profile', 'offline_access'];
const GRAPH_SCOPES = (process.env.GRAPH_SCOPES ? process.env.GRAPH_SCOPES.split(' ') : ['https://graph.microsoft.com/ServiceMessage.Read.All']);

if (!CLIENT_ID || !CLIENT_SECRET || !TENANT_ID) {
  console.error('Missing required env vars: CLIENT_ID, CLIENT_SECRET, TENANT_ID');
  process.exit(1);
}

// MSAL configuration
const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    clientSecret: CLIENT_SECRET,
  }
};
const cca = new msal.ConfidentialClientApplication(msalConfig);

// JWKS client for token validation
const jwksUri = `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`;
const jwks = jwksClient({ jwksUri, cache: true, rateLimit: true });
function getKey(header, callback) {
  jwks.getSigningKey(header.kid, function(err, key) {
    if (err) return callback(err);
    const signingKey = key.getPublicKey ? key.getPublicKey() : key.publicKey;
    callback(null, signingKey);
  });
}

function validateToken(token) {
  return new Promise((resolve, reject) => {
    jwt.verify(token, getKey, { algorithms: ['RS256'], issuer: `https://login.microsoftonline.com/${TENANT_ID}/v2.0` }, (err, decoded) => {
      if (err) return reject(err);
      resolve(decoded);
    });
  });
}

const app = express();
// Trust the first proxy (needed when running behind dev tunnels / ngrok) so req.secure and protocol checks work
app.set('trust proxy', 1);
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Request logging for debugging: log to console and keep recent requests in memory
const recentRequests = [];
app.use((req, res, next) => {
  try {
    const entry = {
      timestamp: new Date().toISOString(),
      method: req.method,
      url: req.originalUrl || req.url,
      headers: req.headers,
      query: req.query,
      body: req.body
    };
    recentRequests.push(entry);
    if (recentRequests.length > 100) recentRequests.shift();
    console.log(`[REQ] ${entry.timestamp} ${entry.method} ${entry.url}`);
    if (Object.keys(req.query || {}).length) console.log('  query:', req.query);
    if (req.method !== 'GET') console.log('  body:', req.body);
  } catch (e) {
    console.warn('Request logging failed', e);
  }
  next();
});

// Simple debug endpoint to inspect recent requests (dev only)
app.get('/__requests', (req, res) => {
  res.json({ recentRequests });
});

// Make the cookie relaxed for local/dev testing to avoid SameSite/Secure issues across redirects.
const isDev = process.env.NODE_ENV !== 'production';
console.log('NODE_ENV:', process.env.NODE_ENV, 'isDev:', isDev, 'REDIRECT_URI:', REDIRECT_URI);
app.use(session({
  secret: process.env.SESSION_SECRET || 'dev_secret',
  resave: false,
  saveUninitialized: false,
  cookie: {
    maxAge: 24 * 60 * 60 * 1000,
    // In development keep SameSite=lax and secure=false so the browser will send the cookie on the OAuth redirect.
    sameSite: isDev ? 'lax' : (REDIRECT_URI && REDIRECT_URI.startsWith('https') ? 'none' : 'lax'),
    secure: !isDev && (REDIRECT_URI && REDIRECT_URI.startsWith('https')),
    httpOnly: true
  }
}));

app.get('/', (req, res) => {
  res.send('SSO sample server. Endpoints: /auth/start, /auth/callback, /messages');
});

app.get('/auth/start', async (req, res) => {
  const state = Math.random().toString(36).substring(2, 15);
  req.session.state = state;

  const authCodeUrlParameters = {
    scopes: CODE_SCOPES,
    redirectUri: REDIRECT_URI,
    state
  };

  try {
    const authUrl = await cca.getAuthCodeUrl(authCodeUrlParameters);
    // save session before redirect to ensure state persists
    req.session.save(err => {
      if (err) console.warn('session save failed', err);
      console.log('Saved session.state =', req.session.state, '-> redirecting to:', authUrl);
      res.redirect(authUrl);
    });
  } catch (err) {
    console.error('getAuthCodeUrl failed', err);
    res.status(500).send('Auth start failed');
  }
});

app.get('/auth/callback', async (req, res) => {
  // DEBUG: print incoming request data to help diagnose missing code / state mismatch
  console.log('--- auth/callback hit ---');
  console.log('method:', req.method);
  console.log('url:', req.originalUrl || req.url);
  console.log('query:', req.query);
  console.log('body:', req.body);
  console.log('headers.cookie:', req.headers.cookie);
  console.log('session id:', req.sessionID);
  console.log('session obj (keys):', req.session ? Object.keys(req.session) : req.session);
  // extra proxy/secure diagnostics
  console.log('req.secure:', req.secure);
  console.log("x-forwarded-proto:", req.headers['x-forwarded-proto']);

  // accept code/state from query or form POST
  const code = (req.body && req.body.code) || (req.query && req.query.code);
  const state = (req.body && req.body.state) || (req.query && req.query.state);

  if (!code) {
    console.warn('Missing code in callback body/query');
    if (req.query.error || req.body.error || req.query.error_description || req.body.error_description) {
      console.warn('AAD error:', req.query.error || req.body.error, req.query.error_description || req.body.error_description);
    }
    return res.status(400).send('Missing code.');
  }
  if (state !== req.session.state) return res.status(400).send('Invalid state.');

  const tokenRequest = {
    code,
    scopes: CODE_SCOPES,
    redirectUri: REDIRECT_URI
  };

  try {
    const tokenResponse = await cca.acquireTokenByCode(tokenRequest);
    // store minimal account to allow acquireTokenSilent later (avoid putting tokens in cookie)
    const acct = tokenResponse.account || {};
    req.session.account = {
      homeAccountId: acct.homeAccountId,
      environment: acct.environment,
      tenantId: acct.tenantId,
      username: acct.username
    };
    // store the backend access token received from the code exchange so we can perform OBO to call Graph
    req.session.backendAccessToken = tokenResponse.accessToken;
    req.session.backendAccessTokenExpiresAt = tokenResponse.expiresOn ? (new Date(tokenResponse.expiresOn)).getTime() : (Date.now() + ((tokenResponse.expiresIn || 3600) * 1000));

    return res.redirect('/messages');
  } catch (err) {
    console.error('acquireTokenByCode failed', err);
    return res.status(500).send('Token exchange failed.');
  }
});

// Token endpoint for host-managed OAuth flows.
// Accepts form-encoded or JSON POST with 'code' and optional 'redirect_uri'.
app.post('/auth/token', async (req, res) => {
  const code = req.body && (req.body.code || req.body['authorization_code']) || req.query && req.query.code;
  const redirectUri = (req.body && (req.body.redirect_uri || req.body.redirectUri)) || req.query && req.query.redirect_uri || REDIRECT_URI;

  if (!code) {
    return res.status(400).json({ error: 'invalid_request', error_description: 'Missing authorization code' });
  }

  const tokenRequest = {
    code,
    scopes: CODE_SCOPES,
    redirectUri
  };

  try {
    const tokenResponse = await cca.acquireTokenByCode(tokenRequest);
    const out = {
      access_token: tokenResponse.accessToken,
      token_type: 'Bearer',
      expires_in: tokenResponse.expiresIn || 3600,
      scope: CODE_SCOPES.join(' ')
    };
    if (tokenResponse.refreshToken) out.refresh_token = tokenResponse.refreshToken;
    if (tokenResponse.idToken) out.id_token = tokenResponse.idToken;

    // Also save backend token in session so the browser-based flow can reuse it if needed
    try {
      if (req.session) {
        req.session.backendAccessToken = tokenResponse.accessToken;
        req.session.backendAccessTokenExpiresAt = tokenResponse.expiresOn ? (new Date(tokenResponse.expiresOn)).getTime() : (Date.now() + ((tokenResponse.expiresIn || 3600) * 1000));
      }
    } catch (e) {
      console.warn('Failed to save token to session (non-fatal)', e);
    }

    return res.json(out);
  } catch (err) {
    console.error('/auth/token exchange failed', err);
    const description = err && err.errorMessage ? err.errorMessage : (err && err.message ? err.message : String(err));
    return res.status(400).json({ error: 'invalid_grant', error_description: description });
  }
});

async function getAccessTokenForSession(req) {
  if (!req.session) return null;

  // return cached Graph token when available and not expired
  if (req.session.graphToken && req.session.graphTokenExpiresAt && Date.now() < req.session.graphTokenExpiresAt) {
    return req.session.graphToken;
  }

  // 1) Try to silently acquire Graph token using the MSAL account (works if the app was consented for Graph)
  if (req.session.account) {
    try {
      const silentRequest = {
        account: req.session.account,
        scopes: GRAPH_SCOPES
      };
      const silentResult = await cca.acquireTokenSilent(silentRequest);
      req.session.graphToken = silentResult.accessToken;
      req.session.graphTokenExpiresAt = Date.now() + ((silentResult.expiresIn || 3600) * 1000);
      return silentResult.accessToken;
    } catch (err) {
      console.warn('acquireTokenSilent for Graph failed, will try OBO if possible', err.message || err);
    }
  }

  // 2) If we have the backend access token from the code exchange, perform OBO to get a Graph token
  if (req.session.backendAccessToken) {
    try {
      const oboResponse = await cca.acquireTokenOnBehalfOf({ oboAssertion: req.session.backendAccessToken, scopes: GRAPH_SCOPES });
      req.session.graphToken = oboResponse.accessToken;
      req.session.graphTokenExpiresAt = Date.now() + ((oboResponse.expiresIn || 3600) * 1000);
      return oboResponse.accessToken;
    } catch (err) {
      console.error('OBO to get Graph token failed', err);
      return null;
    }
  }

  return null;
}

// helper: call Graph with a token
async function callGraphMessagesWithToken(accessToken) {
  const apiRes = await axios.get(
    'https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/messages?$count=true',
    { headers: { Authorization: `Bearer ${accessToken}`, Prefer: 'odata.maxpagesize=10' } }
  );
  return apiRes.data;
}

// Shared handler for both endpoints used in the OpenAPI spec and internal route
async function messagesHandler(req, res) {
  // 1) If host provided a bearer token, prefer it
  const authHeader = req.get('authorization') || req.get('Authorization');
  if (authHeader && authHeader.startsWith('Bearer ')) {
    const incoming = authHeader.slice(7);

    // If token is valid and has not expired
    try {
      const decoded = await validateToken(incoming);

      // If token audience is for Graph, call Graph directly
      const graphAudCandidates = ['00000003-0000-0000-c000-000000000000', 'https://graph.microsoft.com'];
      const aud = decoded && decoded.aud ? decoded.aud : null;
      const isGraphToken = aud && graphAudCandidates.some(x => aud === x || (typeof aud === 'string' && aud.indexOf(x) !== -1));

      if (isGraphToken) {
        try {
          const data = await callGraphMessagesWithToken(incoming);
          return res.json(data);
        } catch (err) {
          console.error('Graph call with host token failed', err.response ? err.response.data : err.message);
          return res.status(err.response ? err.response.status : 502).json({ error: 'graph_error', details: err.response ? err.response.data : err.message });
        }
      }

      // Otherwise assume token is for this backend; perform OBO to get Graph token
      try {
        const oboRequest = {
          oboAssertion: incoming,
          scopes: GRAPH_SCOPES
        };
        const oboResponse = await cca.acquireTokenOnBehalfOf(oboRequest);
        const graphToken = oboResponse.accessToken;
        const data = await callGraphMessagesWithToken(graphToken);
        return res.json(data);
      } catch (err) {
        console.error('OBO or Graph call failed', err.response ? err.response.data : err.message);
        return res.status(502).json({ error: 'obo_error', details: err.response ? err.response.data : err.message });
      }
    } catch (err) {
      console.warn('Invalid or expired token', err.message || err);
      return res.status(401).json({ error: 'invalid_token' });
    }
  }

  // 2) Fallback: session-based flow (existing behavior)
  const accessToken = await getAccessTokenForSession(req);
  if (!accessToken) return res.status(401).json({ error: 'not_authenticated', auth_url: '/auth/start' });

  try {
    const data = await callGraphMessagesWithToken(accessToken);
    return res.json(data);
  } catch (err) {
    console.error('Graph call failed', err.response ? err.response.data : err.message);
    return res.status(502).json({ error: 'graph_error', details: err.response ? err.response.data : err.message });
  }
}

// Register both routes to the same handler
app.get('/messages', messagesHandler);
app.get('/admin/serviceAnnouncement/messages', messagesHandler);

app.get('/signout', (req, res) => {
  req.session = null;
  res.send('Signed out.');
});

app.listen(PORT, () => console.log(`Server listening on ${PORT}`));
