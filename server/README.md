Node backend for OAuth SSO (minimal)

Overview
- Minimal Express server that implements an authorization code flow and proxies the Graph API call used by your plugin (`/messages`).

Files
- `index.js` - server implementation
- `.env.example` - environment variable example
- `package.json` - dependencies

How it works (summary)
1. Visit `/auth/start` to begin sign-in (redirects to Azure login).
2. Azure returns a code to `/auth/callback` which the server exchanges for tokens.
3. Tokens are stored in session and `/messages` calls Microsoft Graph with the access token.

Security note
- This sample stores tokens in a server session for demo purposes only. Use a secure persistent store and rotate secrets for production.

If you want, I can add MSAL-based implementation or an example using a persistent database for token storage.
