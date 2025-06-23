from flask import Flask, redirect, request, session, url_for
import requests
import os
import uuid
import configparser

# Load config
config = configparser.ConfigParser()
config.read("config.cfg")

app = Flask(__name__)
app.secret_key = os.urandom(24)

# ==== Configuration ====
CLIENT_ID = config["azure"]["clientId"]
TENANT_ID = config["azure"]["tenantId"]
CLIENT_SECRET = config["azure"]["clientSecret"]
REDIRECT_URI = 'http://localhost/auth-response'
AUTHORITY = 'https://login.microsoftonline.com/common'  # Or use your tenant ID
SCOPE = ['User.Read']
TOKEN_ENDPOINT = f'{AUTHORITY}/oauth2/v2.0/token'
AUTH_ENDPOINT = f'{AUTHORITY}/oauth2/v2.0/authorize'
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/'

# ==== Routes ====

@app.route('/')
def home():
    return '<a href="/login">Login with Microsoft</a>'

@app.route('/login')
def login():
    session['state'] = str(uuid.uuid4())
    auth_url = (
        f"{AUTH_ENDPOINT}?client_id={CLIENT_ID}"
        f"&response_type=code"
        f"&redirect_uri={REDIRECT_URI}"
        f"&response_mode=query"
        f"&scope={' '.join(SCOPE)}"
        f"&state={session['state']}"
    )
    return redirect(auth_url)

@app.route('/callback')
def callback():
    if request.args.get('state') != session.get('state'):
        return 'State mismatch', 400

    code = request.args.get('code')
    token_data = {
        'client_id': CLIENT_ID,
        'scope': ' '.join(SCOPE),
        'code': code,
        'redirect_uri': REDIRECT_URI,
        'grant_type': 'authorization_code',
        'client_secret': CLIENT_SECRET
    }

    token_res = requests.post(TOKEN_ENDPOINT, data=token_data)
    if token_res.status_code != 200:
        return f"Error getting token: {token_res.text}", 400

    tokens = token_res.json()
    access_token = tokens['access_token']

    # Use the access token to call Microsoft Graph API
    graph_res = requests.get(
        GRAPH_API_ENDPOINT,
        headers={'Authorization': f'Bearer {access_token}'}
    )

    if graph_res.status_code != 200:
        return f"Graph API call failed: {graph_res.text}", 400

    profile = graph_res.json()
    return f"""
    <h2>User Profile</h2>
    <ul>
        <li><strong>Name:</strong> {profile.get('displayName')}</li>
        <li><strong>Email:</strong> {profile.get('mail') or profile.get('userPrincipalName')}</li>
        <li><strong>ID:</strong> {profile.get('id')}</li>
    </ul>
    <pre>{profile}</pre>
    """

if __name__ == '__main__':
    app.run(debug=True)
