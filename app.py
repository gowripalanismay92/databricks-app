from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse
from html import escape

app = FastAPI(title="Prism Databricks Identity Demo")

def get_identity(request: Request):
    headers = request.headers
    return {
        "username": headers.get("x-forwarded-preferred-username"),
        "email": headers.get("x-forwarded-email"),
        "user_id": headers.get("x-forwarded-user"),
        "access_token": headers.get("x-forwarded-access-token")
    }

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    user = get_identity(request)
    username = escape(user["username"] or "Not provided")
    email = escape(user["email"] or "Not provided")
    user_id = escape(user["user_id"] or "Not provided")
    access_token = escape(user["access_token"] or "Not provided")

    if not any(user.values()):
        message = (
            "<p class='warning'>No Databricks identity headers found. "
            "This is expected when running locally. Open the deployed app "
            "through Databricks Apps.</p>"
        )
    else:
        message = "<p class='success'>Authenticated by Databricks Apps proxy</p>"

    return f'''
    <!doctype html>
    <html lang="en">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>Prism App - Identity</title>
      <style>
        body {{ font-family: Segoe UI, Arial, sans-serif; background:#0f172a;
               color:#e2e8f0; min-height:100vh; display:grid; place-items:center; }}
        main {{ background:#1e293b; padding:36px; border-radius:16px;
                width:min(650px,90vw); }}
        h1 {{ margin-top:0; }}
        .success {{ color:#34d399; }}
        .warning {{ color:#fbbf24; }}
        dt {{ color:#94a3b8; margin-top:14px; }}
        dd {{ margin:4px 0 0; font-weight:600; overflow-wrap:anywhere; }}
      </style>
    </head>
    <body>
      <main>
        <p>Databricks Apps • Entra ID SSO identity demo</p>
        <h1>Prism App</h1>
        {message}
        <dl>
          <dt>Preferred username</dt><dd>{username}</dd>
          <dt>Email</dt><dd>{email}</dd>
          <dt>Databricks forwarded user ID</dt><dd>{user_id}</dd>
          <dt>Access token</dt><dd>{access_token}</dd>
        </dl>
        <p>Identity values are read from Databricks Apps request headers along with token.</p>
      </main>
    </body>
    </html>
    '''

@app.get("/api/user")
async def api_user(request: Request):
    return get_identity(request)
