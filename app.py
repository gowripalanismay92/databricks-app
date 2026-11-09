
import os
from html import escape

import msal
import requests

from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse

app = FastAPI(title="Prism Databricks Identity App")


# --------------------------------------------------
# CONFIGURATION
# --------------------------------------------------

TENANT_ID = os.environ.get("ENTRA_TENANT_ID")
CLIENT_ID = os.environ.get("GRAPH_CLIENT_ID")
CLIENT_SECRET = os.environ.get("GRAPH_CLIENT_SECRET")


# --------------------------------------------------
# DATABRICKS IDENTITY
# --------------------------------------------------

def get_databricks_identity(request: Request) -> dict:
    """
    Read identity headers injected by the Databricks Apps
    reverse proxy.

    Trust these headers only when the app is accessed
    through the authenticated Databricks Apps proxy.
    """

    headers = request.headers

    return {
        "username": headers.get(
            "x-forwarded-preferred-username"
        ),
        "email": headers.get(
            "x-forwarded-email"
        ),
        "user_id": headers.get(
            "x-forwarded-user"
        )
    }


def require_databricks_user(request: Request) -> dict:

    user = get_databricks_identity(request)

    if not user["user_id"]:
        raise HTTPException(
            status_code=401,
            detail=(
                "No Databricks user identity was received. "
                "Open this app through its authenticated "
                "Databricks Apps URL."
            )
        )

    return user


# --------------------------------------------------
# MICROSOFT GRAPH AUTHENTICATION
# --------------------------------------------------

def get_graph_access_token():

    if not all([TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
        raise HTTPException(
            status_code=500,
            detail="Microsoft Graph configuration is incomplete"
        )

    authority = (
        f"https://login.microsoftonline.com/{TENANT_ID}"
    )

    graph_app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )

    result = graph_app.acquire_token_for_client(
        scopes=[
            "https://graph.microsoft.com/.default"
        ]
    )

    if "access_token" not in result:
        # Never expose the token or client secret.
        raise HTTPException(
            status_code=502,
            detail="Microsoft Graph authentication failed"
        )

    return result["access_token"]


# --------------------------------------------------
# RETRIEVE EMPLOYEE ID FROM ENTRA ID
# --------------------------------------------------

def get_employee_profile(email: str):

    if not email:
        raise HTTPException(
            status_code=400,
            detail="Advisor email is not available"
        )

    graph_token = get_graph_access_token()

    # Escape apostrophes in the OData string.
    safe_email = email.replace("'", "''")

    graph_url = (
        "https://graph.microsoft.com/v1.0/users"
    )

    params = {
        "$select": (
            "id,displayName,mail,"
            "userPrincipalName,employeeId,"
            "department,jobTitle"
        ),
        "$filter": (
            f"mail eq '{safe_email}' or "
            f"userPrincipalName eq '{safe_email}'"
        )
    }

    response = requests.get(
        graph_url,
        headers={
            "Authorization": f"Bearer {graph_token}"
        },
        params=params,
        timeout=20
    )

    if response.status_code != 200:
        # Do not expose Graph response details or tokens.
        raise HTTPException(
            status_code=502,
            detail="Microsoft Graph user lookup failed"
        )

    users = response.json().get("value", [])

    if len(users) == 0:
        raise HTTPException(
            status_code=404,
            detail="Advisor not found in Entra ID"
        )

    if len(users) > 1:
        raise HTTPException(
            status_code=409,
            detail="Multiple Entra users matched this email"
        )

    user = users[0]

    return {
        "entra_object_id": user.get("id"),
        "display_name": user.get("displayName"),
        "email": (
            user.get("mail")
            or user.get("userPrincipalName")
        ),
        "employee_id": user.get("employeeId"),
        "department": user.get("department"),
        "job_title": user.get("jobTitle")
    }


# --------------------------------------------------
# PRISM HOME PAGE
# --------------------------------------------------

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):

    user = require_databricks_user(request)

    username = escape(
        user["username"] or "Not provided"
    )

    email = escape(
        user["email"] or "Not provided"
    )

    user_id = escape(
        user["user_id"] or "Not provided"
    )

    return f"""
    <!doctype html>
    <html lang="en">

    <head>
        <meta charset="utf-8">
        <meta name="viewport"
              content="width=device-width, initial-scale=1">

        <title>Prism Advisor Portal</title>

        <style>
            body {{
                font-family: Segoe UI, Arial, sans-serif;
                background: #0f172a;
                color: #e2e8f0;
                min-height: 100vh;
                margin: 0;
                display: grid;
                place-items: center;
            }}

            main {{
                background: #1e293b;
                padding: 36px;
                border-radius: 16px;
                width: min(650px, 90vw);
                box-sizing: border-box;
            }}

            .success {{
                color: #34d399;
            }}

            dt {{
                color: #94a3b8;
                margin-top: 14px;
            }}

            dd {{
                margin: 4px 0 0;
                font-weight: 600;
                overflow-wrap: anywhere;
            }}

            .employee {{
                margin-top: 24px;
                padding: 18px;
                background: #334155;
                border-radius: 10px;
            }}
        </style>
    </head>

    <body>

        <main>

            <p>Prism Advisor Portal</p>

            <h1>Welcome to Prism</h1>

            <p class="success">
                Authenticated by Databricks
            </p>

            <h3>Signed-in Databricks user</h3>

            <dl>

                <dt>Preferred username</dt>
                <dd>{username}</dd>

                <dt>Email</dt>
                <dd>{email}</dd>

                <dt>Databricks user ID</dt>
                <dd>{user_id}</dd>

            </dl>

            <div class="employee">

                <h3>Advisor Employee Profile</h3>

                <p id="employee-name">
                    Loading advisor profile...
                </p>

                <p id="employee-email"></p>

                <p id="employee-id"></p>

                <p id="employee-department"></p>

                <p id="employee-job"></p>

                <p id="employee-error"></p>

            </div>

        </main>

        <script>

        async function loadEmployeeProfile() {{

            try {{

                const response = await fetch(
                    "/api/employee-profile"
                );

                const data = await response.json();

                if (!response.ok) {{
                    throw new Error(
                        data.detail || "Profile lookup failed"
                    );
                }}

                document.getElementById(
                    "employee-name"
                ).textContent =
                    "Name: " +
                    (data.display_name || "Not available");

                document.getElementById(
                    "employee-email"
                ).textContent =
                    "Email: " +
                    (data.email || "Not available");

                document.getElementById(
                    "employee-id"
                ).textContent =
                    "Employee ID: " +
                    (data.employee_id || "Not populated in Entra");

                document.getElementById(
                    "employee-department"
                ).textContent =
                    "Department: " +
                    (data.department || "Not available");

                document.getElementById(
                    "employee-job"
                ).textContent =
                    "Job Title: " +
                    (data.job_title || "Not available");

            }} catch (error) {{

                document.getElementById(
                    "employee-name"
                ).textContent =
                    "Unable to load employee profile";

                document.getElementById(
                    "employee-error"
                ).textContent =
                    error.message;

            }}

        }}

        loadEmployeeProfile();

        </script>

    </body>
    </html>
    """


# --------------------------------------------------
# API: DATABRICKS USER
# --------------------------------------------------

@app.get("/api/user")
async def api_user(request: Request):

    user = require_databricks_user(request)

    return {
        "username": user["username"],
        "email": user["email"],
        "user_id": user["user_id"],
        "authenticated": True,
        "identity_provider": "Databricks Apps proxy"
    }


# --------------------------------------------------
# API: AUTH DEBUG
# --------------------------------------------------

@app.get("/api/auth-debug")
async def auth_debug(request: Request):

    user = require_databricks_user(request)

    user_token = request.headers.get(
        "x-forwarded-access-token"
    )

    return {
        "token_received": bool(user_token),
        "token_type": (
            "Bearer" if user_token else None
        ),
        "token_length": (
            len(user_token) if user_token else 0
        ),
        "username": user["username"],
        "email": user["email"],
        "user_id": user["user_id"],
        "authorization_status": (
            "User authorization token received"
            if user_token
            else "User authorization token not received"
        )
    }


# --------------------------------------------------
# API: EMPLOYEE PROFILE
# --------------------------------------------------

@app.get("/api/employee-profile")
def api_employee_profile(request: Request):

    user = require_databricks_user(request)

    email = (
        user["email"]
        or user["username"]
    )

    if not email:
        raise HTTPException(
            status_code=400,
            detail="Advisor email is unavailable"
        )

    profile = get_employee_profile(email)

    return profile

@app.get("/api/config-check")
async def config_check(request: Request):

    user = require_databricks_user(request)

    return {
        "authenticated": True,
        "tenant_id_configured": bool(
            os.getenv("ENTRA_TENANT_ID")
        ),
        "client_id_configured": bool(
            os.getenv("GRAPH_CLIENT_ID")
        ),
        "client_secret_configured": bool(
            os.getenv("GRAPH_CLIENT_SECRET")
        ),
        "client_secret_length": len(
            os.getenv("GRAPH_CLIENT_SECRET", "")
        )
    }
