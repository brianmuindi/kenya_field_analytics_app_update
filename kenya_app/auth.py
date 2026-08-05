"""
Authentication helpers for Kenya Field Analytics.
Credentials live in .streamlit/secrets.toml  (or Streamlit Cloud Secrets).

[users]
admin   = "your_password"
analyst = "another_password"
"""

import streamlit as st

# ── How long a session stays logged in (minutes) ─────────────────────────────
SESSION_TIMEOUT_MINUTES = 120


def _get_users() -> dict:
    """Return {username: password} from secrets, with a safe fallback."""
    try:
        return dict(st.secrets["users"])
    except Exception:
        # Fallback if secrets aren't configured yet — CHANGE THESE
        return {"admin": "kenya2024!"}


def check_login(username: str, password: str) -> bool:
    users = _get_users()
    expected = users.get(username.strip().lower())
    return expected is not None and expected == password


def is_logged_in() -> bool:
    return st.session_state.get("authenticated", False)


def require_login():
    """
    Call at the top of every page.
    If not logged in, shows the login form and stops the page rendering.
    """
    if is_logged_in():
        return  # already authenticated — let the page render normally

    _show_login_page()
    st.stop()


def logout():
    st.session_state["authenticated"] = False
    st.session_state["username"] = ""
    st.rerun()


def is_admin() -> bool:
    """Return True if the currently logged-in user is an admin."""
    user = st.session_state.get("username", "").strip().lower()
    try:
        admin_users = [u.strip().lower() for u in st.secrets.get("admin_users", ["admin"])]
    except Exception:
        admin_users = ["admin"]
    return user in admin_users


# ── Login page UI ─────────────────────────────────────────────────────────────

def _show_login_page():
    st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family:'DM Sans',sans-serif !important; }
#MainMenu, footer, header { visibility:hidden; }
.stApp { background:#0f1117; }
.block-container { max-width:420px !important; padding-top:80px !important; }

/* card */
.login-card {
    background:#1a1d27; border:1px solid #2a2d3a;
    border-radius:20px; padding:40px 36px 36px;
    margin:0 auto;
}
.login-logo  { text-align:center; font-size:48px; margin-bottom:12px; }
.login-title { text-align:center; font-size:22px; font-weight:700; color:#e8eaf0; margin-bottom:4px; }
.login-sub   { text-align:center; font-size:13px; color:#6b7280; margin-bottom:28px; }

/* inputs */
div[data-testid="stTextInput"] label { color:#9ca3af !important; font-size:12px !important; letter-spacing:1px; text-transform:uppercase; }
div[data-testid="stTextInput"] input {
    background:#0f1117 !important; border:1.5px solid #2a2d3a !important;
    border-radius:10px !important; color:#e8eaf0 !important;
    padding:10px 14px !important; font-size:14px !important;
}
div[data-testid="stTextInput"] input:focus { border-color:#4f8ef7 !important; }

/* login button */
div[data-testid="stButton"] > button {
    background:#4f8ef7 !important; color:#fff !important;
    border:none !important; border-radius:10px !important;
    padding:13px !important; font-size:15px !important;
    font-weight:600 !important; width:100% !important;
    margin-top:6px !important;
}
div[data-testid="stButton"] > button:hover { background:#6b9ff8 !important; }

/* error */
div[data-testid="stAlert"] {
    background:rgba(248,113,113,.1) !important;
    border:1px solid rgba(248,113,113,.3) !important;
    border-radius:10px !important; color:#fca5a5 !important;
}

/* hide sidebar on login page */
section[data-testid="stSidebar"] { display:none !important; }
</style>
""", unsafe_allow_html=True)

    st.markdown("""
<div class="login-card">
  <div class="login-logo">📊</div>
  <div class="login-title">Kenya Field Analytics</div>
  <div class="login-sub">Sign in to access the analytics dashboard</div>
</div>
""", unsafe_allow_html=True)

    with st.form("login_form", clear_on_submit=False):
        username = st.text_input("Username", placeholder="Enter your username")
        password = st.text_input("Password", placeholder="Enter your password", type="password")
        submitted = st.form_submit_button("Sign In →", use_container_width=True)

    if submitted:
        if check_login(username, password):
            st.session_state["authenticated"] = True
            st.session_state["username"] = username.strip().lower()
            st.rerun()
        else:
            st.error("Incorrect username or password. Please try again.")
