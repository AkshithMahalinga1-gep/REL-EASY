"""
common.py — Shared utilities for all REL-EASY scripts.

Provides:
  - load_credentials(script_file, config)  : overlay config.json onto a CONFIG dict
  - build_session(config)                  : returns (session, BASE_URL)
  - make_http_fns(session)                 : returns (get, post, patch) helpers
  - Excel colour constants and style helpers
"""

import json
import os

import requests
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ── Excel colour palette ───────────────────────────────────────────────────────
BLUE_DARK  = "1F3864"
BLUE_MID   = "2E75B6"
BLUE_LIGHT = "D6E4F0"
GREEN_BG   = "E2EFDA"
AMBER_BG   = "FFF2CC"
RED_BG     = "FFDDD8"
WHITE      = "FFFFFF"
GREY_ROW   = "F5F7FA"


# ── Excel style helpers ────────────────────────────────────────────────────────

def header_font(size=11, bold=True, color=WHITE):
    return Font(name="Arial", size=size, bold=bold, color=color)


def cell_font(size=10, bold=False, color="000000"):
    return Font(name="Arial", size=size, bold=bold, color=color)


def fill(hex_color):
    return PatternFill("solid", start_color=hex_color, fgColor=hex_color)


def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)


def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)


# ── Credentials ────────────────────────────────────────────────────────────────

def load_credentials(script_file, config):
    """Load org/project/pat from config.json at the project root.

    Mutates *config* in place. Falls back silently to whatever values are
    already in the dict if config.json is not found.

    Args:
        script_file: Pass ``__file__`` from the calling script so the path to
                     config.json can be resolved relative to the project root.
        config:      The script's CONFIG dict (mutated in place).
    """
    config_path = os.path.normpath(
        os.path.join(os.path.dirname(os.path.abspath(script_file)), "..", "config.json")
    )
    try:
        with open(config_path, encoding="utf-8") as f:
            shared = json.load(f)
        config["org"]     = shared.get("org",     config["org"])
        config["project"] = shared.get("project", config["project"])
        config["pat"]     = shared.get("pat",      config["pat"])
    except FileNotFoundError:
        pass


# ── HTTP session ───────────────────────────────────────────────────────────────

def build_session(config):
    """Create a configured requests.Session and return (session, BASE_URL).

    The session has connection pooling and automatic retries on transient
    errors (429, 5xx) to prevent Windows ephemeral-port exhaustion.
    """
    session = requests.Session()
    session.auth = ("", config["pat"])
    session.headers.update({"Content-Type": "application/json"})
    retry = Retry(total=3, backoff_factor=0.5, status_forcelist=[429, 500, 502, 503, 504])
    adapter = HTTPAdapter(pool_connections=20, pool_maxsize=20, max_retries=retry)
    session.mount("https://", adapter)
    base_url = f"https://dev.azure.com/{config['org']}/{config['project']}/_apis"
    return session, base_url


# ── HTTP helpers ───────────────────────────────────────────────────────────────

def make_http_fns(session):
    """Return (get, post, patch) helpers bound to *session*.

    Each helper automatically injects ``api-version=7.1`` and applies
    (connect, read) timeouts of (10s, 30s).
    """
    def get(url, params=None):
        p = {**(params or {}), "api-version": "7.1"}
        r = session.get(url, params=p, timeout=(10, 30))
        r.raise_for_status()
        return r.json()

    def post(url, body, params=None):
        p = {**(params or {}), "api-version": "7.1"}
        r = session.post(url, json=body, params=p, timeout=(10, 30))
        r.raise_for_status()
        return r.json()

    def patch(url, body, params=None):
        p = {**(params or {}), "api-version": "7.1"}
        r = session.patch(url, json=body, params=p, timeout=(10, 30))
        r.raise_for_status()
        return r.json()

    return get, post, patch
