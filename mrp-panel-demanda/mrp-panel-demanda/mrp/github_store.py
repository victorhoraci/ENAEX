"""
Guardado de archivos directamente en el repositorio de GitHub desde la app.

Permite que, al subir un Excel en la pestaña "Agregar datos", el archivo quede
guardado en el repo (de forma permanente) sin tener que entrar a GitHub a mano.

Cómo se configura (una sola vez):
--------------------------------
Se usan los "secrets" de Streamlit. En Streamlit Cloud: menú de la app →
Settings → Secrets. Localmente: archivo .streamlit/secrets.toml. Formato:

    APP_PASSWORD       = "tu-contraseña"
    GITHUB_TOKEN       = "github_pat_xxx"        # token con permiso de escritura
    GITHUB_REPO        = "victorhoraci/ENAEX"
    GITHUB_BRANCH      = "main"
    GITHUB_DATA_PREFIX = "mrp-panel-demanda/mrp-panel-demanda/data"

Si no hay token configurado, la app guarda los archivos en las carpetas locales
(comportamiento anterior). Si no hay contraseña, se permite sin pedirla (pero se
recomienda configurarla).
"""

from __future__ import annotations

import base64

API = "https://api.github.com"


# --------------------------------------------------------------------------
# Acceso seguro a los secrets
# --------------------------------------------------------------------------
def _secret(clave: str, defecto=None):
    """Lee un secret de Streamlit sin reventar si no está configurado."""
    try:
        import streamlit as st
        if clave in st.secrets:
            return st.secrets[clave]
    except Exception:
        pass
    return defecto


def _config():
    """Devuelve la configuración de GitHub, o None si no está completa."""
    token = _secret("GITHUB_TOKEN")
    repo = _secret("GITHUB_REPO")
    if not token or not repo:
        return None
    return {
        "token": token,
        "repo": repo,
        "branch": _secret("GITHUB_BRANCH", "main"),
        "prefix": str(_secret("GITHUB_DATA_PREFIX", "data")).strip("/"),
    }


def disponible() -> bool:
    """True si la app puede guardar en GitHub (hay token y repo configurados)."""
    return _config() is not None


# --------------------------------------------------------------------------
# Contraseña
# --------------------------------------------------------------------------
def password_configurada() -> bool:
    return _secret("APP_PASSWORD") is not None


def password_ok(clave: str) -> bool:
    """Compara la clave ingresada con la configurada. Si no hay clave configurada, permite."""
    real = _secret("APP_PASSWORD")
    if real is None:
        return True
    return bool(clave) and clave == real


# --------------------------------------------------------------------------
# Llamadas a la API de GitHub
# --------------------------------------------------------------------------
def _headers(token: str) -> dict:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def _listar_dir(cfg: dict, carpeta: str) -> list[dict]:
    """Lista los archivos de una carpeta del repo (name, path, sha)."""
    import requests
    url = f"{API}/repos/{cfg['repo']}/contents/{carpeta}"
    r = requests.get(url, headers=_headers(cfg["token"]),
                     params={"ref": cfg["branch"]}, timeout=30)
    if r.status_code == 200 and isinstance(r.json(), list):
        return r.json()
    return []


def _sha_de(cfg: dict, ruta: str) -> str | None:
    """Devuelve el sha del archivo si ya existe (necesario para actualizarlo)."""
    import requests
    url = f"{API}/repos/{cfg['repo']}/contents/{ruta}"
    r = requests.get(url, headers=_headers(cfg["token"]),
                     params={"ref": cfg["branch"]}, timeout=30)
    if r.status_code == 200:
        return r.json().get("sha")
    return None


def _put(cfg: dict, ruta: str, contenido: bytes, mensaje: str):
    """Crea o actualiza un archivo en el repo."""
    import requests
    url = f"{API}/repos/{cfg['repo']}/contents/{ruta}"
    data = {
        "message": mensaje,
        "content": base64.b64encode(contenido).decode("ascii"),
        "branch": cfg["branch"],
    }
    sha = _sha_de(cfg, ruta)
    if sha:
        data["sha"] = sha
    r = requests.put(url, headers=_headers(cfg["token"]), json=data, timeout=120)
    r.raise_for_status()


def _delete(cfg: dict, ruta: str, sha: str, mensaje: str):
    """Borra un archivo del repo."""
    import requests
    url = f"{API}/repos/{cfg['repo']}/contents/{ruta}"
    data = {"message": mensaje, "sha": sha, "branch": cfg["branch"]}
    r = requests.delete(url, headers=_headers(cfg["token"]), json=data, timeout=120)
    r.raise_for_status()


# --------------------------------------------------------------------------
# Operaciones de alto nivel (usadas por el panel)
# --------------------------------------------------------------------------
def guardar_mb51(nombre: str, contenido: bytes) -> str:
    """
    MB51 REEMPLAZA: borra los Excel que haya en la carpeta MB51 del repo y
    sube el nuevo. Devuelve el nombre guardado.
    """
    cfg = _config()
    if cfg is None:
        raise RuntimeError("GitHub no está configurado (falta GITHUB_TOKEN o GITHUB_REPO).")
    carpeta = f"{cfg['prefix']}/MB51"
    for item in _listar_dir(cfg, carpeta):
        if item["name"].lower().endswith((".xlsx", ".xls")):
            _delete(cfg, item["path"], item["sha"], f"Reemplazar MB51: borrar {item['name']}")
    _put(cfg, f"{carpeta}/{nombre}", contenido, f"Actualizar MB51: {nombre}")
    return nombre


def agregar_mb5b(nombre: str, contenido: bytes) -> str:
    """
    MB5B SE AGREGA: sube el archivo del mes sin borrar los anteriores (si el
    nombre ya existe, lo actualiza). Devuelve el nombre guardado.
    """
    cfg = _config()
    if cfg is None:
        raise RuntimeError("GitHub no está configurado (falta GITHUB_TOKEN o GITHUB_REPO).")
    carpeta = f"{cfg['prefix']}/MB5B"
    _put(cfg, f"{carpeta}/{nombre}", contenido, f"Agregar MB5B: {nombre}")
    return nombre
