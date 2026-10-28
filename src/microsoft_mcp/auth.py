import os
import sys
import msal
import pathlib as pl
from typing import NamedTuple
from dotenv import load_dotenv

load_dotenv()

CACHE_FILE = pl.Path.home() / ".microsoft_mcp_token_cache.json"
SCOPES = ["https://graph.microsoft.com/.default"]
# Segundos por petición a login.microsoftonline.com. MSAL pide allí la
# configuración del tenant al crear la aplicación, los metadatos de instancia al
# buscar cuentas en una caché vacía y el token al renovarlo, y sin timeout
# requests espera para siempre: una llamada sin sesión se colgaría hasta que el
# cliente la cortara (radon-api, a los 120 s) en vez de fallar.
TIMEOUT_MSAL = 15


class Account(NamedTuple):
    username: str
    account_id: str


class SinSesion(RuntimeError):
    """La caché de tokens no tiene sesión para la cuenta: hay que autorizarla.

    `get_token` no la pide con un device flow. En un servidor stdio, stdout es
    el canal del protocolo y nadie vería el código, así que la llamada se
    quedaría esperándolo hasta que el cliente la cortara. Se autoriza con
    `authenticate.py`, o con las herramientas `authenticate_account` y
    `complete_authentication`, que escriben en la misma caché.
    """


def _read_cache() -> str | None:
    try:
        return CACHE_FILE.read_text()
    except FileNotFoundError:
        return None


def _write_cache(content: str) -> None:
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    CACHE_FILE.write_text(content)


def get_app() -> msal.PublicClientApplication:
    client_id = os.getenv("MICROSOFT_MCP_CLIENT_ID")
    if not client_id:
        raise ValueError("MICROSOFT_MCP_CLIENT_ID environment variable is required")

    tenant_id = os.getenv("MICROSOFT_MCP_TENANT_ID", "common")
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    cache = msal.SerializableTokenCache()
    cache_content = _read_cache()
    if cache_content:
        cache.deserialize(cache_content)

    app = msal.PublicClientApplication(
        client_id, authority=authority, token_cache=cache, timeout=TIMEOUT_MSAL
    )

    return app


def get_token(account_id: str | None = None) -> str:
    """El token de la cuenta: el de la caché, o uno renovado sin preguntar a nadie.

    Sin `account_id`, el de la primera cuenta de la caché. Si no hay sesión,
    avisa por stderr y lanza `SinSesion`.
    """
    app = get_app()

    accounts = app.get_accounts()
    account = None

    if account_id:
        account = next(
            (a for a in accounts if a["home_account_id"] == account_id), None
        )
    elif accounts:
        account = accounts[0]

    result = app.acquire_token_silent(SCOPES, account=account)

    if not result:
        mensaje = (
            f"microsoft-mcp: no hay sesión de Microsoft en {CACHE_FILE} para la "
            f"cuenta {account_id or 'por defecto'}: hay que autorizar de nuevo; "
            "un servidor stdio no puede pedir el código"
        )
        print(mensaje, file=sys.stderr)
        raise SinSesion(mensaje)

    cache = app.token_cache
    if isinstance(cache, msal.SerializableTokenCache) and cache.has_state_changed:
        _write_cache(cache.serialize())

    return result["access_token"]


def list_accounts() -> list[Account]:
    app = get_app()
    return [
        Account(username=a["username"], account_id=a["home_account_id"])
        for a in app.get_accounts()
    ]


def authenticate_new_account() -> Account | None:
    """Authenticate a new account interactively"""
    app = get_app()

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception(
            f"Failed to get device code: {flow.get('error_description', 'Unknown error')}"
        )

    print("\nTo authenticate:")
    print(
        f"1. Visit: {flow.get('verification_uri', flow.get('verification_url', 'https://microsoft.com/devicelogin'))}"
    )
    print(f"2. Enter code: {flow['user_code']}")
    print("3. Sign in with your Microsoft account")
    print("\nWaiting for authentication...")

    result = app.acquire_token_by_device_flow(flow)

    if "error" in result:
        raise Exception(
            f"Auth failed: {result.get('error_description', result['error'])}"
        )

    cache = app.token_cache
    if isinstance(cache, msal.SerializableTokenCache) and cache.has_state_changed:
        _write_cache(cache.serialize())

    # Get the newly added account
    accounts = app.get_accounts()
    if accounts:
        # Find the account that matches the token we just got
        for account in accounts:
            if (
                account.get("username", "").lower()
                == result.get("id_token_claims", {})
                .get("preferred_username", "")
                .lower()
            ):
                return Account(
                    username=account["username"], account_id=account["home_account_id"]
                )
        # If exact match not found, return the last account
        account = accounts[-1]
        return Account(
            username=account["username"], account_id=account["home_account_id"]
        )

    return None
