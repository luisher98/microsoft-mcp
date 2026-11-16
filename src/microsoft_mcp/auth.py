import os
import sys
import msal
import pathlib as pl
from msal.exceptions import MsalServiceError
from typing import NamedTuple

CACHE_FILE = pl.Path.home() / ".microsoft_mcp_token_cache.json"
SCOPES = ["https://graph.microsoft.com/.default"]
# Segundos por petición a login.microsoftonline.com. MSAL pide allí la
# configuración del tenant al crear la aplicación, los metadatos de instancia al
# buscar cuentas en una caché vacía y el token al renovarlo, y sin timeout
# requests espera para siempre: una llamada sin sesión se colgaría hasta que el
# cliente la cortara (radon-api, a los 120 s) en vez de fallar.
TIMEOUT_MSAL = 15
# Los errores de OAuth con que Microsoft dice que ahora no puede dar el token,
# pero que lo dará: está ocupado o ha fallado él (RFC 6749 §4.1.2.1). Así le
# llega a MSAL un 429 al renovar: como un error en la respuesta, no como una
# excepción. Cualquier otro error al renovar es que la sesión ya no vale: un
# refresh token caducado o revocado (invalid_grant), una política que pide
# volver a entrar (interaction_required)... Son los mismos que separa radon-api
# en conectores/graph.py.
ERRORES_PASAJEROS = frozenset({"temporarily_unavailable", "server_error"})


class Account(NamedTuple):
    username: str
    account_id: str


class SinSesion(RuntimeError):
    """La caché de tokens no tiene sesión para la cuenta, o Microsoft ya no la
    renueva: hay que autorizarla.

    `get_token` no la pide con un device flow. En un servidor stdio, stdout es
    el canal del protocolo y nadie vería el código, así que la llamada se
    quedaría esperándolo hasta que el cliente la cortara. Se autoriza con
    `authenticate.py`, o con las herramientas `authenticate_account` y
    `complete_authentication`, que escriben en la misma caché.
    """


class SinTokenPorAhora(RuntimeError):
    """Microsoft no ha dado el token ahora, pero no es que falte la sesión: no
    contestó, contestó con un 5xx o dijo que ahora no puede (un 429 al
    renovar). Pasa solo, y se repite más tarde sin autorizar nada.

    Antes era `SinSesion`, y su texto mandaba a autorizar de nuevo una sesión
    que estaba bien.
    """

    def __init__(self, cuenta: str, motivo: str) -> None:
        super().__init__(
            f"microsoft-mcp: Microsoft no ha dado ahora el token de {cuenta} "
            f"({motivo}): es pasajero, no falta la sesión; se repite más tarde, "
            "sin autorizar nada"
        )


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

    Sin `account_id`, el de la primera cuenta de la caché. Si no hay sesión, o
    Microsoft ya no la renueva, lanza `SinSesion`; si Microsoft no da el token
    ahora pero lo dará, `SinTokenPorAhora`. Las dos, avisando por stderr.
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

    cuenta = f"la cuenta {account_id or 'por defecto'}"
    try:
        result = app.acquire_token_silent_with_error(SCOPES, account=account)
    except (OSError, MsalServiceError) as e:
        # Microsoft no contestó, o contestó con un 5xx: las excepciones de
        # `requests`, el transporte de MSAL, son OSError, y el 5xx MSAL lo lanza
        # como MsalServiceError
        error = SinTokenPorAhora(cuenta, f"{type(e).__name__}: {e}")
        print(error, file=sys.stderr)
        raise error from e

    if not result or "access_token" not in result:
        error = _sin_token(cuenta, result)
        print(error, file=sys.stderr)
        raise error

    cache = app.token_cache
    if isinstance(cache, msal.SerializableTokenCache) and cache.has_state_changed:
        _write_cache(cache.serialize())

    return result["access_token"]


def _sin_token(cuenta: str, result: dict | None) -> SinSesion | SinTokenPorAhora:
    """Por qué MSAL no ha dado el token, según lo que ha devuelto: nada, si la
    caché no tiene con qué pedirlo, o el error de Microsoft al renovarlo."""
    if not result:
        return SinSesion(
            f"microsoft-mcp: no hay sesión de Microsoft en {CACHE_FILE} para "
            f"{cuenta}: hay que autorizar de nuevo; un servidor stdio no puede "
            "pedir el código"
        )
    # su código y su descripción (AADSTS…), tal cual los da Microsoft
    partes = (result.get("error"), result.get("error_description"))
    dice = ": ".join(str(p) for p in partes if p) or "sin más detalle"
    if result.get("error") in ERRORES_PASAJEROS:
        return SinTokenPorAhora(cuenta, dice)
    return SinSesion(
        f"microsoft-mcp: Microsoft no renueva la sesión de {CACHE_FILE} para "
        f"{cuenta} ({dice}): hay que autorizar de nuevo; un servidor stdio no "
        "puede pedir el código"
    )


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
