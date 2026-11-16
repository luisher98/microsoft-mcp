"""`auth.get_token`: el token que MSAL da sin preguntar, o un fallo al momento.

Nunca un device flow: en un servidor stdio, stdout es el canal del protocolo y
nadie vería el código.
"""

import msal
import pytest
import requests
from msal.exceptions import MsalServiceError

from microsoft_mcp import auth, graph

from .conftest import AplicacionSinSesion, llamar_por_mcp

CUENTA = {"home_account_id": "cuenta-1", "username": "info@example.com"}
TOKEN_VIGENTE = "token-vigente"
TOKEN_RENOVADO = "token-renovado"
REFRESH_TOKEN_RENOVADO = "refresh-token-renovado"


class AplicacionConSesion(AplicacionSinSesion):
    """Una caché con `CUENTA` y un token vigente, que MSAL da sin tocarla.

    Como MSAL, sin una cuenta de la caché no da nada. El device flow sigue
    siendo el cable trampa de `AplicacionSinSesion`.
    """

    def get_accounts(self) -> list[dict[str, str]]:
        return [CUENTA]

    def acquire_token_silent_with_error(
        self, scopes: list[str], account: dict[str, str] | None
    ) -> dict[str, str] | None:
        if account is None:
            return None
        return {"access_token": TOKEN_VIGENTE}


class AplicacionQueRenuevaElToken(AplicacionConSesion):
    """El token de `CUENTA` había caducado: MSAL lo renueva con el refresh token
    y apunta los tokens nuevos en la caché, como en un refresco de verdad."""

    def acquire_token_silent_with_error(
        self, scopes: list[str], account: dict[str, str] | None
    ) -> dict[str, str] | None:
        if account is None:
            return None
        self.token_cache.add(
            {
                "client_id": self.client_id,
                "scope": scopes,
                "token_endpoint": f"{self.authority}/oauth2/v2.0/token",
                "response": {
                    "access_token": TOKEN_RENOVADO,
                    "refresh_token": REFRESH_TOKEN_RENOVADO,
                    "expires_in": 3599,
                },
            }
        )
        return {"access_token": TOKEN_RENOVADO}


def aplicacion_que_no_renueva(
    fallo: dict[str, str] | Exception,
) -> type[AplicacionConSesion]:
    """Una caché con `CUENTA`, cuyo token había caducado, y Microsoft que no lo
    renueva: MSAL devuelve `fallo`, el error de Microsoft, o lo lanza, si es
    una excepción, como cuando Microsoft no contesta."""

    class AplicacionQueNoRenueva(AplicacionConSesion):
        def acquire_token_silent_with_error(
            self, scopes: list[str], account: dict[str, str] | None
        ) -> dict[str, str] | None:
            if account is None:
                return None
            if isinstance(fallo, Exception):
                raise fallo
            return fallo

    return AplicacionQueNoRenueva


# Lo que dice Microsoft al renovar cuando ahora no puede dar el token
AHORA_NO_PUEDE = {
    "error": "temporarily_unavailable",
    "error_description": "The server is temporarily unable to handle the request",
}


def refresh_tokens_guardados() -> list[str] | None:
    """Los refresh tokens de la caché que quedó en disco; None si no se escribió."""
    if not auth.CACHE_FILE.exists():
        return None
    cache = msal.SerializableTokenCache()
    cache.deserialize(auth.CACHE_FILE.read_text())
    return [
        rt["secret"]
        for rt in cache.search(msal.TokenCache.CredentialType.REFRESH_TOKEN)
    ]


@pytest.mark.parametrize(
    ("account_id", "cuenta"),
    [
        pytest.param("cuenta-1", "la cuenta cuenta-1", id="con-cuenta"),
        pytest.param(None, "la cuenta por defecto", id="sin-cuenta"),
    ],
)
def test_sin_sesion_falla_al_momento_sin_escribir_en_stdout(capsys, account_id, cuenta):
    """La caché está vacía: la `AplicacionSinSesion` del conftest. Antes se abría
    aquí un device flow, su código salía por stdout y la llamada lo esperaba
    hasta que radon-api la cortaba a los 120 s. Si se pide, salta el cable
    trampa: queda apuntado y la prueba falla, aunque el código lo envuelva en un
    `except Exception`."""
    with pytest.raises(auth.SinSesion) as error:
        auth.get_token(account_id)

    assert str(error.value) == (
        f"microsoft-mcp: no hay sesión de Microsoft en {auth.CACHE_FILE} para "
        f"{cuenta}: hay que autorizar de nuevo; un servidor stdio no puede pedir "
        "el código"
    )
    salida = capsys.readouterr()
    assert salida.out == ""
    assert salida.err == f"{error.value}\n"
    assert AplicacionSinSesion.device_flows_pedidos == []


def test_una_cuenta_que_no_esta_en_la_cache_es_sin_sesion(monkeypatch):
    """En la caché solo está `cuenta-1`. Su token no vale para otra cuenta, que
    leería o mandaría desde otro buzón: sin sesión, y sin device flow."""
    monkeypatch.setattr(msal, "PublicClientApplication", AplicacionConSesion)

    with pytest.raises(auth.SinSesion, match="para la cuenta cuenta-2:"):
        auth.get_token("cuenta-2")

    assert AplicacionSinSesion.device_flows_pedidos == []


@pytest.mark.parametrize(
    ("fallo", "motivo"),
    [
        pytest.param(
            AHORA_NO_PUEDE,
            "temporarily_unavailable: The server is temporarily unable to handle "
            "the request",
            id="un-429",
        ),
        pytest.param(
            {"error": "server_error", "error_description": "Unexpected error"},
            "server_error: Unexpected error",
            id="server_error",
        ),
        pytest.param(
            # como lo lanza la MSAL que pide pyproject.toml (su
            # throttled_http_client.py); la 1.32.3 lo construía mal, y lo que
            # salía era un KeyError('error')
            MsalServiceError("HTTP Error: 503", error=None, error_description=None),
            "MsalServiceError: HTTP Error: 503",
            id="un-5xx",
        ),
        pytest.param(
            requests.exceptions.ConnectTimeout("login.microsoftonline.com"),
            "ConnectTimeout: login.microsoftonline.com",
            id="sin-respuesta",
        ),
    ],
)
def test_si_microsoft_no_da_el_token_ahora_es_pasajero_y_no_sin_sesion(
    monkeypatch, capsys, fallo, motivo
):
    """Un 429 al renovar llega a MSAL como un error en la respuesta; un 5xx o
    una red que no contesta, como una excepción. Con `acquire_token_silent`, el
    429 era None, lo mismo que una caché vacía, y salía como `SinSesion`, que
    manda a autorizar de nuevo una sesión que está bien. Es pasajero: se dice
    lo que pasó, que se repite más tarde, y la caché no se toca."""
    monkeypatch.setattr(
        msal, "PublicClientApplication", aplicacion_que_no_renueva(fallo)
    )

    with pytest.raises(auth.SinTokenPorAhora) as error:
        auth.get_token("cuenta-1")

    mensaje = str(error.value)
    assert mensaje == (
        "microsoft-mcp: Microsoft no ha dado ahora el token de la cuenta "
        f"cuenta-1 ({motivo}): es pasajero, no falta la sesión; se repite más "
        "tarde, sin autorizar nada"
    )
    salida = capsys.readouterr()
    assert salida.out == ""
    assert salida.err == f"{mensaje}\n"
    assert refresh_tokens_guardados() is None


@pytest.mark.parametrize(
    "fallo",
    [
        pytest.param(
            {
                "error": "invalid_grant",
                "error_description": "AADSTS700082: The refresh token has expired "
                "due to inactivity.",
            },
            id="refresh-token-caducado",
        ),
        pytest.param(
            {
                "error": "interaction_required",
                "error_description": "AADSTS50076: you must use multi-factor "
                "authentication.",
            },
            id="hay-que-volver-a-entrar",
        ),
    ],
)
def test_si_microsoft_no_renueva_la_sesion_es_sin_sesion_y_dice_por_que(
    monkeypatch, fallo
):
    """Un refresh token caducado o revocado, o una política que pide volver a
    entrar, no pasan solos: es `SinSesion`, con lo que dijo Microsoft, que con
    `acquire_token_silent` se perdía."""
    monkeypatch.setattr(
        msal, "PublicClientApplication", aplicacion_que_no_renueva(fallo)
    )

    with pytest.raises(auth.SinSesion) as error:
        auth.get_token("cuenta-1")

    assert str(error.value) == (
        f"microsoft-mcp: Microsoft no renueva la sesión de {auth.CACHE_FILE} para "
        f"la cuenta cuenta-1 ({fallo['error']}: {fallo['error_description']}): hay "
        "que autorizar de nuevo; un servidor stdio no puede pedir el código"
    )
    assert AplicacionSinSesion.device_flows_pedidos == []


def test_por_mcp_el_fallo_pasajero_llega_sin_mandar_a_autorizar(
    monkeypatch, graph_simulado, tmp_path
):
    """Lo que ve radon-api: `mcp_cliente.llamar_herramienta` lanza un McpError
    con este texto, y su worker lo reintenta como un fallo técnico (MCP_ERROR).
    Ya no le dice a nadie que autorice de nuevo, y no ha salido nada hacia
    Graph."""
    monkeypatch.setattr(graph, "get_token", auth.get_token)
    monkeypatch.setattr(
        msal, "PublicClientApplication", aplicacion_que_no_renueva(AHORA_NO_PUEDE)
    )
    informe = tmp_path / "9187497.pdf"
    informe.write_bytes(b"%PDF-1.7\n")

    error = llamar_por_mcp(
        "sharepoint_upload_file",
        {
            "account_id": "cuenta-1",
            "drive_id": "drive-1",
            "local_file_path": str(informe),
            "target_path": "it/9187497.pdf",
        },
    )

    assert error is not None
    assert "Microsoft no ha dado ahora el token de la cuenta cuenta-1" in error
    assert "es pasajero" in error
    assert "hay que autorizar" not in error
    assert graph_simulado.peticiones == []


def test_msal_no_espera_sin_limite_a_microsoft():
    """MSAL habla con login.microsoftonline.com al crear la aplicación y, sin
    sesión, al buscar cuentas en la caché vacía. Sin timeout, si Microsoft acepta
    la conexión y no contesta, la llamada se colgaba hasta que radon-api la
    cortaba a los 120 s en vez de fallar."""
    assert auth.get_app().timeout == auth.TIMEOUT_MSAL


@pytest.mark.parametrize(
    ("aplicacion", "token", "guardados"),
    [
        pytest.param(AplicacionConSesion, TOKEN_VIGENTE, None, id="token-vigente"),
        pytest.param(
            AplicacionQueRenuevaElToken,
            TOKEN_RENOVADO,
            [REFRESH_TOKEN_RENOVADO],
            id="token-renovado",
        ),
    ],
)
def test_con_sesion_devuelve_el_token_y_guarda_la_cache_si_cambia(
    monkeypatch, aplicacion, token, guardados
):
    """Con el token vigente, la caché no cambia y el fichero no se escribe. Si
    MSAL lo renueva, se guarda con el refresh token nuevo: sin él, la sesión se
    perdería cuando caducara el viejo."""
    monkeypatch.setattr(msal, "PublicClientApplication", aplicacion)

    assert auth.get_token("cuenta-1") == token
    assert refresh_tokens_guardados() == guardados
