"""`auth.get_token`: el token que MSAL da sin preguntar, o un fallo al momento.

Nunca un device flow: en un servidor stdio, stdout es el canal del protocolo y
nadie vería el código.
"""

import msal
import pytest

from microsoft_mcp import auth

from .conftest import AplicacionSinSesion

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

    def acquire_token_silent(
        self, scopes: list[str], account: dict[str, str] | None
    ) -> dict[str, str] | None:
        if account is None:
            return None
        return {"access_token": TOKEN_VIGENTE}


class AplicacionQueRenuevaElToken(AplicacionConSesion):
    """El token de `CUENTA` había caducado: MSAL lo renueva con el refresh token
    y apunta los tokens nuevos en la caché, como en un refresco de verdad."""

    def acquire_token_silent(
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
