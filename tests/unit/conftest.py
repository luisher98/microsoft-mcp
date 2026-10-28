"""Las pruebas unitarias no salen de la máquina.

Graph es un `httpx.MockTransport` que contesta lo que cada prueba le prepara,
MSAL es una aplicación sin sesión y la caché de tokens vive en el directorio
temporal de cada prueba. Ninguna necesita una cuenta autorizada, ni ve la
caché o el `.env` reales.
"""

import asyncio
import pathlib as pl
from collections.abc import Iterator
from dataclasses import dataclass, field
from typing import Any, ClassVar, NoReturn

import httpx
import msal
import pytest
from fastmcp import Client
from mcp.types import CallToolResult

from microsoft_mcp import auth, graph, tools

TOKEN_DE_PRUEBA = "token-de-prueba"


class AplicacionSinSesion:
    """Ocupa el sitio de `msal.PublicClientApplication`: una caché sin cuentas.

    No tiene token que dar, y pedirle un device flow es un cable trampa: en una
    prueba unitaria nadie va a teclear el código. Cada paso que se pida queda
    apuntado en `device_flows_pedidos`, que `auth_hermetica` vacía en cada
    prueba, y la prueba falla con `pytest.fail`, que es un BaseException: un
    `except Exception` del código no se lo traga.
    """

    device_flows_pedidos: ClassVar[list[str]] = []

    def __init__(
        self,
        client_id: str,
        authority: str,
        token_cache: msal.SerializableTokenCache,
        timeout: float,
    ) -> None:
        self.client_id = client_id
        self.authority = authority
        self.token_cache = token_cache
        self.timeout = timeout

    def get_accounts(self) -> list[dict[str, str]]:
        return []

    def acquire_token_silent(
        self, scopes: list[str], account: dict[str, str] | None
    ) -> None:
        return None

    def initiate_device_flow(self, scopes: list[str]) -> NoReturn:
        self.device_flows_pedidos.append("initiate_device_flow")
        pytest.fail(
            "una prueba unitaria pidió un device flow: nadie va a teclear el código"
        )

    def acquire_token_by_device_flow(self, flow: dict[str, Any]) -> NoReturn:
        self.device_flows_pedidos.append("acquire_token_by_device_flow")
        pytest.fail(
            "una prueba unitaria esperó el código de un device flow: nadie lo va a "
            "teclear"
        )


@pytest.fixture(autouse=True)
def auth_hermetica(monkeypatch: pytest.MonkeyPatch, tmp_path: pl.Path) -> None:
    """Ninguna prueba llega a la cuenta autorizada.

    `auth` carga el `.env` al importarse; el client id de prueba y el tenant por
    defecto tapan lo que traiga. La caché de tokens es un fichero del directorio
    temporal, y MSAL, una aplicación sin sesión, con su lista de device flows
    pedidos vacía.
    """
    monkeypatch.setenv("MICROSOFT_MCP_CLIENT_ID", "client-id-de-prueba")
    monkeypatch.delenv("MICROSOFT_MCP_TENANT_ID", raising=False)
    monkeypatch.setattr(auth, "CACHE_FILE", tmp_path / "token_cache.json")
    monkeypatch.setattr(msal, "PublicClientApplication", AplicacionSinSesion)
    monkeypatch.setattr(AplicacionSinSesion, "device_flows_pedidos", [])


@dataclass
class GraphSimulado:
    """Graph sin red: contesta, por orden, las respuestas que la prueba le prepara.

    Si lo que toca es un `httpx.TransportError`, lo lanza, como un corte de la
    conexión. `peticiones` apunta cada petición que le llega, y `esperas`, los
    segundos que `graph` habría dormido antes de reintentar.
    """

    respuestas: list[httpx.Response | httpx.TransportError] = field(
        default_factory=list
    )
    peticiones: list[httpx.Request] = field(default_factory=list)
    esperas: list[float] = field(default_factory=list)

    def contestar(self, peticion: httpx.Request) -> httpx.Response:
        self.peticiones.append(peticion)
        if not self.respuestas:
            raise AssertionError(
                f"Graph no tenía respuesta preparada para {peticion.method} {peticion.url}"
            )
        respuesta = self.respuestas.pop(0)
        if isinstance(respuesta, httpx.TransportError):
            raise respuesta
        return respuesta


def token_de_prueba(account_id: str | None = None) -> str:
    return TOKEN_DE_PRUEBA


@pytest.fixture(autouse=True)
def graph_simulado(monkeypatch: pytest.MonkeyPatch) -> Iterator[GraphSimulado]:
    """`graph` habla con un `GraphSimulado`, con un token de prueba y sin dormir.

    Va en todas las pruebas, para que ninguna salga a la red por descuido; la que
    habla con Graph lo pide por su nombre y le prepara las respuestas.
    """
    simulado = GraphSimulado()
    with httpx.Client(
        transport=httpx.MockTransport(simulado.contestar), follow_redirects=True
    ) as cliente:
        monkeypatch.setattr(graph, "_client", cliente)
        monkeypatch.setattr(graph, "get_token", token_de_prueba)
        monkeypatch.setattr(graph.time, "sleep", simulado.esperas.append)
        yield simulado


def llamar_por_mcp(herramienta: str, argumentos: dict[str, Any]) -> CallToolResult:
    """Llama a la herramienta como radon-api: por MCP, con un cliente de FastMCP.

    Así FastMCP valida los argumentos contra la firma de la herramienta, como en
    producción, y un fallo llega como un resultado de error con su texto, que es
    lo que lee radon-api. Llamar a la función que hay detrás se saltaría las dos
    cosas.
    """

    async def llamar() -> CallToolResult:
        async with Client(tools.mcp) as cliente:
            return await cliente.call_tool_mcp(herramienta, argumentos)

    return asyncio.run(llamar())
