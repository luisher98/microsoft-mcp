"""Las herramientas de correo que usa radon-api, llamadas como las llama él: por
MCP y con sus mismos argumentos."""

import json
import pathlib as pl
from typing import Any

import httpx
import pytest

from .conftest import llamar_por_mcp

# Desde aquí, un adjunto no va dentro del POST: el mensaje se crea como borrador
# y el adjunto se sube después, con una sesión de subida (tools.py).
TRES_MIB = 3 * 1024 * 1024

# Los argumentos de `enviar_interno` (radon-api, conectores/correo.py): el aviso
# interno del RF-38, tres enlaces en HTML y sin adjuntos. `account_id` lo pone
# `_llamar`.
AVISO_RF38 = {
    "to": "info@example.com",
    "subject": "Informe desatendido generado — Cliente",
    "body": '<p><a href="https://example.com/informe.docx">DOCX</a></p>',
    "content_type": "HTML",
    "account_id": "cuenta-1",
}


def pdf(tmp_path: pl.Path, tamano: int) -> str:
    """Un PDF de `tamano` bytes en el directorio de la prueba; su ruta."""
    ruta = tmp_path / "plano.pdf"
    ruta.write_bytes(b"%PDF-1.7\n".ljust(tamano, b"\0"))
    return str(ruta)


def borrador(adjunto: str) -> dict[str, Any]:
    """Los argumentos de `crear_borrador` (radon-api, conectores/correo.py), que
    deja el informe en Borradores para el cliente; `account_id` lo pone `_llamar`."""
    return {
        "to": "cliente@example.com",
        "subject": "Informe de radón",
        "body": " ",
        "attachments": [adjunto],
        "account_id": "cuenta-1",
    }


def impresion(adjunto: str) -> dict[str, Any]:
    """Los argumentos de la impresora (radon-api, conectores/impresora.py)."""
    return {
        "account_id": "cuenta-1",
        "to": "impresora@example.com",
        "subject": "plano.pdf",
        "body": "plano.pdf",
        "attachments": [adjunto],
    }


@pytest.mark.parametrize(
    ("adjunto", "respuestas", "peticiones"),
    [
        pytest.param(
            None,
            [httpx.Response(502)],
            [("POST", "/v1.0/me/sendMail")],
            id="aviso",
        ),
        pytest.param(
            9,
            [httpx.Response(502)],
            [("POST", "/v1.0/me/sendMail")],
            id="impresora",
        ),
        pytest.param(
            TRES_MIB + 1,
            [
                httpx.Response(201, json={"id": "m1"}),
                httpx.Response(
                    201, json={"uploadUrl": "https://subida.example.com/sesion-1"}
                ),
                httpx.Response(201, json={}),
                httpx.Response(502),
            ],
            [
                ("POST", "/v1.0/me/messages"),
                ("POST", "/v1.0/me/messages/m1/attachments/createUploadSession"),
                ("PUT", "/sesion-1"),
                ("POST", "/v1.0/me/messages/m1/send"),
            ],
            id="impresora-grande",
        ),
    ],
)
def test_send_email_hace_un_solo_post_aunque_graph_de_502(
    graph_simulado, tmp_path, adjunto, respuestas, peticiones
):
    """Por MCP y con los argumentos de los dos envíos de radon-api: el aviso del
    RF-38, en HTML y sin adjuntos, y la impresora, con el PDF adjunto. Si Graph
    lo mandó antes de contestar el 502, repetirlo sería un segundo aviso o una
    segunda impresión.

    Un PDF de 3 MiB o más no cabe en el sendMail: va en un borrador, con una
    sesión de subida, y sale con /send, que tampoco se repite. El borrador ya
    estaba completo, así que no se retira."""
    argumentos = AVISO_RF38 if adjunto is None else impresion(pdf(tmp_path, adjunto))
    graph_simulado.respuestas = list(respuestas)

    resultado = llamar_por_mcp("send_email", argumentos)

    assert resultado.isError
    assert "502 Bad Gateway" in resultado.content[0].text
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == peticiones
    assert graph_simulado.esperas == []


@pytest.mark.parametrize(
    ("argumentos", "formato"),
    [
        pytest.param(AVISO_RF38, "HTML", id="aviso-rf38"),
        pytest.param(
            {k: v for k, v in AVISO_RF38.items() if k != "content_type"},
            "Text",
            id="sin-content-type",
        ),
    ],
)
def test_send_email_manda_el_cuerpo_en_html_si_se_le_pide(
    graph_simulado, argumentos, formato
):
    """radon-api manda el aviso del RF-38 con content_type="HTML". Sin ese
    parámetro, FastMCP rechazaba la llamada al validar los argumentos, antes de
    llegar a Graph, y el aviso no salía nunca. Sin content_type, el cuerpo sigue
    yendo en texto, como el de la impresora."""
    graph_simulado.respuestas = [httpx.Response(202)]

    resultado = llamar_por_mcp("send_email", argumentos)

    assert not resultado.isError, resultado.content
    (peticion,) = graph_simulado.peticiones
    assert (peticion.method, peticion.url.path) == ("POST", "/v1.0/me/sendMail")
    assert json.loads(peticion.content)["message"]["body"] == {
        "contentType": formato,
        "content": argumentos["body"],
    }


CON_UN_ADJUNTO_GRANDE = pytest.mark.parametrize(
    ("herramienta", "argumentos"),
    [
        pytest.param("create_email_draft", borrador, id="borrador"),
        pytest.param("send_email", impresion, id="impresora"),
    ],
)


@CON_UN_ADJUNTO_GRANDE
def test_si_falla_la_subida_de_un_adjunto_grande_el_borrador_se_retira(
    graph_simulado, tmp_path, herramienta, argumentos
):
    """La sesión de subida es un POST y no se repite tras un 502. El borrador ya
    estaba creado, dirigido a quien iba y sin el adjunto, y quien llama no sabe
    su id: no podría retirarlo y, si repite, crearía otro. Se retira y el 502
    llega a radon-api tal cual."""
    graph_simulado.respuestas = [
        httpx.Response(201, json={"id": "m1"}),
        httpx.Response(502),
        httpx.Response(204),
    ]

    resultado = llamar_por_mcp(herramienta, argumentos(pdf(tmp_path, TRES_MIB + 1)))

    assert resultado.isError
    assert "502 Bad Gateway" in resultado.content[0].text
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/me/messages"),
        ("POST", "/v1.0/me/messages/m1/attachments/createUploadSession"),
        ("DELETE", "/v1.0/me/messages/m1"),
    ]
    assert graph_simulado.esperas == []


@CON_UN_ADJUNTO_GRANDE
def test_si_tampoco_se_puede_retirar_el_borrador_el_error_dice_cual_es(
    graph_simulado, tmp_path, herramienta, argumentos
):
    """El DELETE sí se repite tras un 5xx, porque no duplica nada. Si Graph sigue
    sin contestar, el error dice qué borrador ha quedado, para retirarlo a mano."""
    graph_simulado.respuestas = [
        httpx.Response(201, json={"id": "m1"}),
        httpx.Response(502),
        *[httpx.Response(503) for _ in range(4)],
    ]

    resultado = llamar_por_mcp(herramienta, argumentos(pdf(tmp_path, TRES_MIB + 1)))

    assert resultado.isError
    texto = resultado.content[0].text
    assert "502 Bad Gateway" in texto
    assert "el borrador a medias m1 no se ha podido retirar" in texto
    assert "hay que borrarlo a mano" in texto
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/me/messages"),
        ("POST", "/v1.0/me/messages/m1/attachments/createUploadSession"),
        *[("DELETE", "/v1.0/me/messages/m1")] * 4,
    ]
    assert graph_simulado.esperas == [1, 2, 4]


def test_un_borrador_con_un_adjunto_grande_que_sube_bien_se_queda(
    graph_simulado, tmp_path
):
    """Retirarlo es solo para cuando algo falla: si el adjunto sube, el borrador
    queda en Borradores con él."""
    graph_simulado.respuestas = [
        httpx.Response(201, json={"id": "m1"}),
        httpx.Response(201, json={"uploadUrl": "https://subida.example.com/sesion-1"}),
        httpx.Response(201, json={}),
    ]

    resultado = llamar_por_mcp(
        "create_email_draft", borrador(pdf(tmp_path, TRES_MIB + 1))
    )

    assert not resultado.isError, resultado.content
    assert [(p.method, p.url.host, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "graph.microsoft.com", "/v1.0/me/messages"),
        (
            "POST",
            "graph.microsoft.com",
            "/v1.0/me/messages/m1/attachments/createUploadSession",
        ),
        ("PUT", "subida.example.com", "/sesion-1"),
    ]
