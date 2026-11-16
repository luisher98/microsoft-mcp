"""Las herramientas de correo que usa radon-api, llamadas como las llama él: por
MCP y con sus mismos argumentos."""

import json
import pathlib as pl
from typing import Any

import httpx
import pytest

from microsoft_mcp import graph

from .conftest import llamar_por_mcp

# Desde aquí, un adjunto no va dentro del POST: el mensaje se crea como borrador
# y el adjunto se sube después, con una sesión de subida (tools.py).
TRES_MIB = 3 * 1024 * 1024
# La sesión de subida de un adjunto, que es de Outlook y no de Graph: contesta
# 200 a cada trozo mientras le falta alguno, y 201, sin cuerpo, cuando ya tiene
# el adjunto entero y lo crea. Así lo documenta Graph y así lo simula radon-api
# (tests/conftest.py, `_tramo`)
SESION = "https://subida.example.com/sesion-1"
TROZO_M1 = ("PUT", "/sesion-1")

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
                httpx.Response(201, json={"uploadUrl": SESION}),
                httpx.Response(201),
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

    error = llamar_por_mcp("send_email", argumentos)

    assert error is not None
    assert "502 Bad Gateway" in error
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

    assert llamar_por_mcp("send_email", argumentos) is None
    (peticion,) = graph_simulado.peticiones
    assert (peticion.method, peticion.url.path) == ("POST", "/v1.0/me/sendMail")
    assert json.loads(peticion.content)["message"]["body"] == {
        "contentType": formato,
        "content": argumentos["body"],
    }


ABRIR_SESION_M1 = ("POST", "/v1.0/me/messages/m1/attachments/createUploadSession")
# Retirar un borrador es moverlo a Elementos eliminados
MOVER_M1 = ("POST", "/v1.0/me/messages/m1/move")
A_ELIMINADOS = {"destinationId": "deleteditems"}

CON_UN_ADJUNTO_GRANDE = pytest.mark.parametrize(
    ("herramienta", "argumentos"),
    [
        pytest.param("create_email_draft", borrador, id="borrador"),
        pytest.param("send_email", impresion, id="impresora"),
    ],
)


def sesion_que_no_se_abre() -> list[httpx.Response]:
    """La subida del adjunto que falla: la sesión sigue dando un 502 tras sus tres
    reintentos, tras 1, 2 y 4 s."""
    return [httpx.Response(502) for _ in range(4)]


def destinos(peticiones: list[httpx.Request]) -> list[dict[str, Any]]:
    """El cuerpo de cada move del borrador: adónde se pidió llevarlo."""
    return [
        json.loads(p.content) for p in peticiones if (p.method, p.url.path) == MOVER_M1
    ]


@CON_UN_ADJUNTO_GRANDE
def test_si_falla_la_subida_de_un_adjunto_grande_el_borrador_se_retira(
    graph_simulado, tmp_path, herramienta, argumentos
):
    """Si la sesión de subida del adjunto no se abre, el borrador ya estaba
    creado, dirigido a quien iba y sin el adjunto, y quien llama no sabe su id:
    no podría retirarlo y, si repite, crearía otro. Se retira a Elementos
    eliminados, de donde se recupera, y el 502 llega a radon-api diciendo que
    el borrador está allí, con su destinatario y su asunto. Antes se borraba
    con un DELETE, sin pasar por ahí."""
    graph_simulado.respuestas = [
        httpx.Response(201, json={"id": "m1"}),
        *sesion_que_no_se_abre(),
        httpx.Response(201, json={"id": "m1-en-eliminados"}),
    ]

    error = llamar_por_mcp(herramienta, argumentos(pdf(tmp_path, TRES_MIB + 1)))

    assert error is not None
    assert "502 Bad Gateway" in error
    assert "El borrador a medias se ha movido a Elementos eliminados" in error
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/me/messages"),
        *[ABRIR_SESION_M1] * 4,
        MOVER_M1,
    ]
    assert destinos(graph_simulado.peticiones) == [A_ELIMINADOS]
    assert graph_simulado.esperas == [1, 2, 4]


@CON_UN_ADJUNTO_GRANDE
def test_si_el_borrador_ya_se_movio_el_404_es_que_esta_retirado(
    graph_simulado, tmp_path, herramienta, argumentos
):
    """Mover se repite tras un 5xx: le cambia el id al mensaje, así que
    repetirlo no lo mueve dos veces ni toca otro. Si Graph lo movió antes de
    contestar el 503, el siguiente intento da 404 con ErrorItemNotFound, y es
    que ya no está en Borradores: el error es el 502 de la subida, y lo dice."""
    graph_simulado.respuestas = [
        httpx.Response(201, json={"id": "m1"}),
        *sesion_que_no_se_abre(),
        httpx.Response(503),
        httpx.Response(404, json={"error": {"code": "ErrorItemNotFound"}}),
    ]

    error = llamar_por_mcp(herramienta, argumentos(pdf(tmp_path, TRES_MIB + 1)))

    assert error is not None
    assert "502 Bad Gateway" in error
    assert "El borrador a medias ya no está en Borradores" in error
    assert "no se ha podido retirar" not in error
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/me/messages"),
        *[ABRIR_SESION_M1] * 4,
        MOVER_M1,
        MOVER_M1,
    ]
    assert destinos(graph_simulado.peticiones) == [A_ELIMINADOS] * 2
    assert graph_simulado.esperas == [1, 2, 4, 1]


@pytest.mark.parametrize(
    "no_encontrado",
    [
        pytest.param(
            httpx.Response(404, json={"error": {"code": "ResourceNotFound"}}),
            id="otro-codigo",
        ),
        pytest.param(httpx.Response(404), id="sin-cuerpo"),
    ],
)
@CON_UN_ADJUNTO_GRANDE
def test_un_404_que_no_dice_que_el_borrador_se_haya_ido_no_lo_da_por_retirado(
    graph_simulado, tmp_path, herramienta, argumentos, no_encontrado
):
    """Solo el 404 con ErrorItemNotFound dice que el mensaje ya no está, como en
    radon-api (`_ya_no_esta`, conectores/graph_correo.py). Otro 404 no dice que
    se haya movido: el borrador puede seguir en Borradores, dirigido al
    cliente, y el error da su id, para retirarlo a mano, en vez de callarlo."""
    graph_simulado.respuestas = [
        httpx.Response(201, json={"id": "m1"}),
        *sesion_que_no_se_abre(),
        no_encontrado,
    ]

    error = llamar_por_mcp(herramienta, argumentos(pdf(tmp_path, TRES_MIB + 1)))

    assert error is not None
    assert "502 Bad Gateway" in error
    assert (
        "el borrador a medias m1 no se ha podido retirar (Client error '404 Not "
        "Found'" in error
    )
    assert "si sigue en Borradores, hay que retirarlo a mano" in error
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/me/messages"),
        *[ABRIR_SESION_M1] * 4,
        MOVER_M1,
    ]


@CON_UN_ADJUNTO_GRANDE
def test_si_tampoco_se_puede_retirar_el_borrador_el_error_dice_cual_es(
    graph_simulado, tmp_path, herramienta, argumentos
):
    """Si al moverlo Graph sigue con un 503 tras tres reintentos, el error dice
    qué borrador ha quedado, para retirarlo a mano. Puede que alguno de esos
    503 sí lo moviera, así que no da por hecho que siga en Borradores."""
    graph_simulado.respuestas = [
        httpx.Response(201, json={"id": "m1"}),
        *sesion_que_no_se_abre(),
        *[httpx.Response(503) for _ in range(4)],
    ]

    error = llamar_por_mcp(herramienta, argumentos(pdf(tmp_path, TRES_MIB + 1)))

    assert error is not None
    assert "502 Bad Gateway" in error
    assert (
        "el borrador a medias m1 no se ha podido retirar (Server error '503 Service "
        "Unavailable'" in error
    )
    assert "si sigue en Borradores, hay que retirarlo a mano" in error
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/me/messages"),
        *[ABRIR_SESION_M1] * 4,
        *[MOVER_M1] * 4,
    ]
    assert destinos(graph_simulado.peticiones) == [A_ELIMINADOS] * 4
    assert graph_simulado.esperas == [1, 2, 4, 1, 2, 4]


def test_si_abrir_la_sesion_de_un_adjunto_da_un_5xx_se_repite(
    graph_simulado, tmp_path
):
    """Como la de un fichero del drive: abrirla no crea el adjunto, que solo
    existe cuando la sesión recibe el último trozo, así que repetirla tras un
    503 deja como mucho una sesión vacía, nunca dos adjuntos. Antes el 503
    dejaba el borrador a medias, y había que retirarlo."""
    graph_simulado.respuestas = [
        httpx.Response(201, json={"id": "m1"}),
        httpx.Response(503),
        httpx.Response(201, json={"uploadUrl": SESION}),
        httpx.Response(201),
    ]

    assert (
        llamar_por_mcp("create_email_draft", borrador(pdf(tmp_path, TRES_MIB + 1)))
        is None
    )
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/me/messages"),
        ABRIR_SESION_M1,
        ABRIR_SESION_M1,
        TROZO_M1,
    ]
    assert graph_simulado.esperas == [1]


def test_un_borrador_con_un_adjunto_grande_que_sube_bien_se_queda(
    graph_simulado, tmp_path
):
    """Retirarlo es solo para cuando algo falla: si el adjunto sube, el borrador
    queda en Borradores con él. Un PDF de 3 MiB a 3,75 MiB va en un solo trozo,
    y Outlook lo da por creado con un 201 sin cuerpo; antes se leía como JSON,
    fallaba, y el borrador, ya completo, se retiraba.

    El trozo va sin el token de Graph: la `uploadUrl` ya lleva su credencial,
    y es de Outlook, no de Graph."""
    graph_simulado.respuestas = [
        httpx.Response(201, json={"id": "m1"}),
        httpx.Response(201, json={"uploadUrl": SESION}),
        httpx.Response(201),
    ]

    assert (
        llamar_por_mcp("create_email_draft", borrador(pdf(tmp_path, TRES_MIB + 1)))
        is None
    )
    assert [(p.method, p.url.host, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "graph.microsoft.com", "/v1.0/me/messages"),
        (
            "POST",
            "graph.microsoft.com",
            "/v1.0/me/messages/m1/attachments/createUploadSession",
        ),
        ("PUT", "subida.example.com", "/sesion-1"),
    ]
    trozo = graph_simulado.peticiones[-1]
    assert trozo.headers["Content-Range"] == f"bytes 0-{TRES_MIB}/{TRES_MIB + 1}"
    assert "Authorization" not in trozo.headers


def test_la_impresora_manda_un_pdf_que_va_en_varios_trozos(graph_simulado, tmp_path):
    """Un PDF de más de 3,75 MiB va en trozos de 12 × 320 KiB, por debajo de los
    4 MB que admite Outlook por petición; antes iban de 4,9 MB. Outlook contesta
    200 al primero, porque le falta el resto, y antes ese 200 daba el adjunto
    por subido y el correo salía sin él. Ahora se manda el segundo, el 201 crea
    el adjunto, y solo entonces sale el correo."""
    trozo = graph.TROZO_DE_ADJUNTO
    total = trozo + 1
    graph_simulado.respuestas = [
        httpx.Response(201, json={"id": "m1"}),
        httpx.Response(201, json={"uploadUrl": SESION}),
        httpx.Response(200, json={"nextExpectedRanges": [f"{trozo}-"]}),
        httpx.Response(201),
        httpx.Response(202),
    ]

    assert llamar_por_mcp("send_email", impresion(pdf(tmp_path, total))) is None
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/me/messages"),
        ABRIR_SESION_M1,
        TROZO_M1,
        TROZO_M1,
        ("POST", "/v1.0/me/messages/m1/send"),
    ]
    trozos = [
        p for p in graph_simulado.peticiones if (p.method, p.url.path) == TROZO_M1
    ]
    assert [(p.headers["Content-Range"], len(p.content)) for p in trozos] == [
        (f"bytes 0-{trozo - 1}/{total}", trozo),
        (f"bytes {trozo}-{trozo}/{total}", 1),
    ]
    assert trozo < 4_000_000
    assert all("Authorization" not in p.headers for p in trozos)


@CON_UN_ADJUNTO_GRANDE
def test_si_la_sesion_no_crea_el_adjunto_el_borrador_se_retira(
    graph_simulado, tmp_path, herramienta, argumentos
):
    """Outlook contesta al último trozo con un 200 y lo que le falta, en vez de
    con el 201 que crea el adjunto: el adjunto no existe. El borrador se retira
    como cuando la sesión no se abre, y el correo de la impresora no sale."""
    graph_simulado.respuestas = [
        httpx.Response(201, json={"id": "m1"}),
        httpx.Response(201, json={"uploadUrl": SESION}),
        httpx.Response(200, json={"nextExpectedRanges": ["0-"]}),
        httpx.Response(201, json={"id": "m1-en-eliminados"}),
    ]

    error = llamar_por_mcp(herramienta, argumentos(pdf(tmp_path, TRES_MIB + 1)))

    assert error is not None
    assert "sin crear el adjunto: no se ha adjuntado" in error
    assert "El borrador a medias se ha movido a Elementos eliminados" in error
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/me/messages"),
        ABRIR_SESION_M1,
        TROZO_M1,
        MOVER_M1,
    ]
    assert destinos(graph_simulado.peticiones) == [A_ELIMINADOS]
