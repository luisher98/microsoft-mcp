"""`sharepoint_upload_file`, llamada como la llama radon-api: por MCP y con sus
mismos argumentos (`SharePointMicrosoftMcp.subir`, conectores/sharepoint.py).
Por aquí suben los informes y las facturas."""

import json
import pathlib as pl
from typing import Any

import httpx
import pytest

from microsoft_mcp import graph

from .conftest import llamar_por_mcp

SESION = "https://subida.example.com/sesion-1"
SUBIDO = {
    "id": "item-1",
    "name": "9187497.pdf",
    "webUrl": "https://example.sharepoint.com/it/9187497.pdf",
}
# Lo que hay en el drive con el nombre `it`: la carpeta, o un fichero
CARPETA_IT = {"id": "carpeta-it", "name": "it", "folder": {"childCount": 0}}
FICHERO_IT = {"id": "fichero-it", "name": "it", "file": {}}
# Lo que contesta Graph al crear con `conflictBehavior=fail` un nombre que ya está
YA_EXISTE = {"error": {"code": "nameAlreadyExists"}}
# Lo que pide cada paso de la subida a `it/9187497.pdf`
CARPETA = ("GET", "/v1.0/drives/drive-1/root:/it")
CREAR_CARPETA = ("POST", "/v1.0/drives/drive-1/root/children")
# Lo que lleva cada POST que crea la carpeta. `fail` es lo que deja repetirlo:
# con `rename`, la repetición o la segunda subida crearían «it 1», y con
# `replace` sustituirían la carpeta que ya está, con sus documentos dentro
CARPETA_NUEVA = {
    "name": "it",
    "folder": {},
    "@microsoft.graph.conflictBehavior": "fail",
}
SUBIR_ENTERO = ("PUT", "/v1.0/drives/drive-1/root:/it/9187497.pdf:/content")
ABRIR_SESION = (
    "POST",
    "/v1.0/drives/drive-1/root:/it/9187497.pdf:/createUploadSession",
)
TROZO = ("PUT", "/sesion-1")
# Por encima de esto, la subida va por una sesión, en trozos de este tamaño
TROZO_MAXIMO = graph.UPLOAD_CHUNK_SIZE


def subida(tmp_path: pl.Path, tamano: int) -> dict[str, Any]:
    """Los argumentos de radon-api para subir un PDF de `tamano` bytes a
    `it/9187497.pdf`."""
    ruta = tmp_path / "9187497.pdf"
    ruta.write_bytes(b"%PDF-1.7\n".ljust(tamano, b"\0"))
    return {
        "account_id": "cuenta-1",
        "drive_id": "drive-1",
        "local_file_path": str(ruta),
        "target_path": "it/9187497.pdf",
    }


def rangos(peticiones: list[httpx.Request]) -> list[str]:
    """El Content-Range de cada trozo que llegó a la sesión de subida."""
    return [p.headers["Content-Range"] for p in peticiones if p.url.path == TROZO[1]]


def carpetas_creadas(peticiones: list[httpx.Request]) -> list[dict[str, Any]]:
    """El cuerpo de cada POST que creó la carpeta."""
    return [
        json.loads(p.content)
        for p in peticiones
        if (p.method, p.url.path) == CREAR_CARPETA
    ]


def test_si_abrir_la_sesion_de_subida_da_un_5xx_se_repite(graph_simulado, tmp_path):
    """Abrir la sesión no crea el fichero, solo un sitio donde esperan los
    trozos: repetirla tras un 503 deja como mucho una sesión vacía, que
    Microsoft descarta al caducar. Antes el 503 llegaba a radon-api y el
    informe no subía.

    Un trozo que da un 5xx se repite con el mismo tramo; aquí Graph no lo
    tenía, y lo recibe al repetirlo. Los trozos van sin el token de Graph: la
    `uploadUrl` ya lleva su credencial, y con el token puede contestar 401."""
    graph_simulado.respuestas = [
        httpx.Response(200, json=CARPETA_IT),
        httpx.Response(503),
        httpx.Response(200, json={"uploadUrl": SESION}),
        httpx.Response(503),
        httpx.Response(202, json={"nextExpectedRanges": [f"{TROZO_MAXIMO}-"]}),
        httpx.Response(201, json=SUBIDO),
    ]
    argumentos = subida(tmp_path, TROZO_MAXIMO + 1)

    assert llamar_por_mcp("sharepoint_upload_file", argumentos) is None
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        CARPETA,
        ABRIR_SESION,
        ABRIR_SESION,
        TROZO,
        TROZO,
        TROZO,
    ]
    assert rangos(graph_simulado.peticiones) == [
        f"bytes 0-{TROZO_MAXIMO - 1}/{TROZO_MAXIMO + 1}",
        f"bytes 0-{TROZO_MAXIMO - 1}/{TROZO_MAXIMO + 1}",
        f"bytes {TROZO_MAXIMO}-{TROZO_MAXIMO}/{TROZO_MAXIMO + 1}",
    ]
    assert all(
        "Authorization" not in p.headers
        for p in graph_simulado.peticiones
        if p.url.path == TROZO[1]
    )
    assert graph_simulado.esperas == [1, 1]


def test_si_graph_ya_tenia_el_trozo_repetido_la_subida_acaba_en_un_416(
    graph_simulado, tmp_path
):
    """Un trozo que da un 5xx se repite con el mismo tramo, y Graph pudo
    recibirlo antes de fallar. Entonces contesta 416: no lo escribe dos veces,
    pero la subida acaba en ese error, sin el fichero y sin mandar el resto."""
    graph_simulado.respuestas = [
        httpx.Response(200, json=CARPETA_IT),
        httpx.Response(200, json={"uploadUrl": SESION}),
        httpx.Response(503),
        httpx.Response(416),
    ]

    error = llamar_por_mcp("sharepoint_upload_file", subida(tmp_path, TROZO_MAXIMO + 1))

    assert error is not None
    assert "416 Requested Range Not Satisfiable" in error
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        CARPETA,
        ABRIR_SESION,
        TROZO,
        TROZO,
    ]
    assert rangos(graph_simulado.peticiones) == [
        f"bytes 0-{TROZO_MAXIMO - 1}/{TROZO_MAXIMO + 1}",
    ] * 2
    assert graph_simulado.esperas == [1]


def test_si_la_sesion_de_subida_no_se_abre_no_se_sube_nada(graph_simulado, tmp_path):
    """Tres reintentos, tras 1, 2 y 4 s, como cualquier petición que se repite,
    y el último 503 llega a radon-api sin haber mandado ni un trozo."""
    graph_simulado.respuestas = [
        httpx.Response(200, json=CARPETA_IT),
        *[httpx.Response(503) for _ in range(4)],
    ]

    error = llamar_por_mcp("sharepoint_upload_file", subida(tmp_path, TROZO_MAXIMO + 1))

    assert error is not None
    assert "503 Service Unavailable" in error
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        CARPETA,
        *[ABRIR_SESION] * 4,
    ]
    assert graph_simulado.esperas == [1, 2, 4]


def test_una_sesion_que_no_da_el_fichero_por_subido_es_un_error(
    graph_simulado, tmp_path
):
    """Solo un 200 o un 201 con el elemento dicen que el fichero está en el
    drive. Si el último trozo vuelve con un 202, a la sesión le falta algo y el
    fichero no existe: el error lo dice, y no que la subida acabó."""
    graph_simulado.respuestas = [
        httpx.Response(200, json=CARPETA_IT),
        httpx.Response(200, json={"uploadUrl": SESION}),
        httpx.Response(202, json={"nextExpectedRanges": [f"{TROZO_MAXIMO}-"]}),
        httpx.Response(202, json={"nextExpectedRanges": [f"0-{TROZO_MAXIMO - 1}"]}),
    ]

    error = llamar_por_mcp("sharepoint_upload_file", subida(tmp_path, TROZO_MAXIMO + 1))

    assert error is not None
    assert "no se ha subido" in error
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        CARPETA,
        ABRIR_SESION,
        TROZO,
        TROZO,
    ]


@pytest.mark.parametrize(
    ("al_crearla", "esperas"),
    [
        pytest.param(
            [
                (CREAR_CARPETA, httpx.Response(409, json=YA_EXISTE)),
                (CARPETA, httpx.Response(200, json=CARPETA_IT)),
            ],
            [],
            id="otra-subida-la-acaba-de-crear",
        ),
        pytest.param(
            [
                (CREAR_CARPETA, httpx.Response(502)),
                (CREAR_CARPETA, httpx.Response(409, json=YA_EXISTE)),
                (CARPETA, httpx.Response(200, json=CARPETA_IT)),
            ],
            [1],
            id="graph-la-creo-y-dio-502",
        ),
        pytest.param(
            [
                (CREAR_CARPETA, httpx.Response(502)),
                (CREAR_CARPETA, httpx.Response(201, json=CARPETA_IT)),
            ],
            [1],
            id="graph-no-la-creo-y-dio-502",
        ),
    ],
)
def test_la_carpeta_que_ya_existe_al_crearla_se_da_por_buena(
    graph_simulado, tmp_path, al_crearla, esperas
):
    """Dos subidas a la vez a una carpeta nueva: las dos ven que no está y las
    dos van a crearla. Con `conflictBehavior=fail`, la segunda recibe un 409, y
    antes la subida fallaba por eso. Por lo mismo, crearla se repite tras un
    5xx: con `fail` nunca quedan dos carpetas, y si ya la creó el primer
    intento, el segundo da 409. Un 409 solo dice que con ese nombre ya hay
    algo, así que se vuelve a leer, y es la carpeta. Todo eso se apoya en que
    cada POST lleva `fail`, y se mira."""
    graph_simulado.respuestas = [
        httpx.Response(404),
        *[respuesta for _, respuesta in al_crearla],
        httpx.Response(201, json=SUBIDO),
    ]

    assert llamar_por_mcp("sharepoint_upload_file", subida(tmp_path, 9)) is None
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        CARPETA,
        *[peticion for peticion, _ in al_crearla],
        SUBIR_ENTERO,
    ]
    assert carpetas_creadas(graph_simulado.peticiones) == [CARPETA_NUEVA] * sum(
        peticion == CREAR_CARPETA for peticion, _ in al_crearla
    )
    assert graph_simulado.esperas == esperas


@pytest.mark.parametrize(
    "hasta_el_fichero",
    [
        pytest.param(
            [(CARPETA, httpx.Response(200, json=FICHERO_IT))],
            id="lo-encuentra-el-get",
        ),
        pytest.param(
            [
                (CARPETA, httpx.Response(404)),
                (CREAR_CARPETA, httpx.Response(409, json=YA_EXISTE)),
                (CARPETA, httpx.Response(200, json=FICHERO_IT)),
            ],
            id="choca-con-el-al-crearla",
        ),
    ],
)
def test_si_con_el_nombre_de_la_carpeta_hay_un_fichero_no_se_sube(
    graph_simulado, tmp_path, hasta_el_fichero
):
    """En el drive, `it` es un fichero. Antes se daba por buena la carpeta, la
    encontrara el GET o chocara con ella el POST, y la subida fallaba después,
    con un error de Graph que no decía por qué. Ahora falla aquí y lo dice, sin
    intentar la subida."""
    graph_simulado.respuestas = [respuesta for _, respuesta in hasta_el_fichero]

    error = llamar_por_mcp("sharepoint_upload_file", subida(tmp_path, 9))

    assert error is not None
    assert "en el drive, it es un fichero y no una carpeta: no se ha subido" in error
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        peticion for peticion, _ in hasta_el_fichero
    ]
    # con `replace`, el POST se habría llevado por delante el fichero
    assert carpetas_creadas(graph_simulado.peticiones) == [CARPETA_NUEVA] * sum(
        peticion == CREAR_CARPETA for peticion, _ in hasta_el_fichero
    )


@pytest.mark.parametrize(
    ("al_crearla", "dice", "esperas"),
    [
        pytest.param([httpx.Response(403)], "403 Forbidden", [], id="sin-permiso"),
        pytest.param(
            [httpx.Response(503) for _ in range(4)],
            "503 Service Unavailable",
            [1, 2, 4],
            id="graph-sigue-con-503",
        ),
    ],
)
def test_si_la_carpeta_no_se_puede_crear_no_se_sube(
    graph_simulado, tmp_path, al_crearla, dice, esperas
):
    """Al crearla, solo el 409 se da por bueno. Un 403, o un 503 que sigue tras
    los tres reintentos, llega a radon-api tal cual, sin intentar la subida."""
    graph_simulado.respuestas = [httpx.Response(404), *al_crearla]

    error = llamar_por_mcp("sharepoint_upload_file", subida(tmp_path, 9))

    assert error is not None
    assert dice in error
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        CARPETA,
        *[CREAR_CARPETA] * len(al_crearla),
    ]
    assert carpetas_creadas(graph_simulado.peticiones) == [CARPETA_NUEVA] * len(
        al_crearla
    )
    assert graph_simulado.esperas == esperas
