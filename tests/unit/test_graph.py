"""`graph.request` cuando Graph falla: qué se repite y qué no."""

import httpx
import pytest

from microsoft_mcp import graph, tools


def test_un_get_que_da_503_se_reintenta_y_acaba_bien(graph_simulado):
    """La carpeta que busca `sharepoint_upload_file` antes de subir: un GET, que
    se puede repetir sin duplicar nada."""
    graph_simulado.respuestas = [
        httpx.Response(503),
        httpx.Response(200, json={"id": "carpeta-1", "folder": {}}),
    ]

    resultado = graph.request("GET", "/drives/drive-1/root:/informes", "cuenta-1")

    assert resultado == {"id": "carpeta-1", "folder": {}}
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("GET", "/v1.0/drives/drive-1/root:/informes"),
        ("GET", "/v1.0/drives/drive-1/root:/informes"),
    ]
    assert graph_simulado.esperas == [1]


def test_un_post_que_da_502_no_se_repite(graph_simulado):
    """La carpeta que crea `sharepoint_upload_file` cuando falta: Graph pudo
    crearla antes de contestar el 502. El POST no se repite y el 502 sube a
    quien llama."""
    graph_simulado.respuestas = [httpx.Response(502)]

    with pytest.raises(httpx.HTTPStatusError) as error:
        graph.request(
            "POST",
            "/drives/drive-1/root/children",
            "cuenta-1",
            json={
                "name": "informes",
                "folder": {},
                "@microsoft.graph.conflictBehavior": "fail",
            },
        )

    assert error.value.response.status_code == 502
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/drives/drive-1/root/children"),
    ]
    assert graph_simulado.esperas == []


@pytest.mark.parametrize("con_adjunto", [False, True], ids=["aviso", "impresora"])
def test_send_email_hace_un_solo_post_aunque_graph_de_502(
    graph_simulado, tmp_path, con_adjunto
):
    """Por la función de la herramienta, como la usan los dos envíos de
    radon-api: el aviso interno, sin adjuntos, y la impresora, con el PDF
    adjunto. Si Graph lo mandó antes de contestar el 502, repetirlo sería un
    segundo aviso o una segunda impresión."""
    adjuntos = None
    if con_adjunto:
        pdf = tmp_path / "plano.pdf"
        pdf.write_bytes(b"%PDF-1.7\n")
        adjuntos = [str(pdf)]
    graph_simulado.respuestas = [httpx.Response(502)]

    with pytest.raises(httpx.HTTPStatusError) as error:
        tools.send_email.fn(
            account_id="cuenta-1",
            to="destino@example.com",
            subject="Asunto",
            body="Cuerpo",
            attachments=adjuntos,
        )

    assert error.value.response.status_code == 502
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/me/sendMail"),
    ]
    assert graph_simulado.esperas == []


def test_un_429_se_reintenta_aunque_sea_post_y_respeta_retry_after(graph_simulado):
    """Con un 429, Graph rechazó la petición sin procesarla: repetirla no
    duplica nada, después de esperar lo que pide."""
    graph_simulado.respuestas = [
        httpx.Response(429, headers={"Retry-After": "7"}),
        httpx.Response(202),
    ]

    resultado = graph.request(
        "POST", "/me/sendMail", "cuenta-1", json={"message": {"subject": "Aviso"}}
    )

    assert resultado is None
    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        ("POST", "/v1.0/me/sendMail"),
        ("POST", "/v1.0/me/sendMail"),
    ]
    assert graph_simulado.esperas == [7]


@pytest.mark.parametrize(
    ("metodo", "ruta"),
    [
        pytest.param("PUT", "/drives/drive-1/root:/it/9187497.pdf:/content", id="PUT"),
        pytest.param("PATCH", "/me/messages/mensaje-1", id="PATCH"),
        pytest.param("DELETE", "/me/messages/mensaje-1", id="DELETE"),
    ],
)
def test_put_patch_y_delete_se_reintentan_tras_un_5xx(graph_simulado, metodo, ruta):
    """Repetirlos deja lo mismo que hacerlos una vez: el fichero en su ruta, las
    mismas propiedades en el mensaje, el mensaje en la papelera."""
    graph_simulado.respuestas = [httpx.Response(502), httpx.Response(204)]

    graph.request(metodo, ruta, "cuenta-1")

    assert [(p.method, p.url.path) for p in graph_simulado.peticiones] == [
        (metodo, f"/v1.0{ruta}"),
        (metodo, f"/v1.0{ruta}"),
    ]
    assert graph_simulado.esperas == [1]


def test_un_5xx_que_no_se_arregla_acaba_en_error_tras_tres_reintentos(graph_simulado):
    """Tres reintentos, tras 1, 2 y 4 s, y el último 503 sube a quien llama."""
    graph_simulado.respuestas = [httpx.Response(503) for _ in range(4)]

    with pytest.raises(httpx.HTTPStatusError) as error:
        graph.request("GET", "/drives/drive-1/root:/informes", "cuenta-1")

    assert error.value.response.status_code == 503
    assert len(graph_simulado.peticiones) == 4
    assert graph_simulado.esperas == [1, 2, 4]
