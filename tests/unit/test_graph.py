"""`graph.request` cuando Graph falla: qué se repite y qué no."""

import httpx

from microsoft_mcp import graph


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
