"""El `.env` lo carga el servidor al arrancar, no la importación del paquete."""

import os
import subprocess
import sys

from microsoft_mcp import server


def test_importar_el_paquete_no_carga_ningun_env(tmp_path):
    """Las pruebas importan el paquete. Antes, `auth` cargaba al importarse el
    `.env` de la raíz del checkout, y en el de la máquina de Luis todo lo que
    trajera entraba en el entorno de las pruebas. En un proceso aparte, con un
    `.env` canario en lugar del que se encontrara, importar el servidor no lo
    carga."""
    canario = tmp_path / "canario" / ".env"
    canario.parent.mkdir()
    canario.write_text("MICROSOFT_MCP_CANARIO=si\n")
    codigo = (
        "import os\n"
        "import dotenv.main\n"
        f"dotenv.main.find_dotenv = lambda *args, **kwargs: {str(canario)!r}\n"
        "import microsoft_mcp.server\n"
        "print(os.environ.get('MICROSOFT_MCP_CANARIO'))\n"
    )

    resultado = subprocess.run(
        [sys.executable, "-c", codigo],
        cwd=tmp_path,
        capture_output=True,
        text=True,
        check=True,
    )

    assert resultado.stdout == "None\n"


class ServidorQueNoArranca:
    """Ocupa el sitio del `mcp` de FastMCP: apunta con qué client id arrancaría."""

    def __init__(self) -> None:
        self.arranques: list[str] = []

    def run(self) -> None:
        self.arranques.append(os.environ["MICROSOFT_MCP_CLIENT_ID"])


def test_el_servidor_carga_el_env_antes_de_pedir_el_client_id(monkeypatch):
    """Quien tiene el client id en el `.env`, y no en el entorno, sigue pudiendo
    arrancar el servidor: `main()` carga el `.env` antes de comprobarlo."""
    monkeypatch.delenv("MICROSOFT_MCP_CLIENT_ID")

    def cargar_env() -> bool:
        monkeypatch.setenv("MICROSOFT_MCP_CLIENT_ID", "client-id-del-env")
        return True

    servidor = ServidorQueNoArranca()
    monkeypatch.setattr(server, "load_dotenv", cargar_env)
    monkeypatch.setattr(server, "mcp", servidor)

    server.main()

    assert servidor.arranques == ["client-id-del-env"]
