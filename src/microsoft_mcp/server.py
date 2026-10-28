import os
import sys
from dotenv import load_dotenv
from .tools import mcp


def main() -> None:
    # El .env se carga al arrancar el servidor y no al importar el paquete: las
    # pruebas lo importan, y no deben ver el de la máquina.
    load_dotenv()
    if not os.getenv("MICROSOFT_MCP_CLIENT_ID"):
        print(
            "Error: MICROSOFT_MCP_CLIENT_ID environment variable is required",
            file=sys.stderr,
        )
        sys.exit(1)

    mcp.run()


if __name__ == "__main__":
    main()
