"""Gera o refresh token do OneDrive PESSOAL (rode UMA vez no seu computador).

Uso:
    python get_token.py SEU_CLIENT_ID

O script mostra um código e um link (https://microsoft.com/devicelogin). Abra o link,
digite o código e entre com a conta Microsoft dona do OneDrive. No final ele imprime o
refresh token para colar em [graph].refresh_token no secrets.toml / Secrets do Streamlit.

Guarde o token como uma senha: ele dá acesso de leitura/escrita aos arquivos do OneDrive.
"""

from __future__ import annotations

import sys

import msal

SCOPES = ["Files.ReadWrite"]  # "offline_access" (refresh token) é incluído automaticamente
AUTHORITY = "https://login.microsoftonline.com/consumers"


def main() -> None:
    if len(sys.argv) != 2:
        print("Uso: python get_token.py SEU_CLIENT_ID")
        sys.exit(1)

    app = msal.PublicClientApplication(sys.argv[1], authority=AUTHORITY)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        print("Não foi possível iniciar o login:", flow.get("error_description", flow))
        print("Confira se o app permite 'public client flows' e contas Microsoft pessoais.")
        sys.exit(1)

    print(flow["message"])  # instruções: link + código
    result = app.acquire_token_by_device_flow(flow)  # aguarda você concluir o login

    if "refresh_token" not in result:
        print("Falha no login:", result.get("error_description", result))
        sys.exit(1)

    print("\nLogin OK! Cole o valor abaixo em [graph].refresh_token:\n")
    print(result["refresh_token"])


if __name__ == "__main__":
    main()
