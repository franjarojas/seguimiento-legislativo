"""
Genera el token OAuth de Gmail para guardar como GitHub Secret.

Instrucciones:
  1. Ve a https://console.cloud.google.com
  2. Crea un proyecto (o usa uno existente)
  3. Activa la API "Gmail API"
  4. Crea credenciales → OAuth 2.0 → Aplicación de escritorio
  5. Descarga el JSON de credenciales y guárdalo como credentials.json en esta carpeta
  6. Corre: python scripts/generar_token_gmail.py
  7. Se abrirá el navegador para autorizar → copia el token impreso
  8. En GitHub → Settings → Secrets → New secret:
       Nombre: GMAIL_TOKEN_JSON
       Valor:  (pegar el JSON impreso)
"""

import json
import os
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import google.oauth2.credentials

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
CREDS_FILE = "credentials.json"
TOKEN_FILE  = "token_gmail.json"

def main():
    if not os.path.exists(CREDS_FILE):
        print(f"No se encontró {CREDS_FILE}")
        print("Descárgalo desde: https://console.cloud.google.com → APIs → Credenciales")
        return

    flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
    creds = flow.run_local_server(port=0)

    token_data = {
        "token":         creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri":     creds.token_uri,
        "client_id":     creds.client_id,
        "client_secret": creds.client_secret,
        "scopes":        list(creds.scopes),
    }

    with open(TOKEN_FILE, "w") as f:
        json.dump(token_data, f, indent=2)

    print("\n" + "="*60)
    print("✓ Token generado. Copia todo el JSON de abajo como valor")
    print("  del secret GMAIL_TOKEN_JSON en GitHub:")
    print("="*60)
    print(json.dumps(token_data, indent=2))
    print("="*60)
    print(f"\nTambién guardado en: {TOKEN_FILE}")
    print("⚠️  No subas credentials.json ni token_gmail.json a Git")

if __name__ == "__main__":
    main()
