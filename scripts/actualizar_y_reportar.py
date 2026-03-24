"""
Actualiza el Excel de seguimiento legislativo y envía reporte de cambios por correo.

Flujo:
  1. Carga estado_anterior.json (snapshot de la semana pasada)
  2. Scraping: actualiza Col3, Col4, Col6 de cada boletín via tramitacion.senado.cl
  3. Compara con estado anterior → detecta cambios
  4. Guarda estado_actual.json (pasa a ser el nuevo estado_anterior.json)
  5. Si hay cambios (o es la primera ejecución): envía correo HTML con tabla de cambios + Excel adjunto

Variables de entorno requeridas (GitHub Secrets):
  GMAIL_TOKEN_JSON   — contenido del token OAuth de Gmail (ver README)
  DESTINATARIOS      — emails separados por coma, ej: "a@x.cl,b@x.cl"

Uso local:
  pip install requests openpyxl beautifulsoup4 google-auth google-auth-oauthlib google-api-python-client
  python scripts/actualizar_y_reportar.py
  python scripts/actualizar_y_reportar.py --solo-reporte   # no hace scraping, solo envía correo
"""

import os
import re
import sys
import json
import time
import base64
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# ── Rutas ──────────────────────────────────────────────────────────────────
ROOT            = Path(__file__).parent.parent
EXCEL_PATH      = ROOT / "seguimiento_legislativo.xlsx"
ESTADO_PATH     = ROOT / "estado_anterior.json"

# ── Configuración ──────────────────────────────────────────────────────────
SEMANAS_ALERTA  = 4   # marcar en el correo si un proyecto lleva N semanas sin cambios
BASE_SENADO     = "https://tramitacion.senado.cl/appsenado/index.php"
HEADERS         = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://tramitacion.senado.cl/appsenado/templates/tramitacion/index.php",
    "X-Requested-With": "XMLHttpRequest",
}

# Índices de columna (1-based)
COL_BOLETIN   = 1
COL_NOMBRE    = 2
COL_FECHA_ULT = 3
COL_ULT_MOV   = 4
COL_FECHA_PRS = 5
COL_ESTADO    = 6


# ══════════════════════════════════════════════════════════════════════════════
# SCRAPING
# ══════════════════════════════════════════════════════════════════════════════

def ts():
    return int(time.time() * 1000)


def normalizar_boletin(val: str) -> str:
    val = val.strip().replace(".", "")
    if "-" not in val and len(val) > 2:
        val = val[:-2] + "-" + val[-2:]
    return val


def obtener_proyid(boletin: str) -> str | None:
    params = {"mo": "tramitacion", "ac": "datos_proy", "nboletin": boletin, "etc": ts()}
    try:
        r = requests.get(BASE_SENADO, params=params, headers=HEADERS, timeout=15)
        m = re.search(r'proyid\s*=\s*(\d+)', r.text)
        return m.group(1) if m else None
    except Exception:
        return None


def obtener_ultimo_tramite(proyid: str, boletin: str) -> dict:
    params = {
        "mo": "tramitacion", "ac": "tramites",
        "proyid": proyid, "camara": "S",
        "boletin": boletin, "etc": ts()
    }
    try:
        r = requests.get(BASE_SENADO, params=params, headers=HEADERS, timeout=15)
    except Exception as e:
        return {"error": str(e)}

    soup = BeautifulSoup(r.text, "html.parser")
    tabla = soup.find("table", id="grid_tram") or soup.find("table")
    if not tabla:
        return {"error": "No se encontró tabla"}

    filas_datos = []
    for fila in tabla.find_all("tr"):
        celdas = fila.find_all("td")
        if len(celdas) >= 3:
            textos = [c.get_text(strip=True) for c in celdas]
            if len(textos) > 1 and re.match(r'\d{2}/\d{2}/\d{4}', textos[1]):
                filas_datos.append(textos)

    if not filas_datos:
        return {"error": "No se encontraron filas con fechas"}

    ultimo = filas_datos[-1]
    fecha  = ultimo[1] if len(ultimo) > 1 else "—"
    desc   = ultimo[2] if len(ultimo) > 2 else "—"
    etapa  = ultimo[3] if len(ultimo) > 3 else ""
    camara = ultimo[4] if len(ultimo) > 4 else ""

    estado = " — ".join(p for p in [etapa, camara] if p) or etapa
    descripcion = " — ".join(p for p in [etapa, desc] if p) if etapa else desc

    return {"fecha": fecha, "descripcion": descripcion, "estado": estado}


def scraping_boletin(boletin: str) -> dict:
    proyid = obtener_proyid(boletin)
    if not proyid:
        return {"error": "No se pudo obtener proyid"}
    time.sleep(0.3)
    return obtener_ultimo_tramite(proyid, boletin)


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def leer_estado_excel(ruta: Path) -> dict:
    """Devuelve dict {boletin: {nombre, fecha_ult, ult_mov, estado}} para comparación."""
    wb = load_workbook(ruta)
    ws = wb.active
    estado = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        boletin = str(row[COL_BOLETIN - 1] or "").strip()
        if not boletin or boletin == "—":
            continue
        estado[boletin] = {
            "nombre":    str(row[COL_NOMBRE - 1]    or ""),
            "fecha_ult": str(row[COL_FECHA_ULT - 1] or ""),
            "ult_mov":   str(row[COL_ULT_MOV - 1]   or ""),
            "estado":    str(row[COL_ESTADO - 1]     or ""),
        }
    return estado


def actualizar_excel(ruta: Path) -> dict:
    """Corre scraping y actualiza el Excel. Devuelve {boletin: resultado_scraping}."""
    wb = load_workbook(ruta)
    ws = wb.active
    font_d = Font(name="Arial", size=9)
    a_wrap = Alignment(vertical="top", wrap_text=True)
    a_ctr  = Alignment(horizontal="center", vertical="top", wrap_text=True)

    resultados = {}
    for row in ws.iter_rows(min_row=2):
        val_raw = str(row[COL_BOLETIN - 1].value or "").strip()
        if not val_raw or val_raw == "—":
            continue
        boletin = normalizar_boletin(val_raw)
        print(f"  {boletin}...", end=" ", flush=True)
        r = scraping_boletin(boletin)
        resultados[val_raw] = r

        if "error" in r:
            print(f"ERROR: {r['error']}")
            continue

        row[COL_FECHA_ULT - 1].value     = r["fecha"]
        row[COL_FECHA_ULT - 1].font      = font_d
        row[COL_FECHA_ULT - 1].alignment = a_ctr
        row[COL_ULT_MOV - 1].value       = r["descripcion"]
        row[COL_ULT_MOV - 1].font        = font_d
        row[COL_ULT_MOV - 1].alignment   = a_wrap
        row[COL_ESTADO - 1].value         = r["estado"]
        row[COL_ESTADO - 1].font          = font_d
        row[COL_ESTADO - 1].alignment     = a_wrap
        print(f"OK → {r['fecha']}")

    wb.save(ruta)
    return resultados


# ══════════════════════════════════════════════════════════════════════════════
# COMPARACIÓN DE CAMBIOS
# ══════════════════════════════════════════════════════════════════════════════

def detectar_cambios(estado_anterior: dict, estado_nuevo: dict) -> list[dict]:
    """
    Retorna lista de boletines con cambios.
    Cada item: {boletin, nombre, campo_cambiado, valor_anterior, valor_nuevo, fecha_ult, semanas_sin_cambio}
    """
    cambios = []
    hoy = datetime.date.today()

    for boletin, nuevo in estado_nuevo.items():
        anterior = estado_anterior.get(boletin, {})

        # Calcular semanas sin cambio
        semanas_sin_cambio = None
        try:
            d = datetime.datetime.strptime(nuevo["fecha_ult"], "%d/%m/%Y").date()
            semanas_sin_cambio = (hoy - d).days // 7
        except Exception:
            pass

        cambio = {
            "boletin": boletin,
            "nombre": nuevo["nombre"][:70] + "…" if len(nuevo["nombre"]) > 70 else nuevo["nombre"],
            "fecha_ult": nuevo["fecha_ult"],
            "ult_mov": nuevo["ult_mov"],
            "estado_nuevo": nuevo["estado"],
            "estado_anterior": anterior.get("estado", "—"),
            "semanas_sin_cambio": semanas_sin_cambio,
            "es_nuevo": boletin not in estado_anterior,
        }

        # ¿Hubo cambio en los campos de tramitación?
        if anterior:
            hubo_cambio = (
                anterior.get("fecha_ult") != nuevo["fecha_ult"] or
                anterior.get("ult_mov")   != nuevo["ult_mov"]   or
                anterior.get("estado")    != nuevo["estado"]
            )
            cambio["hubo_cambio"] = hubo_cambio
        else:
            cambio["hubo_cambio"] = True  # nuevo boletín

        cambios.append(cambio)

    return cambios


# ══════════════════════════════════════════════════════════════════════════════
# CORREO HTML
# ══════════════════════════════════════════════════════════════════════════════

SEMANAS_RECIENTE = 4  # destacar proyectos con movimiento en las últimas N semanas


def construir_html(cambios: list[dict], fecha_ejecucion: str) -> str:
    solo_cambios   = [c for c in cambios if c["hubo_cambio"]]
    sin_cambio     = [c for c in cambios if not c["hubo_cambio"]]

    # Proyectos con movimiento reciente (últimas 2 semanas), excluyendo los que
    # ya aparecen en "cambios esta semana"
    con_movimiento_reciente = [
        c for c in sin_cambio
        if c["semanas_sin_cambio"] is not None
        and c["semanas_sin_cambio"] <= SEMANAS_RECIENTE
    ]

    con_alerta = [
        c for c in sin_cambio
        if c["semanas_sin_cambio"] is not None
        and c["semanas_sin_cambio"] >= SEMANAS_ALERTA
    ]

    def fila_cambio(c):
        estado_ant = c["estado_anterior"] if c["estado_anterior"] != "—" else "<em>nuevo</em>"
        flecha = "🆕" if c["es_nuevo"] else "🔄"
        return f"""
        <tr>
          <td style="padding:8px;border-bottom:1px solid #e2e8f0;font-weight:600;white-space:nowrap">{flecha} {c['boletin']}</td>
          <td style="padding:8px;border-bottom:1px solid #e2e8f0;font-size:13px">{c['nombre']}</td>
          <td style="padding:8px;border-bottom:1px solid #e2e8f0;white-space:nowrap;color:#64748b">{c['fecha_ult']}</td>
          <td style="padding:8px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#dc2626">{estado_ant}</td>
          <td style="padding:8px;border-bottom:1px solid #e2e8f0;font-size:13px;color:#16a34a">{c['estado_nuevo']}</td>
        </tr>"""

    def fila_reciente(c):
        return f"""
        <tr>
          <td style="padding:8px;border-bottom:1px solid #dcfce7;font-weight:600;white-space:nowrap">🟢 {c['boletin']}</td>
          <td style="padding:8px;border-bottom:1px solid #dcfce7;font-size:13px">{c['nombre']}</td>
          <td style="padding:8px;border-bottom:1px solid #dcfce7;white-space:nowrap">{c['fecha_ult']}</td>
          <td style="padding:8px;border-bottom:1px solid #dcfce7;font-size:13px">{c['ult_mov']}</td>
          <td style="padding:8px;border-bottom:1px solid #dcfce7;color:#15803d">
            Hace {c['semanas_sin_cambio']} semana{'s' if c['semanas_sin_cambio'] != 1 else ''}
          </td>
        </tr>"""

    def fila_alerta(c):
        return f"""
        <tr>
          <td style="padding:8px;border-bottom:1px solid #fef3c7;font-weight:600;white-space:nowrap">⚠️ {c['boletin']}</td>
          <td style="padding:8px;border-bottom:1px solid #fef3c7;font-size:13px">{c['nombre']}</td>
          <td style="padding:8px;border-bottom:1px solid #fef3c7;white-space:nowrap">{c['fecha_ult']}</td>
          <td style="padding:8px;border-bottom:1px solid #fef3c7;color:#92400e">{c['semanas_sin_cambio']} semanas sin movimiento</td>
        </tr>"""

    # ── Sección cambios esta semana ──
    tabla_cambios = ""
    if solo_cambios:
        filas = "".join(fila_cambio(c) for c in solo_cambios)
        tabla_cambios = f"""
        <h2 style="color:#1e3a5f;font-size:16px;margin-top:28px">
          🔄 Boletines con cambios esta semana ({len(solo_cambios)})
        </h2>
        <table style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px">
          <thead>
            <tr style="background:#1e3a5f;color:white">
              <th style="padding:10px;text-align:left">Boletín</th>
              <th style="padding:10px;text-align:left">Nombre</th>
              <th style="padding:10px;text-align:left">Fecha</th>
              <th style="padding:10px;text-align:left">Estado anterior</th>
              <th style="padding:10px;text-align:left">Estado nuevo</th>
            </tr>
          </thead>
          <tbody>{filas}</tbody>
        </table>"""
    else:
        tabla_cambios = """
        <p style="color:#64748b;font-style:italic;margin-top:20px">
          ✅ No hubo cambios de tramitación esta semana.
        </p>"""

    # ── Sección movimiento reciente (últimas 2 semanas) ──
    tabla_reciente = ""
    if con_movimiento_reciente:
        filas_r = "".join(fila_reciente(c) for c in con_movimiento_reciente)
        tabla_reciente = f"""
        <h2 style="color:#15803d;font-size:16px;margin-top:28px">
          🟢 Con movimiento reciente — último mes ({len(con_movimiento_reciente)})
        </h2>
        <table style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;background:#f0fdf4">
          <thead>
            <tr style="background:#16a34a;color:white">
              <th style="padding:10px;text-align:left">Boletín</th>
              <th style="padding:10px;text-align:left">Nombre</th>
              <th style="padding:10px;text-align:left">Fecha</th>
              <th style="padding:10px;text-align:left">Último movimiento</th>
              <th style="padding:10px;text-align:left">Antigüedad</th>
            </tr>
          </thead>
          <tbody>{filas_r}</tbody>
        </table>"""

    # ── Sección alertas de inactividad ──
    tabla_alertas = ""
    if con_alerta:
        filas_a = "".join(fila_alerta(c) for c in con_alerta)
        tabla_alertas = f"""
        <h2 style="color:#92400e;font-size:16px;margin-top:28px">
          ⚠️ Proyectos sin movimiento hace {SEMANAS_ALERTA}+ semanas ({len(con_alerta)})
        </h2>
        <table style="width:100%;border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px;background:#fffbeb">
          <thead>
            <tr style="background:#fbbf24;color:#1c1917">
              <th style="padding:10px;text-align:left">Boletín</th>
              <th style="padding:10px;text-align:left">Nombre</th>
              <th style="padding:10px;text-align:left">Último movimiento</th>
              <th style="padding:10px;text-align:left">Inactividad</th>
            </tr>
          </thead>
          <tbody>{filas_a}</tbody>
        </table>"""

    return f"""<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;max-width:860px;margin:0 auto;padding:20px;color:#1e293b">
  <div style="background:#1e3a5f;padding:20px 28px;border-radius:8px 8px 0 0">
    <h1 style="color:white;margin:0;font-size:20px">
      📋 Reporte Seguimiento Legislativo
    </h1>
    <p style="color:#93c5fd;margin:6px 0 0">
      Actualización semanal — {fecha_ejecucion}
    </p>
  </div>
  <div style="background:#f8fafc;padding:20px 28px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 8px 8px">
    <p style="margin:0;font-size:14px;color:#475569">
      Se monitorearon <strong>{len(cambios)} boletines</strong>.
      Adjunto encontrarás el Excel actualizado.
    </p>
    {tabla_cambios}
    {tabla_reciente}
    {tabla_alertas}
    <hr style="border:none;border-top:1px solid #e2e8f0;margin-top:32px">
    <p style="font-size:12px;color:#94a3b8;margin:12px 0 0">
      Generado automáticamente por el sistema de seguimiento legislativo del Proyecto REDAR.
      Fuente: tramitacion.senado.cl / camara.cl
    </p>
  </div>
</body>
</html>"""


# ══════════════════════════════════════════════════════════════════════════════
# ENVÍO POR GMAIL API
# ══════════════════════════════════════════════════════════════════════════════

def enviar_correo(html_body: str, excel_path: Path, destinatarios: list[str], fecha: str):
    import google.oauth2.credentials
    from googleapiclient.discovery import build

    token_json = os.environ.get("GMAIL_TOKEN_JSON")
    if not token_json:
        print("⚠️  GMAIL_TOKEN_JSON no configurado — correo no enviado")
        return

    token_data = json.loads(token_json)
    creds = google.oauth2.credentials.Credentials(
        token=token_data.get("token"),
        refresh_token=token_data.get("refresh_token"),
        token_uri=token_data.get("token_uri", "https://oauth2.googleapis.com/token"),
        client_id=token_data.get("client_id"),
        client_secret=token_data.get("client_secret"),
        scopes=token_data.get("scopes", ["https://www.googleapis.com/auth/gmail.send"]),
    )

    service = build("gmail", "v1", credentials=creds)

    msg = MIMEMultipart("mixed")
    msg["Subject"] = f"Seguimiento Legislativo — {fecha}"
    msg["To"] = ", ".join(destinatarios)

    # Cuerpo HTML
    msg.attach(MIMEText(html_body, "html", "utf-8"))

    # Excel adjunto
    with open(excel_path, "rb") as f:
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{excel_path.name}"')
    msg.attach(part)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    service.users().messages().send(userId="me", body={"raw": raw}).execute()
    print(f"✉️  Correo enviado a: {', '.join(destinatarios)}")


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    solo_reporte = "--solo-reporte" in sys.argv
    fecha_hoy = datetime.date.today().strftime("%d/%m/%Y")

    print(f"\n{'='*60}")
    print(f"Seguimiento Legislativo — {fecha_hoy}")
    print(f"{'='*60}\n")

    # 1. Cargar estado anterior
    estado_anterior = {}
    if ESTADO_PATH.exists():
        with open(ESTADO_PATH) as f:
            estado_anterior = json.load(f)
        print(f"Estado anterior cargado ({len(estado_anterior)} boletines)\n")
    else:
        print("Primera ejecución — no hay estado anterior\n")

    # 2. Scraping y actualización del Excel
    if not solo_reporte:
        print("Actualizando Excel...")
        actualizar_excel(EXCEL_PATH)
        print()

    # 3. Leer estado nuevo del Excel ya actualizado
    estado_nuevo = leer_estado_excel(EXCEL_PATH)

    # 4. Guardar nuevo estado como referencia para la próxima semana
    with open(ESTADO_PATH, "w") as f:
        json.dump(estado_nuevo, f, ensure_ascii=False, indent=2)
    print(f"Estado guardado en {ESTADO_PATH}\n")

    # 5. Detectar cambios
    cambios = detectar_cambios(estado_anterior, estado_nuevo)
    n_cambios = sum(1 for c in cambios if c["hubo_cambio"])
    print(f"Cambios detectados: {n_cambios} de {len(cambios)} boletines\n")

    # 6. Construir y enviar correo
    destinatarios_raw = os.environ.get("DESTINATARIOS", "")
    destinatarios = [d.strip() for d in destinatarios_raw.split(",") if d.strip()]

    if not destinatarios:
        print("⚠️  Variable DESTINATARIOS no configurada — correo no enviado")
        print("   Configura: export DESTINATARIOS='tu@email.cl,otro@email.cl'")
    else:
        html = construir_html(cambios, fecha_hoy)
        enviar_correo(html, EXCEL_PATH, destinatarios, fecha_hoy)

    print("\n✓ Proceso completado")


if __name__ == "__main__":
    main()
