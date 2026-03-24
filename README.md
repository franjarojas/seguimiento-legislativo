# Seguimiento Legislativo — Automatización

Sistema de actualización semanal del Excel de seguimiento y envío de reporte por correo.

---

## Estructura del repositorio

```
seguimiento-legislativo/
├── seguimiento_legislativo.xlsx   ← Excel de seguimiento (se actualiza automáticamente)
├── estado_anterior.json           ← Snapshot de la semana anterior (generado automáticamente)
├── scripts/
│   ├── actualizar_y_reportar.py   ← Script principal
│   └── generar_token_gmail.py     ← Utilitario para generar credenciales Gmail
├── .github/
│   └── workflows/
│       └── actualizacion_semanal.yml  ← Workflow de GitHub Actions
└── README.md
```

---

## Configuración inicial (una sola vez)

### Paso 1 — Crear el repositorio en GitHub

1. Ve a https://github.com/new
2. Nombre sugerido: `seguimiento-legislativo`
3. **Privado** (importante: contiene datos internos)
4. Sube todos los archivos de esta carpeta

### Paso 2 — Configurar Gmail API

1. Ve a https://console.cloud.google.com
2. Crea un proyecto nuevo (ej: "Seguimiento Legislativo")
3. Activa la **Gmail API**: APIs y servicios → Biblioteca → buscar "Gmail API" → Activar
4. Crea credenciales: APIs → Credenciales → Crear credenciales → **ID de cliente OAuth 2.0**
   - Tipo: **Aplicación de escritorio**
   - Descarga el JSON → guárdalo como `credentials.json` en la carpeta `scripts/`
5. Corre el script de autorización en tu computador:
   ```bash
   pip install google-auth-oauthlib google-auth google-api-python-client
   python scripts/generar_token_gmail.py
   ```
6. Se abrirá el navegador → autoriza con tu cuenta Gmail
7. El script imprimirá el token JSON completo

### Paso 3 — Configurar GitHub Secrets

En tu repositorio GitHub → **Settings → Secrets and variables → Actions → New repository secret**:

| Nombre | Valor |
|--------|-------|
| `GMAIL_TOKEN_JSON` | El JSON completo impreso por `generar_token_gmail.py` |
| `DESTINATARIOS` | Emails separados por coma, ej: `francisco@nodoxxi.cl,jose@nodoxxi.cl` |

---

## Funcionamiento

El workflow corre **cada lunes a las 11:00 AM (hora Chile)**:

1. Consulta `tramitacion.senado.cl` para cada boletín del Excel
2. Actualiza **Col3** (fecha último movimiento), **Col4** (descripción) y **Col6** (estado)
3. Compara con el `estado_anterior.json` de la semana pasada
4. Envía un correo HTML con:
   - Tabla de boletines con **cambios** respecto a la semana anterior
   - Alerta de proyectos **sin movimiento hace 4+ semanas**
   - **Excel actualizado** como adjunto
5. Hace commit del Excel y estado actualizado al repositorio

### Correr manualmente

En GitHub → Actions → **Seguimiento Legislativo Semanal** → **Run workflow**

### Correr localmente

```bash
pip install requests openpyxl beautifulsoup4 google-auth google-auth-oauthlib google-api-python-client

# Actualización completa + correo
export GMAIL_TOKEN_JSON='{"token": "...", ...}'
export DESTINATARIOS='francisco@nodoxxi.cl,jose@nodoxxi.cl'
python scripts/actualizar_y_reportar.py

# Solo correo (sin scraping)
python scripts/actualizar_y_reportar.py --solo-reporte
```

---

## Notas técnicas

- El scraping usa los endpoints AJAX internos de `tramitacion.senado.cl`
- El token de Gmail se refresca automáticamente (no expira)
- Si un boletín falla el scraping, se conserva el valor anterior en el Excel
- El commit automático solo ocurre si hay cambios en el Excel o el estado
- Los proyectos sin información pública (ej: 17.446-17) mantienen sus datos manuales
