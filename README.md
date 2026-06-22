# Fichas de Personal

App web en Flask para gestionar varias fichas de personal, cada una con su historial de actos administrativos y asignaciones familiares. Incluye login, auditoría, papelera, validaciones y un dashboard.

## Instalación

```bash
python -m venv venv
source venv/bin/activate      # En Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## Ejecutar

```bash
python app.py
```

Abre http://127.0.0.1:5000

Al primer arranque se crea un usuario administrador con el username y contraseña de las variables de entorno `ADMIN_USER` / `ADMIN_PASSWORD` (por defecto `admin` / `admin123`). Cámbialos antes de usar esto en serio.

## Qué incluye

- **Dashboard** (`/dashboard`): fichas activas, total de actos administrativos y asignaciones familiares, fichas en la papelera, y carpetas creadas por mes.
- **Fichas de Personal** (`/fichas`): listado con buscador, exportar todo a Excel, y entrar a cada ficha para ver/editar sus datos, foto, actos administrativos y asignación familiar en una sola pantalla. Agregar o editar un acto administrativo o una asignación familiar abre una ventana (modal) en vez de ocupar espacio fijo en la pantalla, y cada uno puede llevar varios archivos adjuntos a la vez (PDF, imagen o Word) — se pueden ver, agregar más después, o quitar uno por uno. El campo "Tipo de Acto Administrativo" sugiere valores comunes (Decreto, Resolución, Honorarios, etc.) pero acepta cualquier texto.
- **Auditoría**: cada creación, edición o eliminación queda registrada con quién y cuándo, visible en el dashboard.
- **Papelera** (`/fichas/papelera`): eliminar una ficha ya no es definitivo — se mueve a la papelera, desde donde se puede restaurar o eliminar para siempre.
- **Validaciones**: el RUT se valida con su dígito verificador real (no solo el formato), y en actos/asignaciones la fecha "Hasta" no puede ser anterior a "Desde".
- **Exportar**: Excel con todas las fichas activas, y PDF por persona con sus actos administrativos.

## Estructura del proyecto

```
ficha_personas/
├── app.py
├── requirements.txt
├── fichas.db                 # se crea solo
├── static/
│   ├── estilo.css
│   ├── rut.js
│   └── uploads/
└── templates/
    ├── base.html
    ├── login.html
    ├── dashboard.html
    ├── listado.html
    ├── papelera.html
    └── ficha.html
```

## Despliegue en Render

### Pasos

1. Sube este proyecto a GitHub (si no lo has hecho ya).
2. En Render → **New Web Service** → conecta tu repositorio.
3. Render detectará el `Procfile` y usará `gunicorn app:app` automáticamente.
4. En **Environment Variables** agrega estas tres (obligatorio antes del primer deploy):

| Variable | Ejemplo | Descripción |
|---|---|---|
| `SECRET_KEY` | `un-string-largo-y-aleatorio` | Clave de sesión Flask. Genera una con `python -c "import secrets; print(secrets.token_hex(32))"` |
| `ADMIN_USER` | `admin` | Usuario administrador inicial |
| `ADMIN_PASSWORD` | `una-clave-segura` | Contraseña administrador inicial |

5. Haz clic en **Deploy** — debería quedar en verde.

### Advertencias importantes para el plan gratuito de Render

- **El sistema de archivos es efímero**: la base de datos SQLite (`fichas.db`) y los archivos subidos (`static/uploads/`) se borran en cada redeploy o reinicio. Para producción real necesitas migrar a Postgres (Supabase) y Supabase Storage para los archivos.
- **El servicio se duerme tras 15 min de inactividad**: el primer request tarda ~50 segundos en responder (esto lo avisa la misma pantalla de Render).
- Por ahora sirve perfectamente para probar y mostrar el sistema. Avísame cuando quieras migrar a Supabase y lo hacemos juntos.

