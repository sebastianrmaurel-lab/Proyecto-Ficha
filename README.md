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

## Notas para cuando lo subas a Render + Supabase

- SQLite vive en un archivo local. En Render el sistema de archivos es efímero (se borra en cada redeploy/reinicio), así que tanto la base de datos como las fotos/archivos adjuntos subidos a `static/uploads/` se perderían.
- Antes de desplegar, hay que migrar la base de datos a Postgres (Supabase) y mover los archivos a Supabase Storage. Avísame cuando quieras seguir con eso y lo armamos paso a paso.
