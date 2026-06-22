import os
import sqlite3
import uuid
import json
from functools import wraps
from io import BytesIO
from datetime import date

from flask import (
    Flask, render_template, request, redirect, url_for, flash, g, session, send_file, jsonify
)
from werkzeug.security import generate_password_hash, check_password_hash
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'fichas.db')
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static', 'uploads')
EXT_IMAGEN = {'png', 'jpg', 'jpeg', 'gif', 'webp'}
EXT_DOCUMENTO = {'pdf', 'jpg', 'jpeg', 'png', 'doc', 'docx'}

# Usadas solo para crear el primer usuario administrador la primera vez que corre la app.
ADMIN_USER = os.environ.get('ADMIN_USER', 'admin')
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin123')

app = Flask(__name__)
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'cambia-esta-clave-por-una-segura')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 30 * 1024 * 1024  # máx 30 MB por solicitud (varios archivos a la vez)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def login_required(f):
    @wraps(f)
    def decorada(*args, **kwargs):
        if not session.get('usuario_id'):
            return redirect(url_for('login', next=request.path))
        return f(*args, **kwargs)
    return decorada


# ---------- Base de datos ----------

def get_db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
        g.db.execute('PRAGMA foreign_keys = ON')
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop('db', None)
    if db is not None:
        db.close()


def init_db():
    db = sqlite3.connect(DB_PATH)
    db.execute('PRAGMA foreign_keys = ON')

    db.execute('''
        CREATE TABLE IF NOT EXISTS fichas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            rut TEXT NOT NULL UNIQUE,
            titulo TEXT,
            grado_academico TEXT,
            jerarquia TEXT,
            direccion TEXT,
            observaciones TEXT,
            fecha_nacimiento TEXT,
            estado_civil TEXT,
            telefono TEXT,
            sexo TEXT,
            link TEXT,
            foto TEXT,
            activo INTEGER NOT NULL DEFAULT 1,
            fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    db.execute('''
        CREATE TABLE IF NOT EXISTS historial_institucional (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ficha_id INTEGER NOT NULL,
            cargo TEXT,
            unidad TEXT,
            departamento TEXT,
            tipo_contrato TEXT,
            estamento TEXT,
            grado TEXT,
            jornada TEXT,
            desde TEXT,
            hasta TEXT,
            observaciones TEXT,
            fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (ficha_id) REFERENCES fichas (id) ON DELETE CASCADE
        )
    ''')
    db.execute('''
        CREATE TABLE IF NOT EXISTS archivos_historial (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            historial_id INTEGER NOT NULL,
            nombre_archivo TEXT NOT NULL,
            nombre_original TEXT,
            fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (historial_id) REFERENCES historial_institucional (id) ON DELETE CASCADE
        )
    ''')
    db.execute('''
        CREATE TABLE IF NOT EXISTS actos_administrativos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ficha_id INTEGER NOT NULL,
            decreto TEXT,
            fecha TEXT,
            glosa TEXT,
            tipo_acto TEXT,
            desde TEXT,
            hasta TEXT,
            observaciones TEXT,
            soporte TEXT,
            fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (ficha_id) REFERENCES fichas (id) ON DELETE CASCADE
        )
    ''')
    db.execute('''
        CREATE TABLE IF NOT EXISTS asignaciones_familiares (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ficha_id INTEGER NOT NULL,
            nombre TEXT,
            parentesco TEXT,
            resolucion TEXT,
            fecha TEXT,
            desde TEXT,
            hasta TEXT,
            fecha_nacimiento TEXT,
            observaciones TEXT,
            soporte TEXT,
            fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (ficha_id) REFERENCES fichas (id) ON DELETE CASCADE
        )
    ''')
    db.execute('''
        CREATE TABLE IF NOT EXISTS archivos_actos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            acto_id INTEGER NOT NULL,
            nombre_archivo TEXT NOT NULL,
            nombre_original TEXT,
            fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (acto_id) REFERENCES actos_administrativos (id) ON DELETE CASCADE
        )
    ''')
    db.execute('''
        CREATE TABLE IF NOT EXISTS archivos_familiares (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            familiar_id INTEGER NOT NULL,
            nombre_archivo TEXT NOT NULL,
            nombre_original TEXT,
            fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (familiar_id) REFERENCES asignaciones_familiares (id) ON DELETE CASCADE
        )
    ''')
    db.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            nombre TEXT,
            activo INTEGER NOT NULL DEFAULT 1,
            fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    db.execute('''
        CREATE TABLE IF NOT EXISTS auditoria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario_id INTEGER,
            usuario_nombre TEXT,
            accion TEXT,
            entidad TEXT,
            entidad_id INTEGER,
            descripcion TEXT,
            fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    db.execute('''
        CREATE TABLE IF NOT EXISTS tipos_acto_administrativo (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL UNIQUE,
            fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    # Tabla genérica para catálogos editables (contrato, estamento, jornada)
    db.execute('''
        CREATE TABLE IF NOT EXISTS catalogos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            categoria TEXT NOT NULL,
            nombre TEXT NOT NULL,
            fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(categoria, nombre)
        )
    ''')

    if db.execute('SELECT COUNT(*) FROM usuarios').fetchone()[0] == 0:
        db.execute(
            'INSERT INTO usuarios (username, password_hash, nombre, activo) VALUES (?, ?, ?, 1)',
            (ADMIN_USER, generate_password_hash(ADMIN_PASSWORD), 'Administrador')
        )

    if db.execute('SELECT COUNT(*) FROM tipos_acto_administrativo').fetchone()[0] == 0:
        for nombre in ['Decreto', 'Resolución', 'Honorarios', 'Contrata', 'Nombramiento', 'Comisión de Servicio']:
            db.execute('INSERT INTO tipos_acto_administrativo (nombre) VALUES (?)', (nombre,))

    semillas_catalogos = {
        'contrato':  ['Planta', 'Contrata', 'Honorarios', 'Suplencia', 'Interino'],
        'estamento': ['Directivo', 'Profesional', 'Técnico', 'Administrativo', 'Auxiliar', 'Docente'],
        'jornada':   ['Completa (44h)', 'Media jornada (22h)', 'Parcial', 'Otra'],
    }
    for categoria, valores in semillas_catalogos.items():
        if db.execute('SELECT COUNT(*) FROM catalogos WHERE categoria=?', (categoria,)).fetchone()[0] == 0:
            for v in valores:
                db.execute('INSERT INTO catalogos (categoria, nombre) VALUES (?, ?)', (categoria, v))

    db.commit()
    db.close()


def extension_permitida(filename, permitidas):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in permitidas


def guardar_archivo(campo_form, permitidas):
    """Guarda un archivo subido y devuelve su nombre final, o None si no se subió nada."""
    archivo = request.files.get(campo_form)
    if archivo and archivo.filename and extension_permitida(archivo.filename, permitidas):
        extension = archivo.filename.rsplit('.', 1)[1].lower()
        nombre_final = f"{uuid.uuid4().hex}.{extension}"
        archivo.save(os.path.join(UPLOAD_FOLDER, nombre_final))
        return nombre_final
    return None


def guardar_archivos_multiples(campo_form, permitidas):
    """Guarda todos los archivos subidos bajo un campo <input multiple>.
    Devuelve una lista de tuplas (nombre_guardado, nombre_original)."""
    guardados = []
    for archivo in request.files.getlist(campo_form):
        if archivo and archivo.filename and extension_permitida(archivo.filename, permitidas):
            extension = archivo.filename.rsplit('.', 1)[1].lower()
            nombre_final = f"{uuid.uuid4().hex}.{extension}"
            archivo.save(os.path.join(UPLOAD_FOLDER, nombre_final))
            guardados.append((nombre_final, archivo.filename))
    return guardados


def _borrar_archivo_fisico(nombre_archivo):
    if not nombre_archivo:
        return
    ruta = os.path.join(UPLOAD_FOLDER, nombre_archivo)
    if os.path.exists(ruta):
        os.remove(ruta)


def rut_valido(rut):
    """Valida el dígito verificador de un RUT chileno (módulo 11)."""
    limpio = rut.upper().replace('.', '').replace('-', '').strip()
    if len(limpio) < 2:
        return False
    cuerpo, dv = limpio[:-1], limpio[-1]
    if not cuerpo.isdigit():
        return False
    suma = 0
    multiplicador = 2
    for digito in reversed(cuerpo):
        suma += int(digito) * multiplicador
        multiplicador = 2 if multiplicador == 7 else multiplicador + 1
    resto = 11 - (suma % 11)
    dv_esperado = {11: '0', 10: 'K'}.get(resto, str(resto))
    return dv == dv_esperado


def fechas_validas(desde, hasta):
    """True si no se puede determinar un problema, o si Desde <= Hasta."""
    if desde and hasta:
        try:
            return date.fromisoformat(desde) <= date.fromisoformat(hasta)
        except ValueError:
            return True
    return True


def _registrar_auditoria(accion, entidad, entidad_id, descripcion=''):
    db = get_db()
    db.execute('''
        INSERT INTO auditoria (usuario_id, usuario_nombre, accion, entidad, entidad_id, descripcion)
        VALUES (?, ?, ?, ?, ?, ?)
    ''', (session.get('usuario_id'), session.get('usuario_nombre', ''), accion, entidad, entidad_id, descripcion))
    db.commit()


MESES_ES = {
    '01': 'Ene', '02': 'Feb', '03': 'Mar', '04': 'Abr', '05': 'May', '06': 'Jun',
    '07': 'Jul', '08': 'Ago', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dic',
}


def _resumen_carpetas_por_mes():
    db = get_db()
    filas = db.execute('''
        SELECT strftime('%Y-%m', fecha_creacion) AS periodo, COUNT(*) AS total
        FROM fichas
        WHERE activo = 1
        GROUP BY periodo
        ORDER BY periodo
    ''').fetchall()

    meses = []
    for fila in filas:
        if not fila['periodo']:
            continue
        anio, mes = fila['periodo'].split('-')
        meses.append({'label': f"{MESES_ES.get(mes, mes)} {anio}", 'total': fila['total']})

    total_general = sum(m['total'] for m in meses)
    return meses, total_general


# ---------- Login ----------

@app.route('/')
def home():
    return redirect(url_for('dashboard'))


@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('usuario_id'):
        return redirect(url_for('dashboard'))

    if request.method == 'POST':
        username = request.form.get('usuario', '').strip()
        clave = request.form.get('clave', '').strip()
        db = get_db()
        usuario = db.execute(
            'SELECT * FROM usuarios WHERE username = ? AND activo = 1', (username,)
        ).fetchone()
        if usuario and check_password_hash(usuario['password_hash'], clave):
            session['usuario_id'] = usuario['id']
            session['usuario_nombre'] = usuario['nombre'] or usuario['username']
            siguiente = request.args.get('next') or url_for('dashboard')
            return redirect(siguiente)
        flash('Usuario o contraseña incorrectos.', 'error')

    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    flash('Sesión cerrada.', 'success')
    return redirect(url_for('login'))


# ---------- Dashboard ----------

@app.route('/dashboard')
@login_required
def dashboard():
    db = get_db()
    total_fichas = db.execute('SELECT COUNT(*) FROM fichas WHERE activo = 1').fetchone()[0]
    total_actos = db.execute('SELECT COUNT(*) FROM actos_administrativos').fetchone()[0]
    total_familiares = db.execute('SELECT COUNT(*) FROM asignaciones_familiares').fetchone()[0]
    en_papelera = db.execute('SELECT COUNT(*) FROM fichas WHERE activo = 0').fetchone()[0]

    meses_resumen, total_general_carpetas = _resumen_carpetas_por_mes()

    return render_template(
        'dashboard.html',
        total_fichas=total_fichas, total_actos=total_actos, total_familiares=total_familiares,
        en_papelera=en_papelera,
        meses_resumen=meses_resumen, total_general_carpetas=total_general_carpetas
    )


# ---------- Listado / papelera / exportar ----------

@app.route('/fichas')
@login_required
def fichas_listado():
    q = request.args.get('q', '').strip()
    db = get_db()
    sql = '''
        SELECT f.*, COUNT(a.id) AS total_actos
        FROM fichas f
        LEFT JOIN actos_administrativos a ON a.ficha_id = f.id
        WHERE f.activo = 1
    '''
    params = []
    if q:
        sql += ' AND (f.nombre LIKE ? OR f.rut LIKE ?) '
        params += [f'%{q}%', f'%{q}%']
    sql += ' GROUP BY f.id ORDER BY f.nombre COLLATE NOCASE'
    fichas = db.execute(sql, params).fetchall()

    en_papelera = db.execute('SELECT COUNT(*) FROM fichas WHERE activo = 0').fetchone()[0]

    return render_template('listado.html', fichas=fichas, q=q, en_papelera=en_papelera)


# ---------- Importación desde Excel ----------

def _leer_excel_funcionarios(archivo_stream):
    """
    Lee el Excel y devuelve una lista de personas, cada una con sus actos y un historial.
    Agrupa las filas por IdFuncionario (o Funcionario si no hay ID).
    """
    from openpyxl import load_workbook
    wb = load_workbook(archivo_stream, data_only=True)
    ws = wb.active

    encabezados_raw = [str(c.value).strip() if c.value else '' for c in next(ws.iter_rows(min_row=1, max_row=1))]

    def _col(fila, nombre):
        """Busca una columna por nombre de encabezado (case-insensitive, parcial)."""
        nombre_lower = nombre.lower()
        for i, enc in enumerate(encabezados_raw):
            if nombre_lower in enc.lower():
                return fila[i]
        return None

    def _str(val):
        if val is None:
            return ''
        if hasattr(val, 'strftime'):
            return val.strftime('%Y-%m-%d')
        return str(val).strip()

    personas = {}  # id_funcionario → dict

    for fila in ws.iter_rows(min_row=2, values_only=True):
        if not any(c is not None for c in fila):
            continue

        id_func = _str(_col(fila, 'IdFuncionario')) or _str(_col(fila, 'Funcionario'))
        if not id_func:
            continue

        if id_func not in personas:
            personas[id_func] = {
                'id_original': id_func,
                'nombre':          _str(_col(fila, 'Funcionario')),
                'fecha_nacimiento': _str(_col(fila, 'FechaNac')),
                'estado_civil':    _str(_col(fila, 'EstadoCivil')),
                'titulo':          _str(_col(fila, 'Titulo')),
                'direccion':       _str(_col(fila, 'Direccion')),
                'telefono':        _str(_col(fila, 'Telefono')),
                'link':            _str(_col(fila, 'Correo')),
                'rut': '',  # debe rellenarse a mano
                'actos': [],
                'historial': [],
            }

        decreto  = _str(_col(fila, 'Decreto'))
        glosa    = _str(_col(fila, 'GLOSA'))
        calidad  = _str(_col(fila, 'Calidad'))
        fecha    = _str(_col(fila, 'Fecha'))
        desde    = _str(_col(fila, 'Desde'))
        hasta    = _str(_col(fila, 'Hasta'))
        jornada  = _str(_col(fila, 'Jornada'))
        escalafon= _str(_col(fila, 'Escalafon'))
        unidad   = _str(_col(fila, 'Unidad'))
        macroun  = _str(_col(fila, 'MacroUn'))
        campus   = _str(_col(fila, 'Campus'))

        if decreto or glosa:
            personas[id_func]['actos'].append({
                'decreto':    decreto,
                'fecha':      fecha,
                'glosa':      glosa,
                'tipo_acto':  calidad,
                'desde':      desde,
                'hasta':      hasta,
                'observaciones': '',
                'soporte': '',
            })

        if unidad or jornada or escalafon:
            hist = {
                'cargo': calidad,
                'unidad': unidad,
                'departamento': macroun,
                'tipo_contrato': calidad,
                'estamento': '',
                'grado': escalafon,
                'jornada': jornada,
                'desde': desde,
                'hasta': hasta,
                'observaciones': campus,
            }
            existe = any(
                h['unidad'] == hist['unidad'] and h['desde'] == hist['desde']
                for h in personas[id_func]['historial']
            )
            if not existe:
                personas[id_func]['historial'].append(hist)

    return list(personas.values())


@app.route('/fichas/importar', methods=['GET'])
@login_required
def fichas_importar_get():
    return render_template('importar.html', personas=None, error=None)


@app.route('/fichas/importar', methods=['POST'])
@login_required
def fichas_importar_post():
    archivo = request.files.get('excel')
    if not archivo or not archivo.filename:
        return render_template('importar.html', personas=None, error='Selecciona un archivo Excel (.xlsx).')
    if not archivo.filename.lower().endswith(('.xlsx', '.xls')):
        return render_template('importar.html', personas=None, error='Solo se aceptan archivos .xlsx o .xls.')
    try:
        personas = _leer_excel_funcionarios(archivo.stream)
        if not personas:
            return render_template('importar.html', personas=None, error='El archivo no contiene datos o no tiene el formato esperado.')
        return render_template('importar.html', personas=personas, error=None)
    except Exception as e:
        return render_template('importar.html', personas=None, error=f'No se pudo leer el archivo: {e}')


@app.route('/fichas/confirmar-importacion', methods=['POST'])
@login_required
def fichas_confirmar_importacion():
    cantidad = int(request.form.get('cantidad_personas', 0))
    db = get_db()
    creadas = 0
    omitidas = []

    for i in range(cantidad):
        rut       = request.form.get(f'rut_{i}', '').strip()
        nombre    = request.form.get(f'nombre_{i}', '').strip()
        f_nac     = request.form.get(f'fecha_nacimiento_{i}', '').strip()
        est_civil = request.form.get(f'estado_civil_{i}', '').strip()
        titulo    = request.form.get(f'titulo_{i}', '').strip()
        direccion = request.form.get(f'direccion_{i}', '').strip()
        telefono  = request.form.get(f'telefono_{i}', '').strip()
        link      = request.form.get(f'link_{i}', '').strip()

        if not nombre:
            continue
        if not rut or not rut_valido(rut):
            omitidas.append(nombre + ' (RUT inválido o vacío)')
            continue
        if db.execute('SELECT id FROM fichas WHERE rut = ?', (rut,)).fetchone():
            omitidas.append(nombre + ' (RUT ya existe)')
            continue

        cursor = db.execute('''
            INSERT INTO fichas (nombre, rut, titulo, direccion, fecha_nacimiento, estado_civil, telefono, link, activo)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 1)
        ''', (nombre, rut, titulo, direccion, f_nac, est_civil, telefono, link))
        ficha_id = cursor.lastrowid
        _registrar_auditoria('crear', 'ficha', ficha_id, f'Importada desde Excel: {nombre} ({rut})')

        # Actos
        n_actos = int(request.form.get(f'n_actos_{i}', 0))
        for j in range(n_actos):
            db.execute('''
                INSERT INTO actos_administrativos (ficha_id, decreto, fecha, glosa, tipo_acto, desde, hasta, observaciones, soporte)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                ficha_id,
                request.form.get(f'acto_{i}_{j}_decreto', ''),
                request.form.get(f'acto_{i}_{j}_fecha', ''),
                request.form.get(f'acto_{i}_{j}_glosa', ''),
                request.form.get(f'acto_{i}_{j}_tipo_acto', ''),
                request.form.get(f'acto_{i}_{j}_desde', ''),
                request.form.get(f'acto_{i}_{j}_hasta', ''),
                request.form.get(f'acto_{i}_{j}_observaciones', ''),
                request.form.get(f'acto_{i}_{j}_soporte', ''),
            ))

        # Historial
        n_hist = int(request.form.get(f'n_hist_{i}', 0))
        for k in range(n_hist):
            db.execute('''
                INSERT INTO historial_institucional (ficha_id, cargo, unidad, departamento, tipo_contrato, grado, jornada, desde, hasta, observaciones)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                ficha_id,
                request.form.get(f'hist_{i}_{k}_cargo', ''),
                request.form.get(f'hist_{i}_{k}_unidad', ''),
                request.form.get(f'hist_{i}_{k}_departamento', ''),
                request.form.get(f'hist_{i}_{k}_tipo_contrato', ''),
                request.form.get(f'hist_{i}_{k}_grado', ''),
                request.form.get(f'hist_{i}_{k}_jornada', ''),
                request.form.get(f'hist_{i}_{k}_desde', ''),
                request.form.get(f'hist_{i}_{k}_hasta', ''),
                request.form.get(f'hist_{i}_{k}_observaciones', ''),
            ))

        db.commit()
        creadas += 1

    if omitidas:
        flash(f'{creadas} ficha(s) importada(s). Omitidas: {"; ".join(omitidas)}.', 'error')
    else:
        flash(f'{creadas} ficha(s) importada(s) correctamente desde Excel.', 'success')
    return redirect(url_for('fichas_listado'))

@login_required
def fichas_exportar():
    db = get_db()
    fichas = db.execute('SELECT * FROM fichas WHERE activo = 1 ORDER BY nombre COLLATE NOCASE').fetchall()

    wb = Workbook()
    hoja = wb.active
    hoja.title = 'Fichas de Personal'

    encabezados = ['Nombre', 'RUT', 'Título', 'Grado Académico', 'Jerarquía', 'Dirección', 'Teléfono',
                   'Fecha Nacimiento', 'Estado Civil', 'Sexo', 'Link', 'Observaciones']
    hoja.append(encabezados)
    for celda in hoja[1]:
        celda.font = Font(name='Arial', bold=True, color='FFFFFF')
        celda.fill = PatternFill('solid', start_color='1C3D5A')
        celda.alignment = Alignment(horizontal='left')

    for f in fichas:
        hoja.append([
            f['nombre'], f['rut'], f['titulo'], f['grado_academico'], f['jerarquia'], f['direccion'],
            f['telefono'], f['fecha_nacimiento'], f['estado_civil'], f['sexo'], f['link'], f['observaciones']
        ])

    for fila in hoja.iter_rows(min_row=2):
        for celda in fila:
            celda.font = Font(name='Arial')

    anchos = [24, 14, 26, 26, 18, 28, 16, 14, 14, 12, 28, 30]
    for columna, ancho in zip('ABCDEFGHIJKL', anchos):
        hoja.column_dimensions[columna].width = ancho

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return send_file(
        buffer, as_attachment=True, download_name='fichas_de_personal.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


@app.route('/fichas/papelera')
@login_required
def fichas_papelera():
    db = get_db()
    fichas = db.execute('SELECT * FROM fichas WHERE activo = 0 ORDER BY nombre COLLATE NOCASE').fetchall()
    return render_template('papelera.html', fichas=fichas)


@app.route('/fichas/<int:ficha_id>/restaurar', methods=['POST'])
@login_required
def ficha_restaurar(ficha_id):
    db = get_db()
    db.execute('UPDATE fichas SET activo = 1 WHERE id = ?', (ficha_id,))
    db.commit()
    _registrar_auditoria('restaurar', 'ficha', ficha_id, 'Ficha restaurada desde la papelera')
    flash('Ficha restaurada.', 'success')
    return redirect(url_for('fichas_papelera'))


@app.route('/fichas/<int:ficha_id>/eliminar_definitivo', methods=['POST'])
@login_required
def ficha_eliminar_definitivo(ficha_id):
    db = get_db()
    ficha = db.execute('SELECT nombre, rut FROM fichas WHERE id = ?', (ficha_id,)).fetchone()
    db.execute('DELETE FROM fichas WHERE id = ?', (ficha_id,))
    db.commit()
    if ficha:
        _registrar_auditoria(
            'eliminar_definitivo', 'ficha', ficha_id,
            f"Ficha eliminada definitivamente: {ficha['nombre']} ({ficha['rut']})"
        )
    flash('Ficha eliminada definitivamente, junto con sus actos y asignaciones familiares.', 'success')
    return redirect(url_for('fichas_papelera'))


# ---------- Crear / ver / editar una ficha ----------

@app.route('/fichas/nueva', methods=['GET', 'POST'])
@login_required
def ficha_nueva():
    if request.method == 'POST':
        datos = _leer_datos_ficha()
        if not datos['nombre'] or not datos['rut']:
            flash('Nombre y RUT son obligatorios.', 'error')
            return redirect(url_for('ficha_nueva'))

        if not rut_valido(datos['rut']):
            flash('El RUT ingresado no es válido (revisa el dígito verificador).', 'error')
            return redirect(url_for('ficha_nueva'))

        db = get_db()
        existente = db.execute('SELECT id FROM fichas WHERE rut = ?', (datos['rut'],)).fetchone()
        if existente:
            flash('Ya existe una ficha con ese RUT.', 'error')
            return redirect(url_for('ficha_nueva'))

        foto = guardar_archivo('foto', EXT_IMAGEN)

        cursor = db.execute('''
            INSERT INTO fichas (nombre, rut, titulo, grado_academico, jerarquia, direccion,
                                 observaciones, fecha_nacimiento, estado_civil, telefono, sexo,
                                 link, foto)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            datos['nombre'], datos['rut'], datos['titulo'], datos['grado_academico'], datos['jerarquia'],
            datos['direccion'], datos['observaciones'], datos['fecha_nacimiento'], datos['estado_civil'],
            datos['telefono'], datos['sexo'], datos['link'], foto
        ))
        db.commit()
        _registrar_auditoria('crear', 'ficha', cursor.lastrowid, f"Ficha creada: {datos['nombre']} ({datos['rut']})")
        flash('Ficha creada correctamente.', 'success')
        return redirect(url_for('ficha_detalle', ficha_id=cursor.lastrowid))

    return render_template('ficha.html', ficha=None, actos=[], familiares=[], tipos_acto=[], historial=[],
                           cat_contrato=[], cat_estamento=[], cat_jornada=[])


@app.route('/fichas/<int:ficha_id>')
@login_required
def ficha_detalle(ficha_id):
    db = get_db()
    ficha = db.execute('SELECT * FROM fichas WHERE id = ?', (ficha_id,)).fetchone()
    if ficha is None:
        flash('Esa ficha no existe.', 'error')
        return redirect(url_for('fichas_listado'))

    actos = []
    for fila in db.execute('SELECT * FROM actos_administrativos WHERE ficha_id = ? ORDER BY id DESC', (ficha_id,)):
        acto = dict(fila)
        archivos = db.execute(
            'SELECT * FROM archivos_actos WHERE acto_id = ? ORDER BY id', (acto['id'],)
        ).fetchall()
        acto['archivos'] = archivos
        acto['archivos_json'] = json.dumps([
            {
                'id': a['id'],
                'nombre': a['nombre_original'] or a['nombre_archivo'],
                'url': url_for('static', filename='uploads/' + a['nombre_archivo'])
            }
            for a in archivos
        ])
        actos.append(acto)

    familiares = []
    for fila in db.execute('SELECT * FROM asignaciones_familiares WHERE ficha_id = ? ORDER BY id DESC', (ficha_id,)):
        familiar = dict(fila)
        archivos = db.execute(
            'SELECT * FROM archivos_familiares WHERE familiar_id = ? ORDER BY id', (familiar['id'],)
        ).fetchall()
        familiar['archivos'] = archivos
        familiar['archivos_json'] = json.dumps([
            {
                'id': a['id'],
                'nombre': a['nombre_original'] or a['nombre_archivo'],
                'url': url_for('static', filename='uploads/' + a['nombre_archivo'])
            }
            for a in archivos
        ])
        familiares.append(familiar)

    tipos_acto = [{'id': fila['id'], 'nombre': fila['nombre']} for fila in db.execute('SELECT id, nombre FROM tipos_acto_administrativo ORDER BY nombre COLLATE NOCASE')]

    def _cat(categoria):
        return [{'id': f['id'], 'nombre': f['nombre']} for f in db.execute(
            'SELECT id, nombre FROM catalogos WHERE categoria=? ORDER BY nombre COLLATE NOCASE', (categoria,)
        )]

    cat_contrato  = _cat('contrato')
    cat_estamento = _cat('estamento')
    cat_jornada   = _cat('jornada')

    historial = []
    for fila in db.execute('SELECT * FROM historial_institucional WHERE ficha_id = ? ORDER BY desde DESC, id DESC', (ficha_id,)):
        registro = dict(fila)
        archivos = db.execute(
            'SELECT * FROM archivos_historial WHERE historial_id = ? ORDER BY id', (registro['id'],)
        ).fetchall()
        registro['archivos'] = archivos
        registro['archivos_json'] = json.dumps([
            {
                'id': a['id'],
                'nombre': a['nombre_original'] or a['nombre_archivo'],
                'url': url_for('static', filename='uploads/' + a['nombre_archivo'])
            }
            for a in archivos
        ])
        historial.append(registro)

    return render_template('ficha.html', ficha=ficha, actos=actos, familiares=familiares,
                           tipos_acto=tipos_acto, historial=historial,
                           cat_contrato=cat_contrato, cat_estamento=cat_estamento, cat_jornada=cat_jornada)


@app.route('/fichas/<int:ficha_id>/actualizar', methods=['POST'])
@login_required
def ficha_actualizar(ficha_id):
    db = get_db()
    actual = db.execute('SELECT * FROM fichas WHERE id = ?', (ficha_id,)).fetchone()
    if actual is None:
        flash('Esa ficha no existe.', 'error')
        return redirect(url_for('fichas_listado'))

    datos = _leer_datos_ficha()
    if not datos['nombre'] or not datos['rut']:
        flash('Nombre y RUT son obligatorios.', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    if not rut_valido(datos['rut']):
        flash('El RUT ingresado no es válido (revisa el dígito verificador).', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    conflicto = db.execute(
        'SELECT id FROM fichas WHERE rut = ? AND id != ?', (datos['rut'], ficha_id)
    ).fetchone()
    if conflicto:
        flash('Ya existe otra ficha con ese RUT.', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    foto = guardar_archivo('foto', EXT_IMAGEN) or actual['foto']

    db.execute('''
        UPDATE fichas SET nombre=?, rut=?, titulo=?, grado_academico=?, jerarquia=?, direccion=?,
                           observaciones=?, fecha_nacimiento=?, estado_civil=?, telefono=?, sexo=?,
                           link=?, foto=?
        WHERE id=?
    ''', (
        datos['nombre'], datos['rut'], datos['titulo'], datos['grado_academico'], datos['jerarquia'],
        datos['direccion'], datos['observaciones'], datos['fecha_nacimiento'], datos['estado_civil'],
        datos['telefono'], datos['sexo'], datos['link'], foto,
        ficha_id
    ))
    db.commit()
    _registrar_auditoria('actualizar', 'ficha', ficha_id, f"Ficha actualizada: {datos['nombre']} ({datos['rut']})")
    flash('Ficha actualizada correctamente.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


@app.route('/fichas/<int:ficha_id>/eliminar', methods=['POST'])
@login_required
def ficha_eliminar(ficha_id):
    db = get_db()
    ficha = db.execute('SELECT nombre, rut FROM fichas WHERE id = ?', (ficha_id,)).fetchone()
    db.execute('UPDATE fichas SET activo = 0 WHERE id = ?', (ficha_id,))
    db.commit()
    if ficha:
        _registrar_auditoria(
            'eliminar', 'ficha', ficha_id, f"Ficha movida a la papelera: {ficha['nombre']} ({ficha['rut']})"
        )
    flash('Ficha movida a la papelera. Sus actos y asignaciones familiares se mantienen y se pueden recuperar.', 'success')
    return redirect(url_for('fichas_listado'))


def _leer_datos_ficha():
    return {
        'nombre': request.form.get('nombre', '').strip(),
        'rut': request.form.get('rut', '').strip(),
        'titulo': request.form.get('titulo', '').strip(),
        'grado_academico': request.form.get('grado_academico', '').strip(),
        'jerarquia': request.form.get('jerarquia', '').strip(),
        'direccion': request.form.get('direccion', '').strip(),
        'observaciones': request.form.get('observaciones', '').strip(),
        'fecha_nacimiento': request.form.get('fecha_nacimiento', '').strip(),
        'estado_civil': request.form.get('estado_civil', '').strip(),
        'telefono': request.form.get('telefono', '').strip(),
        'sexo': request.form.get('sexo', '').strip(),
        'link': request.form.get('link', '').strip(),
    }


# ---------- Actos administrativos (anidados dentro de una ficha) ----------

@app.route('/fichas/<int:ficha_id>/actos/guardar', methods=['POST'])
@login_required
def actos_guardar(ficha_id):
    datos = _leer_datos_acto()
    if not datos['decreto']:
        flash('El campo "Decreto o Resolución" es obligatorio.', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    if not fechas_validas(datos['desde'], datos['hasta']):
        flash('La fecha "Hasta" no puede ser anterior a la fecha "Desde".', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    db = get_db()
    ficha = db.execute('SELECT id FROM fichas WHERE id = ?', (ficha_id,)).fetchone()
    if ficha is None:
        flash('Esa ficha no existe.', 'error')
        return redirect(url_for('fichas_listado'))

    cursor = db.execute('''
        INSERT INTO actos_administrativos (ficha_id, decreto, fecha, glosa, tipo_acto, desde, hasta, observaciones, soporte)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (ficha_id, datos['decreto'], datos['fecha'], datos['glosa'], datos['tipo_acto'],
          datos['desde'], datos['hasta'], datos['observaciones'], datos['soporte']))
    acto_id = cursor.lastrowid

    for nombre_guardado, nombre_original in guardar_archivos_multiples('archivos', EXT_DOCUMENTO):
        db.execute(
            'INSERT INTO archivos_actos (acto_id, nombre_archivo, nombre_original) VALUES (?, ?, ?)',
            (acto_id, nombre_guardado, nombre_original)
        )

    db.commit()
    _registrar_auditoria('crear', 'acto_administrativo', acto_id, f"Acto agregado: {datos['decreto']}")
    flash('Acto administrativo guardado.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


@app.route('/fichas/<int:ficha_id>/actos/modificar', methods=['POST'])
@login_required
def actos_modificar(ficha_id):
    acto_id = request.form.get('acto_id', '').strip()
    if not acto_id:
        flash('Selecciona un acto de la lista antes de modificarlo.', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    datos = _leer_datos_acto()
    if not fechas_validas(datos['desde'], datos['hasta']):
        flash('La fecha "Hasta" no puede ser anterior a la fecha "Desde".', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    db = get_db()
    db.execute('''
        UPDATE actos_administrativos
        SET decreto=?, fecha=?, glosa=?, tipo_acto=?, desde=?, hasta=?, observaciones=?, soporte=?
        WHERE id=? AND ficha_id=?
    ''', (datos['decreto'], datos['fecha'], datos['glosa'], datos['tipo_acto'],
          datos['desde'], datos['hasta'], datos['observaciones'], datos['soporte'], acto_id, ficha_id))

    for nombre_guardado, nombre_original in guardar_archivos_multiples('archivos', EXT_DOCUMENTO):
        db.execute(
            'INSERT INTO archivos_actos (acto_id, nombre_archivo, nombre_original) VALUES (?, ?, ?)',
            (acto_id, nombre_guardado, nombre_original)
        )

    db.commit()
    _registrar_auditoria('actualizar', 'acto_administrativo', acto_id, f"Acto actualizado: {datos['decreto']}")
    flash('Acto administrativo actualizado.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


@app.route('/fichas/<int:ficha_id>/actos/eliminar', methods=['POST'])
@login_required
def actos_eliminar(ficha_id):
    acto_id = request.form.get('acto_id', '').strip()
    if acto_id:
        db = get_db()
        for fila in db.execute('SELECT nombre_archivo FROM archivos_actos WHERE acto_id = ?', (acto_id,)):
            _borrar_archivo_fisico(fila['nombre_archivo'])
        db.execute('DELETE FROM actos_administrativos WHERE id = ? AND ficha_id = ?', (acto_id, ficha_id))
        db.commit()
        _registrar_auditoria('eliminar', 'acto_administrativo', acto_id, 'Acto administrativo eliminado')
        flash('Acto administrativo eliminado.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


@app.route('/fichas/<int:ficha_id>/actos/<int:acto_id>/archivos/<int:archivo_id>/eliminar', methods=['POST'])
@login_required
def actos_archivo_eliminar(ficha_id, acto_id, archivo_id):
    db = get_db()
    archivo = db.execute(
        'SELECT * FROM archivos_actos WHERE id = ? AND acto_id = ?', (archivo_id, acto_id)
    ).fetchone()
    if archivo:
        _borrar_archivo_fisico(archivo['nombre_archivo'])
        db.execute('DELETE FROM archivos_actos WHERE id = ?', (archivo_id,))
        db.commit()
        _registrar_auditoria('eliminar', 'archivo_acto', archivo_id, 'Archivo adjunto eliminado de un acto administrativo')
        flash('Archivo eliminado.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


@app.route('/fichas/<int:ficha_id>/actos/pdf')
@login_required
def actos_pdf(ficha_id):
    db = get_db()
    ficha = db.execute('SELECT * FROM fichas WHERE id = ?', (ficha_id,)).fetchone()
    if ficha is None:
        flash('Esa ficha no existe.', 'error')
        return redirect(url_for('fichas_listado'))

    actos = db.execute(
        'SELECT * FROM actos_administrativos WHERE ficha_id = ? ORDER BY id', (ficha_id,)
    ).fetchall()

    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, topMargin=2 * cm, bottomMargin=2 * cm)
    estilos = getSampleStyleSheet()
    cuerpo = []

    cuerpo.append(Paragraph('Registro de Actos Administrativos', estilos['Title']))
    cuerpo.append(Paragraph(f"{ficha['nombre']} — RUT {ficha['rut']}", estilos['Heading2']))
    cuerpo.append(Spacer(1, 12))

    encabezados = ['Decreto', 'Fecha', 'Glosa', 'Tipo de Acto', 'Desde', 'Hasta', 'Observ.', 'Soporte']
    filas = [encabezados]
    for acto in actos:
        filas.append([
            acto['decreto'] or '', acto['fecha'] or '', acto['glosa'] or '',
            acto['tipo_acto'] or '', acto['desde'] or '', acto['hasta'] or '',
            acto['observaciones'] or '', acto['soporte'] or ''
        ])

    if len(filas) == 1:
        cuerpo.append(Paragraph('No hay registros para mostrar.', estilos['Normal']))
    else:
        tabla = Table(filas, repeatRows=1)
        tabla.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1c3d5a')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f4f5f7')]),
        ]))
        cuerpo.append(tabla)

    doc.build(cuerpo)
    buffer.seek(0)
    nombre_archivo = f"actos_administrativos_{ficha['rut']}.pdf"
    return send_file(buffer, as_attachment=True, download_name=nombre_archivo, mimetype='application/pdf')


@app.route('/tipos-acto/agregar', methods=['POST'])
@login_required
def tipos_acto_agregar():
    nombre = request.form.get('nombre', '').strip()
    if not nombre:
        return jsonify({'ok': False, 'error': 'Escribe un nombre para el nuevo tipo.'}), 400

    db = get_db()
    existente = db.execute(
        'SELECT id, nombre FROM tipos_acto_administrativo WHERE nombre = ? COLLATE NOCASE', (nombre,)
    ).fetchone()
    if existente:
        return jsonify({'ok': True, 'id': existente['id'], 'nombre': existente['nombre']})

    cursor = db.execute('INSERT INTO tipos_acto_administrativo (nombre) VALUES (?)', (nombre,))
    db.commit()
    _registrar_auditoria('crear', 'tipo_acto_administrativo', cursor.lastrowid, f'Tipo de acto agregado: {nombre}')
    return jsonify({'ok': True, 'id': cursor.lastrowid, 'nombre': nombre})


@app.route('/tipos-acto/listar')
@login_required
def tipos_acto_listar():
    db = get_db()
    tipos = db.execute('SELECT id, nombre FROM tipos_acto_administrativo ORDER BY nombre COLLATE NOCASE').fetchall()
    return jsonify([{'id': t['id'], 'nombre': t['nombre']} for t in tipos])


@app.route('/tipos-acto/<int:tipo_id>/eliminar', methods=['POST'])
@login_required
def tipos_acto_eliminar(tipo_id):
    db = get_db()
    tipo = db.execute('SELECT nombre FROM tipos_acto_administrativo WHERE id = ?', (tipo_id,)).fetchone()
    if not tipo:
        return jsonify({'ok': False, 'error': 'Tipo no encontrado.'}), 404
    db.execute('DELETE FROM tipos_acto_administrativo WHERE id = ?', (tipo_id,))
    db.commit()
    _registrar_auditoria('eliminar', 'tipo_acto_administrativo', tipo_id, f'Tipo de acto eliminado: {tipo["nombre"]}')
    return jsonify({'ok': True})


# ---------- Catálogos genéricos (contrato, estamento, jornada) ----------

CATEGORIAS_VALIDAS = {'contrato', 'estamento', 'jornada'}


@app.route('/catalogo/<categoria>/listar')
@login_required
def catalogo_listar(categoria):
    if categoria not in CATEGORIAS_VALIDAS:
        return jsonify({'ok': False, 'error': 'Categoría no válida.'}), 400
    db = get_db()
    items = db.execute(
        'SELECT id, nombre FROM catalogos WHERE categoria=? ORDER BY nombre COLLATE NOCASE', (categoria,)
    ).fetchall()
    return jsonify([{'id': i['id'], 'nombre': i['nombre']} for i in items])


@app.route('/catalogo/<categoria>/agregar', methods=['POST'])
@login_required
def catalogo_agregar(categoria):
    if categoria not in CATEGORIAS_VALIDAS:
        return jsonify({'ok': False, 'error': 'Categoría no válida.'}), 400
    nombre = request.form.get('nombre', '').strip()
    if not nombre:
        return jsonify({'ok': False, 'error': 'Escribe un nombre.'}), 400
    db = get_db()
    existente = db.execute(
        'SELECT id, nombre FROM catalogos WHERE categoria=? AND nombre=? COLLATE NOCASE', (categoria, nombre)
    ).fetchone()
    if existente:
        return jsonify({'ok': True, 'id': existente['id'], 'nombre': existente['nombre']})
    cursor = db.execute('INSERT INTO catalogos (categoria, nombre) VALUES (?, ?)', (categoria, nombre))
    db.commit()
    _registrar_auditoria('crear', f'catalogo_{categoria}', cursor.lastrowid, f'Valor agregado al catálogo {categoria}: {nombre}')
    return jsonify({'ok': True, 'id': cursor.lastrowid, 'nombre': nombre})


@app.route('/catalogo/<categoria>/<int:item_id>/eliminar', methods=['POST'])
@login_required
def catalogo_eliminar(categoria, item_id):
    if categoria not in CATEGORIAS_VALIDAS:
        return jsonify({'ok': False, 'error': 'Categoría no válida.'}), 400
    db = get_db()
    item = db.execute('SELECT nombre FROM catalogos WHERE id=? AND categoria=?', (item_id, categoria)).fetchone()
    if not item:
        return jsonify({'ok': False, 'error': 'Elemento no encontrado.'}), 404
    db.execute('DELETE FROM catalogos WHERE id=?', (item_id,))
    db.commit()
    _registrar_auditoria('eliminar', f'catalogo_{categoria}', item_id, f'Valor eliminado del catálogo {categoria}: {item["nombre"]}')
    return jsonify({'ok': True})


# ---------- Historial institucional ----------

@app.route('/fichas/<int:ficha_id>/historial/guardar', methods=['POST'])
@login_required
def historial_guardar(ficha_id):
    datos = _leer_datos_historial()
    if not datos['cargo']:
        flash('El campo "Cargo" es obligatorio.', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))
    if not fechas_validas(datos['desde'], datos['hasta']):
        flash('La fecha "Hasta" no puede ser anterior a la fecha "Desde".', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    db = get_db()
    cursor = db.execute("""
        INSERT INTO historial_institucional
            (ficha_id, cargo, unidad, departamento, tipo_contrato, estamento, grado, jornada, desde, hasta, observaciones)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (ficha_id, datos['cargo'], datos['unidad'], datos['departamento'], datos['tipo_contrato'],
          datos['estamento'], datos['grado'], datos['jornada'], datos['desde'], datos['hasta'], datos['observaciones']))
    historial_id = cursor.lastrowid

    for nombre_guardado, nombre_original in guardar_archivos_multiples('archivos', EXT_DOCUMENTO):
        db.execute(
            'INSERT INTO archivos_historial (historial_id, nombre_archivo, nombre_original) VALUES (?, ?, ?)',
            (historial_id, nombre_guardado, nombre_original)
        )

    db.commit()
    _registrar_auditoria('crear', 'historial_institucional', historial_id, f'Historial agregado: {datos["cargo"]}')
    flash('Registro de historial guardado.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


@app.route('/fichas/<int:ficha_id>/historial/modificar', methods=['POST'])
@login_required
def historial_modificar(ficha_id):
    historial_id = request.form.get('historial_id', '').strip()
    if not historial_id:
        flash('Selecciona un registro del historial antes de modificarlo.', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    datos = _leer_datos_historial()
    if not fechas_validas(datos['desde'], datos['hasta']):
        flash('La fecha "Hasta" no puede ser anterior a la fecha "Desde".', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    db = get_db()
    db.execute("""
        UPDATE historial_institucional
        SET cargo=?, unidad=?, departamento=?, tipo_contrato=?, estamento=?, grado=?,
            jornada=?, desde=?, hasta=?, observaciones=?
        WHERE id=? AND ficha_id=?
    """, (datos['cargo'], datos['unidad'], datos['departamento'], datos['tipo_contrato'],
          datos['estamento'], datos['grado'], datos['jornada'], datos['desde'], datos['hasta'],
          datos['observaciones'], historial_id, ficha_id))

    for nombre_guardado, nombre_original in guardar_archivos_multiples('archivos', EXT_DOCUMENTO):
        db.execute(
            'INSERT INTO archivos_historial (historial_id, nombre_archivo, nombre_original) VALUES (?, ?, ?)',
            (historial_id, nombre_guardado, nombre_original)
        )

    db.commit()
    _registrar_auditoria('actualizar', 'historial_institucional', historial_id, f'Historial actualizado: {datos["cargo"]}')
    flash('Registro de historial actualizado.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


@app.route('/fichas/<int:ficha_id>/historial/eliminar', methods=['POST'])
@login_required
def historial_eliminar(ficha_id):
    historial_id = request.form.get('historial_id', '').strip()
    if historial_id:
        db = get_db()
        for fila in db.execute('SELECT nombre_archivo FROM archivos_historial WHERE historial_id = ?', (historial_id,)):
            _borrar_archivo_fisico(fila['nombre_archivo'])
        db.execute('DELETE FROM historial_institucional WHERE id = ? AND ficha_id = ?', (historial_id, ficha_id))
        db.commit()
        _registrar_auditoria('eliminar', 'historial_institucional', historial_id, 'Registro de historial eliminado')
        flash('Registro de historial eliminado.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


@app.route('/fichas/<int:ficha_id>/historial/<int:historial_id>/archivos/<int:archivo_id>/eliminar', methods=['POST'])
@login_required
def historial_archivo_eliminar(ficha_id, historial_id, archivo_id):
    db = get_db()
    archivo = db.execute(
        'SELECT * FROM archivos_historial WHERE id = ? AND historial_id = ?', (archivo_id, historial_id)
    ).fetchone()
    if archivo:
        _borrar_archivo_fisico(archivo['nombre_archivo'])
        db.execute('DELETE FROM archivos_historial WHERE id = ?', (archivo_id,))
        db.commit()
        _registrar_auditoria('eliminar', 'archivo_historial', archivo_id, 'Archivo adjunto eliminado del historial')
        flash('Archivo eliminado.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


def _leer_datos_historial():
    return {
        'cargo':         request.form.get('cargo', '').strip(),
        'unidad':        request.form.get('unidad', '').strip(),
        'departamento':  request.form.get('departamento', '').strip(),
        'tipo_contrato': request.form.get('tipo_contrato', '').strip(),
        'estamento':     request.form.get('estamento', '').strip(),
        'grado':         request.form.get('grado', '').strip(),
        'jornada':       request.form.get('jornada', '').strip(),
        'desde':         request.form.get('desde', '').strip(),
        'hasta':         request.form.get('hasta', '').strip(),
        'observaciones': request.form.get('observaciones', '').strip(),
    }

def _leer_datos_acto():
    return {
        'decreto': request.form.get('decreto', '').strip(),
        'fecha': request.form.get('fecha', '').strip(),
        'glosa': request.form.get('glosa', '').strip(),
        'tipo_acto': request.form.get('tipo_acto', '').strip(),
        'desde': request.form.get('desde', '').strip(),
        'hasta': request.form.get('hasta', '').strip(),
        'observaciones': request.form.get('observaciones', '').strip(),
        'soporte': request.form.get('soporte', '').strip(),
    }


# ---------- Asignación familiar (anidada dentro de una ficha) ----------

@app.route('/fichas/<int:ficha_id>/familiares/guardar', methods=['POST'])
@login_required
def familiares_guardar(ficha_id):
    datos = _leer_datos_familiar()
    if not datos['nombre']:
        flash('El nombre del familiar es obligatorio.', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    if not fechas_validas(datos['desde'], datos['hasta']):
        flash('La fecha "Hasta" no puede ser anterior a la fecha "Desde".', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    db = get_db()
    ficha = db.execute('SELECT id FROM fichas WHERE id = ?', (ficha_id,)).fetchone()
    if ficha is None:
        flash('Esa ficha no existe.', 'error')
        return redirect(url_for('fichas_listado'))

    cursor = db.execute('''
        INSERT INTO asignaciones_familiares
            (ficha_id, nombre, parentesco, resolucion, fecha, desde, hasta, fecha_nacimiento, observaciones, soporte)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (ficha_id, datos['nombre'], datos['parentesco'], datos['resolucion'], datos['fecha'],
          datos['desde'], datos['hasta'], datos['fecha_nacimiento'], datos['observaciones'], datos['soporte']))
    familiar_id = cursor.lastrowid

    for nombre_guardado, nombre_original in guardar_archivos_multiples('archivos', EXT_DOCUMENTO):
        db.execute(
            'INSERT INTO archivos_familiares (familiar_id, nombre_archivo, nombre_original) VALUES (?, ?, ?)',
            (familiar_id, nombre_guardado, nombre_original)
        )

    db.commit()
    _registrar_auditoria('crear', 'asignacion_familiar', familiar_id, f"Asignación familiar agregada: {datos['nombre']}")
    flash('Asignación familiar guardada.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


@app.route('/fichas/<int:ficha_id>/familiares/modificar', methods=['POST'])
@login_required
def familiares_modificar(ficha_id):
    familiar_id = request.form.get('familiar_id', '').strip()
    if not familiar_id:
        flash('Selecciona un familiar de la lista antes de modificarlo.', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    datos = _leer_datos_familiar()
    if not fechas_validas(datos['desde'], datos['hasta']):
        flash('La fecha "Hasta" no puede ser anterior a la fecha "Desde".', 'error')
        return redirect(url_for('ficha_detalle', ficha_id=ficha_id))

    db = get_db()
    db.execute('''
        UPDATE asignaciones_familiares
        SET nombre=?, parentesco=?, resolucion=?, fecha=?, desde=?, hasta=?, fecha_nacimiento=?, observaciones=?, soporte=?
        WHERE id=? AND ficha_id=?
    ''', (datos['nombre'], datos['parentesco'], datos['resolucion'], datos['fecha'], datos['desde'],
          datos['hasta'], datos['fecha_nacimiento'], datos['observaciones'], datos['soporte'], familiar_id, ficha_id))

    for nombre_guardado, nombre_original in guardar_archivos_multiples('archivos', EXT_DOCUMENTO):
        db.execute(
            'INSERT INTO archivos_familiares (familiar_id, nombre_archivo, nombre_original) VALUES (?, ?, ?)',
            (familiar_id, nombre_guardado, nombre_original)
        )

    db.commit()
    _registrar_auditoria('actualizar', 'asignacion_familiar', familiar_id, f"Asignación familiar actualizada: {datos['nombre']}")
    flash('Asignación familiar actualizada.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


@app.route('/fichas/<int:ficha_id>/familiares/eliminar', methods=['POST'])
@login_required
def familiares_eliminar(ficha_id):
    familiar_id = request.form.get('familiar_id', '').strip()
    if familiar_id:
        db = get_db()
        for fila in db.execute('SELECT nombre_archivo FROM archivos_familiares WHERE familiar_id = ?', (familiar_id,)):
            _borrar_archivo_fisico(fila['nombre_archivo'])
        db.execute('DELETE FROM asignaciones_familiares WHERE id = ? AND ficha_id = ?', (familiar_id, ficha_id))
        db.commit()
        _registrar_auditoria('eliminar', 'asignacion_familiar', familiar_id, 'Asignación familiar eliminada')
        flash('Asignación familiar eliminada.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


@app.route('/fichas/<int:ficha_id>/familiares/<int:familiar_id>/archivos/<int:archivo_id>/eliminar', methods=['POST'])
@login_required
def familiares_archivo_eliminar(ficha_id, familiar_id, archivo_id):
    db = get_db()
    archivo = db.execute(
        'SELECT * FROM archivos_familiares WHERE id = ? AND familiar_id = ?', (archivo_id, familiar_id)
    ).fetchone()
    if archivo:
        _borrar_archivo_fisico(archivo['nombre_archivo'])
        db.execute('DELETE FROM archivos_familiares WHERE id = ?', (archivo_id,))
        db.commit()
        _registrar_auditoria('eliminar', 'archivo_familiar', archivo_id, 'Archivo adjunto eliminado de una asignación familiar')
        flash('Archivo eliminado.', 'success')
    return redirect(url_for('ficha_detalle', ficha_id=ficha_id))


def _leer_datos_familiar():
    return {
        'nombre': request.form.get('nombre_familiar', '').strip(),
        'parentesco': request.form.get('parentesco', '').strip(),
        'resolucion': request.form.get('resolucion', '').strip(),
        'fecha': request.form.get('fecha_familiar', '').strip(),
        'desde': request.form.get('desde_familiar', '').strip(),
        'hasta': request.form.get('hasta_familiar', '').strip(),
        'fecha_nacimiento': request.form.get('fecha_nacimiento_familiar', '').strip(),
        'observaciones': request.form.get('observaciones_familiar', '').strip(),
        'soporte': request.form.get('soporte_familiar', '').strip(),
    }
if __name__ == '__main__':
    init_db()
    app.run(debug=True)
