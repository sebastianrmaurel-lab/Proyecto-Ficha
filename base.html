{% extends "base.html" %}
{% block titulo %}Importar Hoja de Vida{% endblock %}
{% block contenido %}

<a class="boton boton-fantasma" href="{{ url_for('fichas_listado') }}" style="margin-bottom:16px">← Volver</a>

<div class="timbre-modulo">
  <div>
    <p class="titulo-principal">Importar Hoja de Vida</p>
    <p class="subtitulo">Carga el Excel con los datos de una o varias personas</p>
  </div>
</div>

{% if error %}
  <div class="mensaje error" style="margin-bottom:16px">{{ error }}</div>
{% endif %}

{% if not personas %}
  <div class="seccion-ficha" style="max-width:580px;">
    <p class="eyebrow">Formato del archivo</p>
    <p style="font-size:.85rem;color:var(--ink-soft);margin:0 0 14px;line-height:1.6">
      El Excel debe tener una fila de encabezados con estas columnas:<br>
      <span style="font-family:var(--font-mono);font-size:.78rem">
        Run (RUT) · ROL · Pasaporte · País Nacionalidad · Nombre Completo · Correo · Región · Comuna · Ciudad ·
        Fecha Nacimiento · Estado Civil · Jerarquía · Grado Académico · Título · Teléfono · Dirección · Observaciones
      </span><br><br>
      Cada fila es una persona. Se importan todas las que tengan Nombre o RUT.
    </p>
    <form method="post" enctype="multipart/form-data">
      <div class="campo" style="margin-bottom:14px">
        <label for="excel" style="font-size:.78rem;font-weight:600;color:var(--ink-soft)">Archivo Excel (.xlsx)</label>
        <input type="file" id="excel" name="excel" accept=".xlsx,.xls" required style="margin-top:6px">
      </div>
      <button class="boton boton-primario" type="submit">Leer archivo →</button>
    </form>
  </div>

{% else %}
  <div class="seccion-ficha" style="margin-bottom:16px">
    <p class="eyebrow">Previsualización</p>
    <p style="font-size:.85rem;color:var(--ink-soft);margin:0">
      Se encontraron <strong>{{ personas|length }}</strong> persona(s). Revisa los datos — el RUT ya viene del Excel.
      Si algún RUT está vacío o inválido esa persona se omitirá y te avisará.
    </p>
  </div>

  <form method="post" action="{{ url_for('fichas_confirmar_importacion') }}">
    <input type="hidden" name="cantidad_personas" value="{{ personas|length }}">

    {% for persona in personas %}
    {% set i = loop.index0 %}
    <div class="seccion-ficha" style="margin-bottom:14px">
      <p class="eyebrow" style="margin-bottom:14px">Persona {{ loop.index }} — {{ persona.nombre or '(sin nombre)' }}</p>
      <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px 18px">
        {% for key,label in [
          ('nombre','Nombre Completo'),('rut','Run (RUT)'),('correo','Correo'),
          ('fecha_nacimiento','Fecha Nacimiento'),('estado_civil','Estado Civil'),('sexo','Sexo'),
          ('titulo','Título'),('grado_academico','Grado Académico'),('jerarquia','Jerarquía'),
          ('telefono','Teléfono'),('region','Región'),('comuna','Comuna'),
          ('ciudad','Ciudad'),('pais_nacionalidad','País Nacionalidad'),('rol','ROL'),
          ('pasaporte','Pasaporte'),('direccion','Dirección'),('observaciones','Observaciones')
        ] %}
        <div class="campo" {% if key in ('direccion','observaciones') %}style="grid-column:1/-1"{% endif %}>
          <label style="font-size:.72rem;font-weight:600;color:var(--ink-soft)">
            {{ label }}{% if key == 'rut' %} <span style="color:var(--accent)">*</span>{% endif %}
          </label>
          <input type="text" name="{{ key }}_{{ i }}" value="{{ persona.get(key,'') }}"
                                  style="background:var(--surface-muted);border:1px solid var(--border);border-radius:7px;padding:7px 10px;font-size:.84rem;width:100%;margin-top:4px">
        </div>
        {% endfor %}
      </div>
    </div>
    {% endfor %}

    <div style="display:flex;gap:10px;padding:14px 0;border-top:1px solid var(--border);position:sticky;bottom:0;background:var(--bg)">
      <button class="boton boton-primario" type="submit">Importar {{ personas|length }} hoja(s) de vida</button>
      <a class="boton boton-fantasma" href="{{ url_for('fichas_importar_get') }}">← Cargar otro archivo</a>
      <a class="boton boton-fantasma" href="{{ url_for('fichas_listado') }}">Cancelar</a>
    </div>
  </form>
{% endif %}

{% endblock %}
