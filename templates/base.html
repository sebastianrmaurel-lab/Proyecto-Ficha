<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{% block titulo %}Sistema de Personal{% endblock %}</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="stylesheet" href="{{ url_for('static', filename='estilo.css') }}">
  <style>
    /* ── estilos específicos del nuevo formato ── */
    .campo-hv{display:flex;align-items:baseline;margin-bottom:9px}
    .campo-hv .etq{font-size:.73rem;font-weight:600;color:var(--ink-soft);min-width:160px;text-align:right;padding-right:12px;white-space:nowrap}
    .campo-hv .val{font-size:.87rem;border-bottom:1px solid var(--border);flex:1;padding-bottom:2px;min-height:19px;font-family:var(--font-mono)}
    .campo-hv .val.txt{font-family:var(--font-body)}
    .bloque-datos{background:var(--surface);border:1px solid var(--border);border-top:none;display:grid;grid-template-columns:1fr 170px}
    .col-datos{padding:18px 22px}
    .col-foto{border-left:1px solid var(--border);display:flex;flex-direction:column;align-items:center;justify-content:center;padding:14px;gap:8px}
    .caja-foto-hv{width:120px;height:120px;border:1px solid var(--border);border-radius:var(--radius);background:var(--surface-muted);display:flex;align-items:center;justify-content:center;flex-direction:column;gap:5px;color:var(--ink-soft);font-size:.7rem;overflow:hidden}
    .caja-foto-hv img{width:100%;height:100%;object-fit:cover}
    .fila-periodo{display:flex;gap:20px;padding:10px 22px;border:1px solid var(--border);border-top:none;background:var(--surface)}
    .campo-periodo{display:flex;align-items:center;gap:8px;font-size:.83rem}
    .campo-periodo label{color:var(--ink-soft);font-weight:600;font-size:.75rem;min-width:38px}
    .campo-periodo input{border:1px solid var(--border);border-radius:6px;padding:4px 9px;font-family:var(--font-mono);font-size:.8rem;background:var(--surface-muted)}
    details.seccion{margin-top:14px;border:1px solid var(--border);border-radius:var(--radius);background:var(--surface);overflow:hidden}
    details.seccion>summary{list-style:none;cursor:pointer;padding:9px 16px;background:var(--surface-muted);border-bottom:1px solid transparent;display:flex;align-items:center;justify-content:space-between;gap:10px;font-family:var(--font-mono);font-size:.7rem;font-weight:500;letter-spacing:.07em;text-transform:uppercase;color:var(--primary);user-select:none}
    details.seccion>summary::-webkit-details-marker{display:none}
    details.seccion>summary::before{content:"▶";font-size:.6rem;margin-right:6px;transition:transform .15s;display:inline-block}
    details[open].seccion>summary::before{transform:rotate(90deg)}
    details[open].seccion>summary{border-bottom-color:var(--border)}
    .sum-right{display:flex;align-items:center;gap:8px}
    .badge-num{background:var(--primary);color:#fff;font-size:.68rem;font-weight:700;padding:1px 7px;border-radius:999px;font-family:var(--font-mono)}
    .badge-num.cero{background:var(--surface-muted);color:var(--ink-soft)}
    .btn-sec{font-family:var(--font-body);font-size:.75rem;font-weight:600;background:var(--primary);color:#fff;border:none;border-radius:6px;padding:4px 11px;cursor:pointer}
    .btn-sec:hover{opacity:.88}
    .scroll-tabla{overflow-x:auto}
    table.tabla-hv{width:100%;border-collapse:collapse;font-size:.79rem}
    table.tabla-hv th{text-align:left;padding:7px 11px;background:var(--surface-muted);color:var(--ink-soft);font-weight:600;font-size:.68rem;text-transform:uppercase;letter-spacing:.04em;border-bottom:1px solid var(--border);white-space:nowrap}
    table.tabla-hv td{padding:8px 11px;border-bottom:1px solid var(--border);white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:200px}
    table.tabla-hv tr:last-child td{border-bottom:none}
    table.tabla-hv tr:not(.fila-vacia):hover td{background:var(--primary-soft);cursor:pointer}
    .fila-vacia td{text-align:center;color:var(--ink-soft);font-size:.77rem;padding:14px!important;font-style:italic}
    .barra-acciones{display:flex;gap:8px;padding:14px 0;border-top:1px solid var(--border);position:sticky;bottom:0;background:var(--bg);margin-top:20px;z-index:10}
    dialog.modal-doc{border:none;border-radius:var(--radius);padding:0;width:min(700px,96vw);max-height:90vh;overflow-y:auto}
    dialog.modal-doc::backdrop{background:rgba(15,20,30,.52)}
    .modal-head{display:flex;justify-content:space-between;align-items:center;padding:16px 22px;border-bottom:1px solid var(--border);position:sticky;top:0;background:var(--surface);z-index:1}
    .modal-titulo{font-family:var(--font-display);font-size:1.05rem;font-weight:700;margin:0}
    .modal-body{padding:18px 22px}
    .rejilla-modal{display:grid;grid-template-columns:1fr 1fr;gap:13px 18px}
    .campo-modal{display:flex;flex-direction:column;gap:5px}
    .campo-modal.ancho{grid-column:1/-1}
    .campo-modal label{font-size:.75rem;font-weight:600;color:var(--ink-soft)}
    .campo-modal input,.campo-modal select{background:var(--surface-muted);border:1px solid var(--border);border-radius:7px;padding:8px 10px;font-size:.87rem;font-family:var(--font-body);color:var(--ink);width:100%}
    .campo-modal input[type=file]{padding:6px 10px}
    .cerrar-modal{background:none;border:none;font-size:1.35rem;color:var(--ink-soft);cursor:pointer;padding:2px 6px;border-radius:5px}
    .cerrar-modal:hover{background:var(--surface-muted)}
    .modal-acciones{padding:14px 22px;border-top:1px solid var(--border);display:flex;gap:8px;flex-wrap:wrap}
    .lista-archivos-adj{display:flex;flex-direction:column;gap:6px;margin-top:8px}
    .chip-doc{display:flex;align-items:center;justify-content:space-between;gap:8px;background:var(--surface-muted);border:1px solid var(--border);border-radius:7px;padding:6px 10px;font-size:.76rem}
    .chip-doc a{color:var(--primary);font-weight:600;text-decoration:none}
    .chip-doc a.quitar{color:var(--accent)}
    .seccion-datos{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:18px 22px;margin-bottom:14px}
    @media(max-width:820px){.bloque-datos{grid-template-columns:1fr}.col-foto{border-left:none;border-top:1px solid var(--border);flex-direction:row;justify-content:flex-start}.rejilla-modal{grid-template-columns:1fr}}
  </style>
</head>
<body>
<div class="contenedor">
  {% if session.usuario_id %}
  <div class="barra-superior">
    <nav>
      <a class="enlace-nav {{ 'activo' if request.path=='/dashboard' }}" href="{{ url_for('dashboard') }}">Dashboard</a>
      <a class="enlace-nav {{ 'activo' if request.path.startswith('/fichas') }}" href="{{ url_for('fichas_listado') }}">Hojas de Vida</a>
    </nav>
    <div class="acciones-pagina">
      <span style="color:var(--ink-soft);font-size:.8rem;">{{ session.usuario_nombre }}</span>
      <a class="boton boton-fantasma" href="{{ url_for('logout') }}">Cerrar sesión</a>
    </div>
  </div>
  {% endif %}
  {% with msgs = get_flashed_messages(with_categories=true) %}
    {% if msgs %}<div class="mensajes">{% for cat,txt in msgs %}<div class="mensaje {{ 'error' if cat=='error' else 'success' }}">{{ txt }}</div>{% endfor %}</div>{% endif %}
  {% endwith %}
  {% block contenido %}{% endblock %}
</div>
<script src="{{ url_for('static', filename='rut.js') }}"></script>
</body>
</html>
