// Formatea un RUT chileno mientras el usuario escribe: 12.345.678-5
function formatearRut(valor) {
  let limpio = valor.replace(/[^0-9kK]/g, '').toUpperCase();
  if (limpio.length === 0) return '';

  let cuerpo = limpio.slice(0, -1);
  let dv = limpio.slice(-1);

  let cuerpoConPuntos = '';
  while (cuerpo.length > 3) {
    cuerpoConPuntos = '.' + cuerpo.slice(-3) + cuerpoConPuntos;
    cuerpo = cuerpo.slice(0, -3);
  }
  cuerpoConPuntos = cuerpo + cuerpoConPuntos;

  return cuerpoConPuntos + (cuerpoConPuntos ? '-' : '') + dv;
}

document.addEventListener('DOMContentLoaded', function () {
  document.querySelectorAll('.formato-rut').forEach(function (campo) {
    campo.addEventListener('input', function (evento) {
      const posicionCursorDesdeElFinal = evento.target.value.length - evento.target.selectionStart;
      evento.target.value = formatearRut(evento.target.value);
      const nuevaPosicion = evento.target.value.length - posicionCursorDesdeElFinal;
      evento.target.setSelectionRange(nuevaPosicion, nuevaPosicion);
    });
  });
});
