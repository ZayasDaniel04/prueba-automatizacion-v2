// ══════════════════════════════════════════════════════════════════════════════
// Generador de Documentos de Cobro
// Banco de México · GSSPIM (Gerencia de Supervisión de Sistemas de Pagos
//                            e Infraestructuras de Mercado)
//
// DESCRIPCIÓN GENERAL:
//   Esta aplicación web cliente (sin backend) automatiza la generación de dos
//   documentos para el cobro de sanciones a instituciones financieras:
//     1. Layout Excel (.xlsx) 
//     2. Memorando Word (.docx) 
//
// DEPENDENCIAS (librerías cargadas en el HTML):
//   - xlsx.full.min.js  → SheetJS: lectura del Excel de entrada (sanciones)
//   - exceljs.min.js    → ExcelJS: escritura del Excel de salida con estilos
//   - jszip.min.js      → JSZip: construcción del archivo .docx (ZIP + XML)
//
// FLUJO PRINCIPAL:
//   1. Usuario carga Sanciones_impuestas.xlsm
//   2. Se muestra tabla con registros para seleccionar
//   3. Al seleccionar, se autocompletan los campos del formulario
//   4. Usuario ingresa el importe y verifica los datos
//   5. Se generan y descargan los documentos
// ══════════════════════════════════════════════════════════════════════════════


// ── ESTADO GLOBAL ──────────────────────────────────────────────────────────────
// registros: almacena todos los objetos cargados del Excel de sanciones.
//            Cada objeto tiene: institucion, casfim, expediente, no_res,
//            f_escrito, f_recep, memo, f_memo, monto, nombre_corto.
let registros = [];

// selReg: guarda el objeto del registro actualmente seleccionado en la tabla.
//         Se usa para referencia pero todos los datos de generación se leen
//         directamente de los campos del formulario (editables).
let selReg   = null;


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 1 — UTILIDADES DE FORMATO
// Funciones de propósito general para transformar datos antes de mostrarlos
// o insertarlos en los documentos generados.
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Formatea un número como moneda mexicana (MXN).
 * Ejemplo: fmtMoney(1500.5) → "$1,500.50"
 * @param {number} n - Valor numérico a formatear
 * @returns {string} Cadena con formato $X,XXX.XX
 */
function fmtMoney(n){
  return '$'+n.toLocaleString('es-MX',{minimumFractionDigits:2,maximumFractionDigits:2});
}

/**
 * Convierte una fecha en formato ISO (YYYY-MM-DD) al formato de texto
 * usado en los documentos oficiales del Banco de México.
 * Ejemplo: fmtDate("2026-05-15") → "15 de mayo de 2026"
 * @param {string} d - Fecha en formato ISO YYYY-MM-DD
 * @returns {string} Fecha en texto español, o cadena vacía si no hay valor
 */
function fmtDate(d){
  if(!d) return '';
  const[y,m,day]=d.split('-');
  // Nombres de meses en español (índice 0 = enero)
  const M=['enero','febrero','marzo','abril','mayo','junio',
           'julio','agosto','septiembre','octubre','noviembre','diciembre'];
  return `${parseInt(day)} de ${M[parseInt(m)-1]} de ${y}`;
}

/**
 * Normaliza fechas provenientes de SheetJS a formato ISO (YYYY-MM-DD).
 * SheetJS puede devolver fechas de tres formas distintas según la configuración:
 *   - Objeto Date  → se serializa a ISO directamente
 *   - Número serial de Excel (días desde 1/1/1900) → se convierte a Date
 *   - String       → se intenta parsear con el constructor Date
 *
 * @param {Date|number|string} val - Valor de fecha proveniente de SheetJS
 * @returns {string} Fecha en formato ISO YYYY-MM-DD, o cadena vacía si es inválida
 */
function excelDateToISO(val){
  if(!val) return '';

  // Caso 1: SheetJS devolvió un objeto Date nativo
  if(val instanceof Date){
    return val.toISOString().split('T')[0];
  }

  // Caso 2: SheetJS devolvió el número serial de Excel.
  // El serial 25569 corresponde al 1 de enero de 1970 (epoch Unix).
  // Fórmula: (serial - 25569) * 86400 segundos * 1000 ms
  if(typeof val === 'number'){
    const d = new Date(Math.round((val - 25569)*86400*1000));
    return d.toISOString().split('T')[0];
  }

  // Caso 3: Es una cadena de texto; intentar parsear
  if(typeof val === 'string'){
    const d = new Date(val);
    if(!isNaN(d)) return d.toISOString().split('T')[0];
  }

  return ''; // No se pudo interpretar la fecha
}

/**
 * Convierte un importe numérico (pesos mexicanos) a su representación
 * en texto en español en MAYÚSCULAS, tal como se requiere en documentos
 * oficiales del Banco de México.
 *
 * Ejemplos:
 *   numToWords(1500.50)    → "MIL QUINIENTOS PESOS 50/100 M.N."
 *   numToWords(1000000)    → "UN MILLÓN PESOS 00/100 M.N."
 *   numToWords(1234567.89) → "UN MILLÓN DOSCIENTOS TREINTA Y CUATRO MIL QUINIENTOS SESENTA Y SIETE PESOS 89/100 M.N."
 *
 * Soporta rangos hasta billones (1,000,000,000,000).
 *
 * @param {number} num - Importe en pesos mexicanos
 * @returns {string} Importe en letras en español mayúsculas con sufijo "PESOS XX/100 M.N."
 */
function numToWords(num){
  // Tablas de referencia para la conversión
  const ones=['','UNO','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE',
              'DIEZ','ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISÉIS',
              'DIECISIETE','DIECIOCHO','DIECINUEVE'];
  const tens=['','DIEZ','VEINTE','TREINTA','CUARENTA','CINCUENTA',
              'SESENTA','SETENTA','OCHENTA','NOVENTA'];
  const huns=['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS',
              'SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS'];

  /**
   * Convierte un número de 0-999 a texto en español.
   * @param {number} n - Número entre 0 y 999
   * @returns {string}
   */
  function three(n){
    if(n===0) return '';
    if(n===100) return 'CIEN'; // Caso especial: CIEN (no CIENTO cuando es exacto)
    const h=Math.floor(n/100), rem=n%100;
    let r = h ? huns[h]+' ' : '';
    if(rem < 20 && rem > 0){
      r += ones[rem]; // Números del 1 al 19 van directamente en la tabla
    } else {
      const t=Math.floor(rem/10), o=rem%10;
      if(t) r += tens[t];
      if(o) r += (t ? ' Y ' : '') + ones[o]; // "TREINTA Y UNO", etc.
    }
    return r.trim();
  }

  /**
   * Convierte un número de 0-999,999 a texto en español.
   * Maneja miles: "MIL", "DOS MIL", "CIEN MIL", etc.
   * @param {number} n - Número entre 0 y 999,999
   * @returns {string}
   */
  function thousands(n){
    if(n===0) return '';
    const th=Math.floor(n/1000), rest=n%1000;
    let r='';
    if(th > 0){
      // "MIL" cuando es exactamente 1000 (no "UN MIL")
      r += (th===1 ? 'MIL' : three(th)+' MIL') + ' ';
    }
    if(rest > 0) r += three(rest);
    return r.trim();
  }

  // Separar parte entera y centavos
  const ip = Math.floor(num);
  const dp = Math.round((num - ip) * 100); // Centavos (0-99)

  let r = '';

  // Conversión según magnitud
  if(ip >= 1000000000000){
    // Billones (1,000,000,000,000)
    const b = Math.floor(ip/1000000000000);
    r += (b===1 ? 'UN BILLÓN' : thousands(b)+' BILLONES') + ' ';
    const rest = ip % 1000000000000;
    if(rest >= 1000000){
      const m = Math.floor(rest/1000000);
      r += (m===1 ? 'UN MILLÓN' : thousands(m)+' MILLONES') + ' ';
      const rest2 = rest % 1000000;
      if(rest2 > 0) r += thousands(rest2);
    } else if(rest > 0){
      r += thousands(rest);
    }
  } else if(ip >= 1000000){
    // Millones
    const m = Math.floor(ip/1000000);
    r += (m===1 ? 'UN MILLÓN' : thousands(m)+' MILLONES') + ' ';
    const rest = ip % 1000000;
    if(rest > 0) r += thousands(rest);
  } else {
    // Menos de un millón
    r = thousands(ip);
  }

  // Formato final: "... PESOS XX/100 M.N."
  // Los centavos se formatean con padding a 2 dígitos (ej: "05", "00")
  return r.trim() + ` PESOS ${String(dp).padStart(2,'0')}/100 M.N.`;
}


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 2 — CARGA DEL ARCHIVO EXCEL DE SANCIONES
// Maneja la lectura del archivo Sanciones_impuestas.xlsm usando SheetJS.
// El proceso es asíncrono para no bloquear la UI mientras se procesa el archivo.
// ══════════════════════════════════════════════════════════════════════════════

// Al hacer clic en el botón visible, se dispara el input file oculto
document.getElementById('btn-cargar').addEventListener('click', () => {
  document.getElementById('file-input').click();
});

/**
 * EVENT: Cambio en el input de archivo.
 * Lee el Excel seleccionado, extrae los registros y los muestra en la tabla.
 * Usa FileReader para leer el archivo como ArrayBuffer en el navegador.
 */
document.getElementById('file-input').addEventListener('change', function(e){
  const file = e.target.files[0];
  if(!file){ return; } // No hay archivo seleccionado

  // Mostrar estado de carga inmediatamente para retroalimentación visual
  const lbl = document.getElementById('lbl-archivo');
  const btn = document.getElementById('btn-cargar');
  lbl.textContent = '⏳ Cargando archivo...';
  lbl.style.color = '#c8a96e'; // Color dorado
  btn.disabled = true;         // Evitar doble clic durante la carga
  btn.style.opacity = '0.6';

  const reader = new FileReader();

  reader.onload = function(ev){
    // Se usa setTimeout con 50ms para permitir que el navegador actualice
    // la UI (spinner) ANTES de ejecutar el procesamiento pesado del Excel
    setTimeout(() => {
      try{
        // Convertir el resultado del FileReader a Uint8Array para SheetJS
        const data = new Uint8Array(ev.target.result);

        // Leer el Excel con opciones optimizadas para velocidad:
        // - cellDates:true  → convierte números seriales a objetos Date
        // - cellStyles:false → no procesar estilos visuales (no los necesitamos)
        // - cellNF:false     → no procesar formatos de número
        // - cellHTML:false   → no generar HTML para cada celda
        // - sheetRows:3000   → limitar a 3000 filas máximas para rendimiento
        const wb = XLSX.read(data, {
          type      : 'array',
          cellDates : true,
          cellStyles: false,
          cellNF    : false,
          cellHTML  : false,
          sheetRows : 3000,
        });

        // Acceder específicamente a la hoja "Sanciones Impuestas"
        // (el nombre debe coincidir exactamente, incluyendo tildes y mayúsculas)
        const ws = wb.Sheets['Sanciones Impuestas'];
        if(!ws){
          showAlert('err','No se encontró la hoja "Sanciones Impuestas".');
          resetBtn(btn, lbl, file.name, 0, true);
          return;
        }

        // Convertir la hoja a un array de objetos JSON.
        // range:1 indica que los headers están en la FILA 2 del Excel (índice 1).
        // defval:'' → campos vacíos retornan '' en lugar de undefined
        // raw:false → SheetJS convierte fechas a strings automáticamente
        const rows = XLSX.utils.sheet_to_json(ws, {
          range  : 1,
          defval : '',
          raw    : false,
        });

        // Mapear cada fila del Excel a un objeto con las propiedades que usa la app.
        // Solo se incluyen filas que tengan Institución Y No. Resolución (obligatorios).
        registros = rows
          .filter(r =>
            r['Institución'] && String(r['Institución']).trim() !== '' &&
            r['No. Resolución'] && String(r['No. Resolución']).trim() !== ''
          )
          .map(r => ({
            // Nombre completo de la institución sancionada
            institucion : String(r['Institución']||'').trim(),
            // Código CASFIM: se elimina ".0" que SheetJS agrega a números enteros
            casfim      : String(r['CASFIM']||'').replace(/\.0$/,'').trim(),
            // Número de expediente del procedimiento de sanción
            expediente  : String(r['Expediente']||'').trim(),
            // Número de resolución de la sanción (ej: S22-118-2026)
            no_res      : String(r['No. Resolución']||'').trim(),
            // Fecha en que la institución autorizó el cobro (escrito)
            f_escrito   : excelDateToISO(r['Fecha autorización de cobro']),
            // Fecha en que Banco de México recibió la autorización
            f_recep     : excelDateToISO(r['Fecha de recepción de autorización de cobro']),
            // Número del memorando de solicitud de cobro
            memo        : String(r['No. de Memorando']||'').trim(),
            // Fecha del memorando (solicitud de cobro)
            f_memo      : excelDateToISO(r['Fecha de solicitud de cobro']),
            // Monto original de la sanción (puede diferir del importe a cobrar)
            monto       : parseFloat(String(r['Monto impuesto de la sanción']).replace(/,/g,''))||0,
            // Nombre corto de la institución (para uso interno)
            nombre_corto: String(r['Nombre corto']||'').trim(),
          }));

        // Renderizar la tabla con los registros cargados y mostrar la sección
        mostrarTabla(registros);
        document.getElementById('sec-tabla').style.display='block';
        resetBtn(btn, lbl, file.name, registros.length, false);
        showAlert('ok', `✅ Se cargaron <strong>${registros.length}</strong> registros correctamente.`);

      } catch(err){
        // Error al parsear el Excel (archivo corrupto, formato incorrecto, etc.)
        showAlert('err','Error al leer el archivo: '+err.message);
        resetBtn(btn, lbl, file.name, 0, true);
      }
    }, 50); // Delay mínimo para que el spinner sea visible antes del procesamiento
  };

  // Leer el archivo como ArrayBuffer (requerido por SheetJS)
  reader.readAsArrayBuffer(file);

  // Resetear el input para permitir cargar el mismo archivo nuevamente si es necesario
  this.value='';
});

/**
 * Restaura el estado visual del botón de carga después de procesar el archivo.
 * Muestra el nombre del archivo y la cantidad de registros, o un mensaje de error.
 *
 * @param {HTMLButtonElement} btn   - El botón de cargar
 * @param {HTMLElement}       lbl   - El label de estado del archivo
 * @param {string}            filename - Nombre del archivo procesado
 * @param {number}            count    - Cantidad de registros cargados
 * @param {boolean}           error    - true si hubo un error durante la carga
 */
function resetBtn(btn, lbl, filename, count, error){
  btn.disabled = false;
  btn.style.opacity = '1';
  if(error){
    lbl.textContent = '❌ Error al cargar el archivo';
    lbl.style.color = '#b83232'; // Rojo de error
  } else {
    lbl.textContent = `✓ ${filename} — ${count} registros cargados`;
    lbl.style.color = '#2d7a4f'; // Verde de éxito
  }
}


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 3 — TABLA DE REGISTROS
// Renderiza la lista de sanciones cargadas y gestiona la selección de una fila.
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Renderiza un array de registros en la tabla HTML de selección.
 * Usa DocumentFragment para insertar todas las filas en un único repaint del DOM,
 * lo que mejora significativamente el rendimiento con muchos registros.
 *
 * @param {Array<Object>} datos - Array de objetos de registro a mostrar en la tabla
 */
function mostrarTabla(datos){
  const tbody = document.getElementById('tbl-body');
  tbody.innerHTML = ''; // Limpiar filas anteriores

  // DocumentFragment: acumula los nodos en memoria antes de insertarlos en el DOM.
  // Esto evita múltiples repaints del navegador (uno por fila) y es mucho más rápido.
  const frag = document.createDocumentFragment();

  datos.forEach((r,i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${r.institucion}</td>
      <td>${r.memo||'—'}</td>
      <td>${r.expediente}</td>
      <td>${r.no_res}</td>
    `;
    tr.style.cursor = 'pointer'; // Indicar que la fila es clickeable

    // Guardar el índice del registro como atributo data- para recuperarlo rápido en el click
    // (más eficiente que buscar por valor en el array)
    tr.dataset.idx = i;

    tr.addEventListener('click', function(){
      seleccionarRegistro(datos[this.dataset.idx]);
      resaltarFila(this);
    });

    frag.appendChild(tr);
  });

  tbody.appendChild(frag); // Un único repaint del DOM
}

/**
 * Resalta visualmente la fila seleccionada en la tabla.
 * Quita la clase 'sel' de todas las filas y la agrega solo a la recibida.
 *
 * @param {HTMLTableRowElement} tr - La fila a resaltar
 */
function resaltarFila(tr){
  // Quitar selección anterior (puede haber solo una fila activa a la vez)
  document.querySelectorAll('#tbl-body tr').forEach(t=>t.classList.remove('sel'));
  tr.classList.add('sel'); // Aplicar fondo azul claro (definido en styles.css)
}


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 4 — BÚSQUEDA EN TIEMPO REAL
// Filtra los registros de la tabla mientras el usuario escribe en el buscador.
// ══════════════════════════════════════════════════════════════════════════════

/**
 * EVENT: Input en el campo buscador.
 * Si el texto tiene 2 o más caracteres, filtra el array global 'registros'
 * buscando coincidencias en institución, no. resolución, expediente y memorando.
 * Si hay menos de 2 caracteres, muestra todos los registros.
 */
document.getElementById('buscador').addEventListener('input', function(){
  const q = this.value.trim().toLowerCase();

  const filtrados = q.length < 2
    ? registros // Mostrar todo si la búsqueda es muy corta
    : registros.filter(r =>
        r.institucion.toLowerCase().includes(q) ||
        r.no_res.toLowerCase().includes(q) ||
        r.expediente.toLowerCase().includes(q) ||
        r.memo.toLowerCase().includes(q)
      );

  mostrarTabla(filtrados); // Re-renderizar la tabla con los resultados filtrados
});


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 5 — SELECCIÓN Y AUTOCOMPLETADO DEL FORMULARIO
// Al hacer clic en una fila de la tabla, se autocompletan todos los campos
// del formulario con los datos del registro seleccionado.
// El importe queda vacío ya que puede diferir del monto original por
// actualizaciones de la sanción o descuentos por pronto pago.
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Autocompleta el formulario con los datos de un registro seleccionado.
 * Todos los campos son editables después del autocompletado.
 *
 * @param {Object} r - Objeto de registro con las propiedades del Excel
 */
function seleccionarRegistro(r){
  selReg = r; // Guardar referencia al registro seleccionado

  // Autocompletar todos los campos del formulario
  setVal('inst-nombre', r.institucion); // Nombre completo de la institución
  setVal('casfim',      r.casfim);      // Código CASFIM
  setVal('no-res',      r.no_res);      // No. Resolución
  setVal('exp',         r.expediente);  // Expediente
  setVal('importe',     '');            // Importe VACÍO (debe ingresarse manualmente)
  setVal('f-escrito',   r.f_escrito);   // Fecha autorización de cobro
  setVal('f-recep',     r.f_recep);     // Fecha recepción de autorización
  setVal('memo',        r.memo);        // No. Memorando
  setVal('f-memo',      r.f_memo);      // Fecha del memorando

  // Regenerar el campo TEXTO automáticamente con los nuevos valores
  // (resetear el flag de edición manual para que autoTxt() pueda escribirlo)
  const tf = document.getElementById('texto');
  tf.removeAttribute('data-manual'); // Quitar bloqueo de edición manual
  autoTxt(); // Generar: "No.Resolución/Expediente" (máx. 42 chars)

  // Mostrar las secciones que estaban ocultas
  document.getElementById('sec-form').style.display='block';
  document.getElementById('sec-preview').style.display='block';
  document.getElementById('sec-generar').style.display='block';

  // Scroll suave al inicio del formulario para que el usuario lo vea
  document.getElementById('sec-form').scrollIntoView({behavior:'smooth', block:'start'});

  // Actualizar contadores y vista previa con los nuevos datos
  updImporte();
  updCC();
  updPreview();
}

/**
 * Establece el valor de un campo del formulario de forma segura.
 * Si el elemento no existe, no hace nada (evita errores silenciosos).
 *
 * @param {string} id  - ID del elemento HTML
 * @param {*}      val - Valor a establecer (se convierte a string vacío si es null/undefined)
 */
function setVal(id, val){
  const el = document.getElementById(id);
  if(el) el.value = val || '';
}


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 6 — AUTO-TEXTO Y CONTADOR DE CARACTERES
// El campo TEXTO del layout BBVA se autogenera como "No.Resolución/Expediente"
// pero puede editarse manualmente. El contador muestra cuántos de los 42
// caracteres máximos se están usando, con alertas visuales por colores.
// ══════════════════════════════════════════════════════════════════════════════

// Cuando cambian No. Resolución o Expediente, regenerar el TEXTO automáticamente
document.getElementById('no-res').addEventListener('input', ()=>{ autoTxt(); updPreview(); });
document.getElementById('exp').addEventListener('input',    ()=>{ autoTxt(); updPreview(); });

/**
 * Autogenera el valor del campo TEXTO como "No.Resolución/Expediente"
 * truncado a 42 caracteres (límite del layout BBVA).
 * Solo se ejecuta si el usuario NO ha editado el campo manualmente
 * (verificado mediante el atributo data-manual).
 */
function autoTxt(){
  const r  = document.getElementById('no-res').value.trim();
  const e  = document.getElementById('exp').value.trim();
  const tf = document.getElementById('texto');

  // Si el usuario no ha editado manualmente el campo Y hay valores disponibles,
  // generar el texto automáticamente
  if(!tf.dataset.manual && r && e){
    tf.value = (r+'/'+e).slice(0,42); // Concatenar y truncar al límite
    updCC(); // Actualizar el contador de caracteres
  }
}

// Si el usuario edita el campo TEXTO manualmente, marcarlo para que
// autoTxt() no lo sobreescriba con cambios futuros en otros campos
document.getElementById('texto').addEventListener('input', function(){
  this.dataset.manual='1'; // Flag: el usuario tomó control del campo
  updCC();
  updPreview();
});

/**
 * Actualiza el contador de caracteres del campo TEXTO.
 * Cambia el color del contador según el uso:
 *   - Normal  (≤35): color gris muted
 *   - Aviso   (36-42): color dorado (clase 'warn')
 *   - Error   (>42): color rojo negrita (clase 'over')
 */
function updCC(){
  const v  = document.getElementById('texto').value;
  const el = document.getElementById('cc');
  el.textContent = v.length + ' / 42';
  // Aplicar clase según el nivel de uso
  el.className = 'cc' + (v.length>42 ? ' over' : v.length>35 ? ' warn' : '');
}


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 7 — ACTUALIZACIÓN DE IMPORTE Y LISTENERS GENERALES
// El importe tiene su propio listener para actualizar la preview.
// El resto de campos también están conectados para mantener la preview actualizada.
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Disparador de actualización de preview cuando cambia el importe.
 * Función separada para poder ser llamada desde seleccionarRegistro().
 */
function updImporte(){
  updPreview();
}

// Listener del campo importe
document.getElementById('importe').addEventListener('input', updImporte);

// Registrar listeners de input y change en todos los campos del formulario
// para mantener la vista previa actualizada en tiempo real
['inst-nombre','casfim','no-res','exp','memo','f-memo','f-escrito','f-recep','operacion'].forEach(id=>{
  const el = document.getElementById(id);
  if(el){
    el.addEventListener('input',  updPreview);
    el.addEventListener('change', updPreview);
  }
});


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 8 — VISTA PREVIA EN TIEMPO REAL
// Muestra un resumen de los datos que se usarán en los documentos,
// dividido en sección "Layout Excel (BBVA)" y sección "Memorando Word".
// Se actualiza automáticamente con cada cambio en el formulario.
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Actualiza el HTML de la caja de vista previa con los datos actuales del formulario.
 * Se llama automáticamente desde todos los listeners de cambio en el formulario.
 * Muestra placeholders ("—") cuando un campo está vacío.
 */
function updPreview(){
  // Leer los valores actuales del formulario para la preview
  const inst   = document.getElementById('inst-nombre').value || '[Institución]';
  const casfim = document.getElementById('casfim').value || '—';
  const raw    = parseFloat(document.getElementById('importe').value) || 0;
  const texto  = document.getElementById('texto').value || '—';
  const memo   = document.getElementById('memo').value || '—';
  const fm     = document.getElementById('f-memo').value;    // Fecha memorando (ISO)
  const fe     = document.getElementById('f-escrito').value; // Fecha escrito (ISO)

  // Generar HTML de la preview con dos secciones:
  // 1. Layout Excel (BBVA): datos que van en el archivo .xlsx
  // 2. Memorando Word: datos principales del documento .docx
  document.getElementById('preview').innerHTML=`
    <div style="margin-bottom:10px;padding-bottom:10px;border-bottom:1px solid #e8e4de">
      <strong style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#6b6b6b">Layout Excel (BBVA)</strong>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px 16px;font-size:13px;margin-bottom:14px">
      <div><span class="tag">Institución</span><br>${inst}</div>
      <div><span class="tag">CASFIM</span><br>${casfim}</div>
      <div><span class="tag">Importe</span><br>${raw>0 ? fmtMoney(raw) : '—'}</div>
      <div><span class="tag">Cuenta Resultados</span><br>385050311</div>
      <div><span class="tag">TEXTO</span><br>${texto}</div>
    </div>
    <div style="margin:10px 0;padding:10px 0;border-top:1px solid #e8e4de;border-bottom:1px solid #e8e4de">
      <strong style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#6b6b6b">Memorando Word</strong>
    </div>
    <div style="font-size:13px;line-height:2">
      <strong>Ref.:</strong> ${memo}<br>
      <strong>Fecha:</strong> ${fm ? fmtDate(fm) : '—'}<br>
      <strong>Institución:</strong> ${inst}<br>
      <strong>Monto:</strong> ${raw>0 ? fmtMoney(raw) : '—'}<br>
      <strong>Fecha escrito:</strong> ${fe ? fmtDate(fe) : '—'}
    </div>`;
}


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 9 — VALIDACIÓN Y ALERTAS
// Valida los campos antes de generar documentos.
// Los errores se muestran como alerta y bloquean la generación.
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Valida los campos del formulario según el tipo de documento a generar.
 * Para el Excel solo se validan los campos del layout.
 * Para el Word (word=true) se validan también las fechas del memorando.
 *
 * @param {boolean} word - Si es true, incluye validaciones adicionales para el .docx
 * @returns {string[]} Array de mensajes de error. Vacío si todo es válido.
 */
function validate(word=false){
  const e = []; // Acumulador de errores

  // Validaciones comunes (requeridas para Excel y Word)
  if(!document.getElementById('inst-nombre').value.trim())
    e.push('Nombre de institución es obligatorio');
  if(!document.getElementById('no-res').value.trim())
    e.push('No. Resolución es obligatorio');
  if(!document.getElementById('exp').value.trim())
    e.push('Expediente es obligatorio');
  if(!(parseFloat(document.getElementById('importe').value) > 0))
    e.push('Ingresa el importe');

  const t = document.getElementById('texto').value.trim();
  if(!t) e.push('TEXTO es obligatorio');
  if(t.length > 42) e.push('TEXTO excede 42 caracteres'); // Límite del layout BBVA

  // Validaciones adicionales solo para el memorando Word
  if(word){
    if(!document.getElementById('memo').value.trim())
      e.push('No. Memorando es obligatorio');
    if(!document.getElementById('f-memo').value)
      e.push('Fecha del memorando es obligatoria');
    if(!document.getElementById('f-escrito').value)
      e.push('Fecha del escrito es obligatoria');
    if(!document.getElementById('f-recep').value)
      e.push('Fecha de recepción es obligatoria');
  }

  return e;
}

/**
 * Muestra una alerta en la parte superior de la página.
 * La alerta se autodestruye después de 7 segundos.
 *
 * @param {string} type - Tipo de alerta: 'ok' (éxito) o 'err' (error)
 * @param {string} msg  - Mensaje a mostrar (puede contener HTML, ej: <strong>)
 */
function showAlert(type, msg){
  document.getElementById('alert-area').innerHTML =
    `<div class="alert al-${type}">${msg}</div>`;
  // Hacer scroll al top para que la alerta sea visible sin tener que hacer scroll
  window.scrollTo({top:0, behavior:'smooth'});
  // Remover la alerta después de 7 segundos
  setTimeout(()=>document.getElementById('alert-area').innerHTML='', 7000);
}

/**
 * Descarga un Blob como archivo con el nombre especificado.
 * Crea un enlace temporal, simula el clic y lo destruye.
 *
 * @param {Blob}   blob - Contenido del archivo a descargar
 * @param {string} name - Nombre del archivo con extensión (ej: "archivo.xlsx")
 */
function download(blob, name){
  const u = URL.createObjectURL(blob);       // Crear URL temporal para el blob
  const a = document.createElement('a');    // Crear enlace invisible
  a.href = u;
  a.download = name;
  a.click();                                 // Simular clic para iniciar descarga
  URL.revokeObjectURL(u);                    // Liberar la URL temporal de memoria
}


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 10 — GENERACIÓN DEL LAYOUT EXCEL (.xlsx)
// Genera el archivo de instrucción de cobro para BBVA usando la librería ExcelJS.
// ExcelJS se usa (en lugar de SheetJS) porque permite aplicar estilos completos:
// bordes, rellenos, fuentes, formatos numéricos y tamaños de fila/columna.
//
// Nombre del archivo: UL_LAYOUT_COBRO_{no_res}_{casfim}.xlsx
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Genera y descarga el layout de cobro en formato Excel (.xlsx).
 * Estructura del archivo:
 *   Filas 1-2: Espaciado vacío
 *   Fila 3:    Título fusionado B3:H3
 *   Fila 4:    Espaciado vacío
 *   Fila 5:    Encabezados de columnas (con fondo gris)
 *   Fila 6:    Datos del cobro
 *   Fila 7:    Espaciado vacío
 *   Filas 8-11: Notas informativas
 */
async function generateExcel(){
  // Validar antes de generar (solo validaciones del Excel, no del Word)
  const errs = validate(false);
  if(errs.length){ showAlert('err','⚠ '+errs.join(' · ')); return; }

  // Leer valores del formulario
  const inst      = document.getElementById('inst-nombre').value.trim();
  const casfim    = document.getElementById('casfim').value.trim();
  const importe   = parseFloat(document.getElementById('importe').value);
  const operacion = document.getElementById('operacion').value[0]; // Solo primer carácter: 'S' o 'D'
  const texto     = document.getElementById('texto').value.trim();
  const res       = document.getElementById('no-res').value.trim();
  const CUENTA    = '385050311'; // Cuenta fija: Multas Art. 36 Bis L.B.M. – DGSPIM

  // Crear libro y hoja de trabajo con ExcelJS
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('LAYOUT', {
    properties: { tabColor: {argb:'FFFFFFFF'} } // Pestaña blanca
  });

  // Ocultar líneas de cuadrícula para un look más limpio y profesional
  ws.views = [{showGridLines: false}];

  // Definir anchos de columna en caracteres (1 = aprox. 7 píxeles)
  ws.columns = [
    {width: 6},   // A: columna vacía izquierda (margen)
    {width: 12},  // B: Número de sanción
    {width: 11},  // C: *Operación (S/D)
    {width: 9},   // D: CASFIM
    {width: 88},  // E: Nombre Institución (ancha para nombres largos)
    {width: 26},  // F: IMPORTE
    {width: 26},  // G: CUENTA DE RESULTADOS
    {width: 36},  // H: * TEXTO
  ];

  // ── Alturas de filas ──
  // Las filas 1, 2, 4 y 7 son vacías de espaciado pero necesitan contenido
  // (espacio en blanco) para que ExcelJS respete su altura definida
  [1, 2, 4, 7].forEach(n => {
    const row  = ws.getRow(n);
    row.height = 15;  // Altura en puntos
    // Insertar espacio en A para forzar la altura (sin esto ExcelJS puede comprimir la fila)
    const cell = ws.getCell(`A${n}`);
    cell.value = ' ';
    cell.font  = {name:'Calibri', size:11};
  });
  ws.getRow(3).height  = 22;  // Fila del título
  ws.getRow(5).height  = 42;  // Fila de encabezados (más alta por el texto con saltos de línea)
  ws.getRow(6).height  = 16;  // Fila de datos
  ws.getRow(8).height  = 14;  // Filas de notas
  ws.getRow(9).height  = 14;
  ws.getRow(10).height = 14;
  ws.getRow(11).height = 14;

  // ── Definición de estilos reutilizables ──

  // Borde grueso (para el título)
  const borderMedium = {
    top   :{style:'medium',color:{argb:'FF000000'}},
    bottom:{style:'medium',color:{argb:'FF000000'}},
    left  :{style:'medium',color:{argb:'FF000000'}},
    right :{style:'medium',color:{argb:'FF000000'}},
  };
  // Borde fino (para encabezados y datos)
  const borderThin = {
    top   :{style:'thin',color:{argb:'FF000000'}},
    bottom:{style:'thin',color:{argb:'FF000000'}},
    left  :{style:'thin',color:{argb:'FF000000'}},
    right :{style:'thin',color:{argb:'FF000000'}},
  };

  const fillGray  = {type:'pattern',pattern:'solid',fgColor:{argb:'FFD9D9D9'}}; // Gris claro para encabezados
  const fillWhite = {type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFFFF'}}; // Blanco para datos

  const alignCC  = {horizontal:'center', vertical:'middle', wrapText:false};  // Centrado sin wrap
  const alignCCW = {horizontal:'center', vertical:'middle', wrapText:true};   // Centrado con wrap (headers)
  const alignLC  = {horizontal:'left',   vertical:'middle', wrapText:false};  // Izquierda (notas)

  // ── Fila 3: Título principal (celdas B3:H3 fusionadas) ──
  ws.mergeCells('B3:H3'); // Fusionar columnas B a H en la fila 3
  const titleCell  = ws.getCell('B3');
  titleCell.value  = 'Instrucciones de sanciones a aplicar';
  titleCell.font   = {name:'Calibri', size:14, bold:true};
  titleCell.alignment = alignCC;   // Centrado horizontal y vertical
  titleCell.border = borderMedium; // Borde grueso alrededor del título
  titleCell.fill   = fillWhite;

  // ── Fila 5: Encabezados de columnas ──
  // Definición de cada encabezado con su columna y texto
  const hdrDefs = [
    {col:'B', v:'Número de\nsanción'},          // \n = salto de línea dentro de la celda
    {col:'C', v:'*Operación'},
    {col:'D', v:'CASFIM'},
    {col:'E', v:'Nombre Institución'},
    {col:'F', v:'IMPORTE'},
    {col:'G', v:'CUENTA DE RESULTADOS\n(Sin dígito verificador)'},
    {col:'H', v:'* TEXTO'},
  ];

  hdrDefs.forEach(({col,v}) => {
    const c = ws.getCell(`${col}5`);
    c.value     = v;
    c.font      = {name:'Calibri', size:11, bold:true};
    c.alignment = alignCCW;    // Con wrap para mostrar saltos de línea
    c.border    = borderThin;
    c.fill      = fillGray;   // Fondo gris para diferenciar encabezados de datos
  });

  // ── Fila 6: Datos del cobro ──
  const fontData = {name:'Calibri', size:11, bold:false};

  /**
   * Helper interno para escribir una celda de datos con estilo estándar.
   * @param {string} addr   - Dirección de celda (ej: 'B6')
   * @param {*}      val    - Valor a insertar
   * @param {Object} align  - Objeto de alineación ExcelJS
   * @param {string} [numFmt] - Formato numérico de Excel (ej: '#,##0.00')
   */
  const setData = (addr, val, align, numFmt) => {
    const c = ws.getCell(addr);
    c.value     = val;
    c.font      = fontData;
    c.alignment = align;
    c.border    = borderThin;
    c.fill      = fillWhite;
    if(numFmt) c.numFmt = numFmt; // Formato numérico (solo cuando se especifica)
  };

  setData('B6', 1,                          alignCC);           // Número de sanción: siempre 1
  setData('C6', operacion,                   alignCC);           // 'S' o 'D'
  setData('D6', parseInt(casfim)||casfim,    alignCC);           // CASFIM como número si es posible
  setData('E6', inst,                        alignLC);           // Nombre de la institución
  setData('F6', importe,                     alignCC, '#,##0.00'); // Importe con formato contable
  setData('G6', CUENTA,                      alignCC);           // Cuenta 385050311 (General — sin numFmt)
  setData('H6', texto,                       alignCC);           // TEXTO del layout (máx. 42 chars)

  // ── Filas 8-11: Notas informativas al pie del layout ──
  const setNota = (addr, val) => {
    const c = ws.getCell(addr);
    c.value     = val;
    c.font      = {name:'Calibri', size:11, bold:false};
    c.alignment = alignLC; // Alineación izquierda para texto descriptivo
  };
  setNota('B8',  'Notas:');
  setNota('B9',  '* Las aplicaciones son Valor Día Siguiente.');
  setNota('B10', '* En el campo "Operación" indicar si se trata de una Sanción (S) o Devolución (D).');
  setNota('B11', '* El campo de texto tiene un máximo de 42 caracteres');

  // ── Generar el archivo y descargarlo ──
  // writeBuffer() devuelve un ArrayBuffer que se convierte a Blob para descarga
  const buf  = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

  // Nombre: los "/" en la resolución se reemplazan por "_" para nombres de archivo válidos
  download(blob, `UL_LAYOUT_COBRO_${res.replace(/\//g,'_')}_${casfim}.xlsx`);
  showAlert('ok','✅ Excel generado correctamente');
}


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 11 — GENERACIÓN DEL MEMORANDO WORD (.docx)
// Genera el memorando oficial usando JSZip para construir manualmente
// la estructura OOXML (Office Open XML) del archivo .docx.
//
// Un .docx es un archivo ZIP que contiene varios archivos XML.
// Se construyen manualmente: document.xml, footer1.xml y los archivos
// de relaciones y tipos de contenido requeridos por la especificación OOXML.
//
// Nombre del archivo: UL_{no_memorando}.docx
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Genera y descarga el memorando de solicitud de cobro en formato Word (.docx).
 *
 * El memorando está dirigido a MARY CARMEN CABRERA RUEDA (Gerente de Gestión
 * de Operaciones) y solicita el cargo a la cuenta de la institución sancionada,
 * con el monto actualizado conforme al Art. 67 de la Ley del Banco de México.
 */
async function generateWord(){
  // Validar campos del Word (más estricto que el Excel)
  const errs = validate(true);
  if(errs.length){ showAlert('err','⚠ '+errs.join(' · ')); return; }

  // Leer todos los valores necesarios del formulario
  const inst         = document.getElementById('inst-nombre').value.trim();
  const imp          = parseFloat(document.getElementById('importe').value);
  const impFmt       = fmtMoney(imp);        // Ej: "$1,234.56"
  const impWords     = numToWords(imp);      // Ej: "MIL DOSCIENTOS TREINTA Y CUATRO PESOS 56/100 M.N."
  const memo         = document.getElementById('memo').value.trim();
  const fmemoStr     = fmtDate(document.getElementById('f-memo').value);      // Fecha del memorando en español
  const fescritoStr  = fmtDate(document.getElementById('f-escrito').value);   // Fecha del escrito en español
  // Para el memorando solo se necesita el DÍA de la fecha de recepción
  const diaRecep     = parseInt(document.getElementById('f-recep').value.split('-')[2]);
  const res          = document.getElementById('no-res').value.trim();
  const casfim       = document.getElementById('casfim').value.trim();

  // ── Funciones auxiliares de XML (definidas localmente para este documento) ──

  /**
   * Escapa caracteres especiales para que sean válidos dentro de XML/OOXML.
   * Necesario para evitar XML malformado cuando el texto contiene &, <, >, ".
   * @param {string} s - Texto a escapar
   * @returns {string} Texto con entidades XML
   */
  function x(s){
    return String(s)
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;');
  }

  /**
   * Genera el XML de propiedades de un run de texto (<w:rPr>).
   * Controla la apariencia tipográfica: fuente, tamaño, negrita, superíndice y color.
   *
   * @param {Object} opts - Opciones de formato:
   *   opts.font  {string}  - Nombre de la fuente (default: 'Calibri')
   *   opts.sz    {number}  - Tamaño en puntos (default: 11)
   *   opts.bold  {boolean} - true para negrita
   *   opts.sup   {boolean} - true para superíndice (notas al pie)
   *   opts.color {string}  - Color hexadecimal sin # (ej: 'FF0000')
   * @returns {string} XML del elemento <w:rPr>
   */
  function rPr(opts){
    const font = opts.font || 'Calibri';
    const sz   = (opts.sz || 11) * 2; // En OOXML el tamaño es en medios puntos
    let s = `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>`;
    s += `<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
    if(opts.bold)  s += '<w:b/><w:bCs/>';                              // Negrita
    if(opts.sup)   s += '<w:vertAlign w:val="superscript"/>';          // Superíndice
    if(opts.color) s += `<w:color w:val="${opts.color}"/>`;            // Color
    s += '</w:rPr>';
    return s;
  }

  /**
   * Genera el XML de propiedades de un párrafo (<w:pPr>).
   * Controla la alineación y la sangría izquierda del párrafo.
   *
   * @param {string} align - Alineación: 'both'(justificado), 'center', 'left', 'right'
   * @param {number} [ind] - Sangría izquierda en DXA (1/20 de punto tipográfico; 1440 DXA = 1 pulgada)
   * @returns {string} XML del elemento <w:pPr>
   */
  function pPr(align, ind){
    let s = `<w:pPr><w:jc w:val="${align||'both'}"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>`;
    if(ind) s += `<w:ind w:left="${ind}"/>`;
    // Marca de párrafo con fuente/tamaño fijos, para que párrafos vacíos no
    // hereden la fuente por defecto de Word (Aptos, más alta que Calibri) y
    // así no se acumule espacio extra que empuje la nota al pie a otra página.
    s += `<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>`;
    s += '</w:pPr>';
    return s;
  }

  /**
   * Genera un run de texto completo (<w:r>) con propiedades y contenido.
   * Un "run" en OOXML es el elemento mínimo de texto con formato uniforme.
   *
   * @param {string} text - Texto del run (se escapa automáticamente)
   * @param {Object} opts - Opciones de formato (ver rPr)
   * @returns {string} XML del elemento <w:r>
   */
  function run(text, opts){
    return `<w:r>${rPr(opts||{})}<w:t xml:space="preserve">${x(text)}</w:t></w:r>`;
  }

  /**
   * Genera un párrafo completo (<w:p>) con propiedades y contenido.
   * En OOXML, un párrafo puede contener múltiples runs con diferentes formatos.
   *
   * @param {string} content - XML de los runs del párrafo (generados con run())
   * @param {string} [align] - Alineación del párrafo (default: 'both' = justificado)
   * @param {number} [ind]   - Sangría izquierda en DXA
   * @returns {string} XML del elemento <w:p>
   */
  function para(content, align, ind){
    return `<w:p>${pPr(align, ind)}${content}</w:p>`;
  }

  /**
   * Genera el XML de una referencia REAL a nota al pie de Word (<w:footnoteReference>).
   * A diferencia de un superíndice manual, esto hace que Word dibuje la línea
   * separadora y calcule el espaciado de forma nativa (igual que al insertar
   * una nota al pie manualmente desde el menú Referencias de Word).
   *
   * @param {number} id - Id de la nota al pie (debe coincidir con footnotesXml)
   * @returns {string} XML del elemento <w:r> con la referencia
   */
  function footnoteRef(id){
    return `<w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:b/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
  }

  /**
   * Genera un párrafo vacío para espaciado vertical entre secciones.
   * Equivalente a presionar Enter en un documento Word.
   *
   * @param {string} [align] - Alineación (default: 'both')
   * @returns {string} XML de un párrafo vacío
   */
  function empty(align){
    return para('', align||'both');
  }

  // ── Estilos de texto predefinidos para el memorando ──
  // Cada objeto define la tipografía de un tipo de contenido específico
  const oAnio = {font:'Calibri', sz:10, bold:true};  // Encabezado con el año oficial
  const oMemo = {font:'Calibri', sz:11, bold:true};  // Título "M E M O R A N D O"
  const oCiud = {font:'Calibri', sz:11};              // Ciudad, fecha y referencia
  const oNomb = {font:'Calibri', sz:11, bold:true};  // Nombre del destinatario
  const oGte  = {font:'Calibri', sz:11};              // Cargo del destinatario
  const oC    = {font:'Calibri', sz:11};              // Texto corriente del cuerpo
  const oCB   = {font:'Calibri', sz:11, bold:true};  // Texto destacado (institución, monto, cuenta)
  const oCSup = {font:'Calibri', sz:11, bold:true, sup:true};  // Superíndice del importe (¹)
  const oNota = {font:'Calibri', sz:9};               // Texto de la nota al pie
  const oNSup = {font:'Times New Roman', sz:12, sup:true};     // Número superíndice de nota (Times New Roman)

  // ── Texto de la nota al pie (explicación legal del descuento por pronto pago) ──
  // Conforme al Art. 67, segundo párrafo, de la Ley del Banco de México
  const NOTA_TEXT = 'Monto determinado al aplicar el descuento de 20% por pronto pago al monto original, ' +
    'de conformidad con el artículo 67, segundo párrafo, de la Ley del Banco de México, ' +
    'en virtud de que el escrito fue recibido dentro del plazo de quince días hábiles ' +
    'siguientes a la fecha de notificación de la resolución.';

  // ── Construcción del cuerpo del documento (body XML) ──
  const body=`
    ${para(run('\u201c2026, A\u00f1o de Margarita Maza Parada\u201d',oAnio),'center')}
    ${empty('center')}
    ${para(run('M E M O R A N D O',oMemo),'center')}
    ${empty('center')}${empty()}
    ${para(run('Ciudad de M\u00e9xico, a '+fmemoStr+'.',oCiud),'both',4300)}
    ${para(run('Ref.: '+memo,oCiud),'both',4300)}
    ${empty()}${empty()}
    ${para(run('MARY CARMEN CABRERA RUEDA',oNomb))}
    ${para(run('Gerente de Gesti\u00f3n de Operaciones',oGte))}
    ${empty()}
    ${para(
      run('Nos referimos al escrito de fecha '+fescritoStr+', recibido por este Banco Central el d\u00eda '+diaRecep+' del mismo mes y a\u00f1o, mediante el cual ',oC) +
      run(inst,oCB) +
      run(', autoriz\u00f3 el cargo en la cuenta que le lleva Banco de M\u00e9xico, para cubrir el importe de la multa que le fue impuesta por este Instituto Central.',oC)
    )}
    ${empty()}
    ${para(
      run('Sobre el particular, y con la finalidad de concluir con el procedimiento de imposici\u00f3n de sanci\u00f3n, les solicitamos que efect\u00faen el cargo por ',oC) +
      run(impFmt+' ('+impWords+')',oCB) +    // Importe en número y letras (negrita)
      footnoteRef(1) +                       // Referencia REAL a nota al pie de Word (¹)
      run(', monto actualizado de conformidad con el art\u00edculo 67, de la Ley del Banco de M\u00e9xico, a la cuenta que le lleva Banco de M\u00e9xico a ',oC) +
      run(inst+',',oCB) +
      run(' y los fondos sean acreditados en la cuenta ',oC) +
      run('385.05.03.11-5',oCB) +           // Cuenta contable (negrita)
      run(' \u201cMultas impuestas conforme al Art. 36 Bis de la L.B.M. \u2013 DGSPIM.\u201d',oCB)
    )}
    ${empty()}
    ${para(run('Agradeceremos a ustedes realizar el tr\u00e1mite respectivo tomando en cuenta el archivo que se adjunta y les solicitamos nos env\u00eden copia del documento contable que se genere.',oC))}
    ${empty()}${empty()}
    ${para(run('A t e n t a m e n t e ,',oCB),'center')}
    ${empty('center')}${empty('center')}${empty('center')}${empty('center')}
    ${para(run('Gerencia de Supervisi\u00f3n de Sistemas de Pagos',oCB),'center')}
    ${para(run('e Infraestructuras de Mercado',oCB),'center')}
    <w:p>
      <w:pPr><w:pageBreakBefore/><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
        <w:t>Documento firmado digitalmente, su validaci\u00f3n requiere hacerse electr\u00f3nicamente.</w:t>
      </w:r>
    </w:p>
  `;

  // ── XML del pie de página ──
  // Aparece en todas las páginas del documento con la leyenda "Uso Limitado"
  const footerXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="C00000"/></w:rPr><w:t>Uso Limitado</w:t></w:r></w:p>
  <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr><w:t>Informaci\u00f3n cuyo acceso est\u00e1 restringido a un grupo limitado de personas empleadas por el Banco de M\u00e9xico y, en su caso, personas ajenas al mismo.</w:t></w:r></w:p>
</w:ftr>`;

  // ── XML de la nota al pie REAL (word/footnotes.xml) ──
  // Contiene el separador (la línea, la dibuja Word), y el texto de la nota
  // con id="1", que es el que referencia footnoteRef(1) en el cuerpo.
  const footnotesXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1">
    <w:p>
      <w:pPr><w:jc w:val="both"/><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="18"/><w:szCs w:val="18"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve"> ${x(NOTA_TEXT)}</w:t></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>`;

  // ── XML principal del documento ──
  // Define la página (tamaño carta US) y márgenes, y referencia el footer
  const docXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>${body}
<w:sectPr>
  <w:footerReference w:type="default" r:id="rId1"/>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1701" w:bottom="720" w:left="1701" w:header="709" w:footer="709" w:gutter="0"/>
</w:sectPr>
</w:body></w:document>`;

  // ── Archivos de relaciones y tipos de contenido OOXML ──
  // Son obligatorios para que Word pueda abrir el archivo correctamente

  // Relaciones dentro de la carpeta word/ (document.xml → footer1.xml)
  const relsXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>`;

  // Tipos de contenido MIME para cada parte del ZIP
  const contentTypes=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>`;

  // Relación raíz del paquete ZIP (_rels/.rels → word/document.xml)
  const mainRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  // ── Construir el ZIP (= el .docx) con JSZip ──
  const zip = new JSZip();
  zip.file('[Content_Types].xml', contentTypes);       // Tipos MIME del paquete
  zip.file('_rels/.rels',         mainRels);           // Relación raíz
  zip.file('word/document.xml',   docXml);             // Cuerpo del documento
  zip.file('word/footer1.xml',    footerXml);          // Pie de página
  zip.file('word/footnotes.xml',  footnotesXml);        // Nota al pie real
  zip.file('word/_rels/document.xml.rels', relsXml);  // Relaciones del documento

  // Generar el blob con el tipo MIME correcto para .docx
  const blob = await zip.generateAsync({
    type:'blob',
    mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  });

  download(blob, `UL_${memo}.docx`);
  showAlert('ok', `✅ Memorando generado: UL_${memo}.docx`);
}


// ══════════════════════════════════════════════════════════════════════════════
// MÓDULO 12 — GENERACIÓN CONJUNTA
// Genera ambos documentos (Excel y Word) en secuencia con un pequeño delay
// entre ellos para evitar conflictos al descargar dos archivos simultáneos.
// ══════════════════════════════════════════════════════════════════════════════

/**
 * Genera el layout Excel y el memorando Word en secuencia.
 * Valida los campos de ambos documentos antes de iniciar.
 * Usa un delay de 600ms entre la generación del Excel y el Word
 * para que el navegador no bloquee la segunda descarga.
 */
function generateAll(){
  // Combinar errores de ambas validaciones, eliminando duplicados con Set
  const errs = [...new Set([...validate(false), ...validate(true)])];
  if(errs.length){ showAlert('err','\u26a0 '+errs.join(' \u00b7 ')); return; }

  // Generar Excel primero, luego Word con delay
  generateExcel().then(() => setTimeout(() => generateWord(), 600));
}


// ── INICIALIZACIÓN ────────────────────────────────────────────────────────────
// Renderizar la vista previa con valores vacíos al cargar la página,
// para que la sección de preview esté lista cuando el usuario empiece a llenar
// el formulario directamente (sin cargar un Excel).
updPreview();