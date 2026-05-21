// ══════════════════════════════════════════════════════
// PROYECTO V2 — Generador con carga desde Excel
// ══════════════════════════════════════════════════════

let registros = [];   // datos cargados del Excel
let selReg   = null;  // registro seleccionado

// ── UTILS ──
function fmtMoney(n){ return '$'+n.toLocaleString('es-MX',{minimumFractionDigits:2,maximumFractionDigits:2}); }
function fmtDate(d){
  if(!d) return '';
  const[y,m,day]=d.split('-');
  const M=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  return `${parseInt(day)} de ${M[parseInt(m)-1]} de ${y}`;
}
function excelDateToISO(val){
  // SheetJS puede devolver Date o número serial de Excel
  if(!val) return '';
  if(val instanceof Date){
    return val.toISOString().split('T')[0];
  }
  if(typeof val === 'number'){
    // número serial Excel → Date
    const d = new Date(Math.round((val - 25569)*86400*1000));
    return d.toISOString().split('T')[0];
  }
  if(typeof val === 'string'){
    // intentar parsear
    const d = new Date(val);
    if(!isNaN(d)) return d.toISOString().split('T')[0];
  }
  return '';
}
function numToWords(num){
  const ones=['','UNO','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE','DIEZ','ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISÉIS','DIECISIETE','DIECIOCHO','DIECINUEVE'];
  const tens=['','DIEZ','VEINTE','TREINTA','CUARENTA','CINCUENTA','SESENTA','SETENTA','OCHENTA','NOVENTA'];
  const huns=['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS','SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS'];
  function three(n){
    if(n===0)return'';if(n===100)return'CIEN';
    const h=Math.floor(n/100),rem=n%100;
    let r=h?huns[h]+' ':'';
    if(rem<20&&rem>0)r+=ones[rem];
    else{const t=Math.floor(rem/10),o=rem%10;if(t)r+=tens[t];if(o)r+=(t?' Y ':'')+ones[o];}
    return r.trim();
  }
  function thousands(n){
    // convierte hasta 999,999
    if(n===0) return '';
    const th=Math.floor(n/1000), rest=n%1000;
    let r='';
    if(th>0) r+=(th===1?'MIL':three(th)+' MIL')+' ';
    if(rest>0) r+=three(rest);
    return r.trim();
  }
  const ip=Math.floor(num),dp=Math.round((num-ip)*100);
  let r='';
  if(ip>=1000000000000){
    // billones (1,000,000,000,000)
    const b=Math.floor(ip/1000000000000);
    r+=(b===1?'UN BILLÓN':thousands(b)+' BILLONES')+' ';
    const rest=ip%1000000000000;
    if(rest>=1000000){
      const m=Math.floor(rest/1000000);
      r+=(m===1?'UN MILLÓN':thousands(m)+' MILLONES')+' ';
      const rest2=rest%1000000;
      if(rest2>0) r+=thousands(rest2);
    } else if(rest>0){ r+=thousands(rest); }
  } else if(ip>=1000000){
    const m=Math.floor(ip/1000000);
    r+=(m===1?'UN MILLÓN':thousands(m)+' MILLONES')+' ';
    const rest=ip%1000000;
    if(rest>0) r+=thousands(rest);
  } else {
    r=thousands(ip);
  }
  return r.trim()+` PESOS ${String(dp).padStart(2,'0')}/100 M.N.`;
}

// ── CARGA DEL EXCEL ──
document.getElementById('btn-cargar').addEventListener('click', () => {
  document.getElementById('file-input').click();
});

document.getElementById('file-input').addEventListener('change', function(e){
  const file = e.target.files[0];
  if(!file){ return; }

  // Mostrar spinner inmediatamente
  const lbl = document.getElementById('lbl-archivo');
  const btn = document.getElementById('btn-cargar');
  lbl.textContent = '⏳ Cargando archivo...';
  lbl.style.color = '#c8a96e';
  btn.disabled = true;
  btn.style.opacity = '0.6';

  const reader = new FileReader();
  reader.onload = function(ev){
    // Usar setTimeout para no bloquear el UI mientras procesa
    setTimeout(() => {
      try{
        const data = new Uint8Array(ev.target.result);

        // Opciones optimizadas: solo leer valores, sin estilos ni fórmulas
        const wb = XLSX.read(data, {
          type      : 'array',
          cellDates : true,
          cellStyles: false,   // no leer estilos → más rápido
          cellNF    : false,   // no leer formatos de número
          cellHTML  : false,   // no generar HTML
          sheetRows : 3000,    // limitar filas máximas
        });

        const ws = wb.Sheets['Sanciones Impuestas'];
        if(!ws){
          showAlert('err','No se encontró la hoja "Sanciones Impuestas".');
          resetBtn(btn, lbl, file.name, 0, true);
          return;
        }

        // Leer solo las columnas que necesitamos usando header array
        // para evitar procesar todas las columnas del Excel
        const rows = XLSX.utils.sheet_to_json(ws, {
          range  : 1,        // header en fila 2
          defval : '',
          raw    : false,    // convertir fechas a string directamente
        });

        // Columnas que nos interesan (ignorar el resto)
        const COLS_NEEDED = new Set([
          'Institución','CASFIM','Expediente','No. Resolución',
          'Fecha autorización de cobro',
          'Fecha de recepción de autorización de cobro',
          'No. de Memorando','Fecha de solicitud de cobro',
          'Monto impuesto de la sanción','Nombre corto'
        ]);

        registros = rows
          .filter(r => r['Institución'] && String(r['Institución']).trim() !== ''
                    && r['No. Resolución'] && String(r['No. Resolución']).trim() !== '')
          .map(r => ({
            institucion : String(r['Institución']||'').trim(),
            casfim      : String(r['CASFIM']||'').replace(/\.0$/,'').trim(),
            expediente  : String(r['Expediente']||'').trim(),
            no_res      : String(r['No. Resolución']||'').trim(),
            f_escrito   : excelDateToISO(r['Fecha autorización de cobro']),
            f_recep     : excelDateToISO(r['Fecha de recepción de autorización de cobro']),
            memo        : String(r['No. de Memorando']||'').trim(),
            f_memo      : excelDateToISO(r['Fecha de solicitud de cobro']),
            monto       : parseFloat(String(r['Monto impuesto de la sanción']).replace(/,/g,''))||0,
            nombre_corto: String(r['Nombre corto']||'').trim(),
          }));

        mostrarTabla(registros);
        document.getElementById('sec-tabla').style.display='block';
        resetBtn(btn, lbl, file.name, registros.length, false);
        showAlert('ok', `✅ Se cargaron <strong>${registros.length}</strong> registros correctamente.`);

      } catch(err){
        showAlert('err','Error al leer el archivo: '+err.message);
        resetBtn(btn, lbl, file.name, 0, true);
      }
    }, 50); // pequeño delay para que el spinner se muestre antes de procesar
  };

  reader.readAsArrayBuffer(file);
  this.value='';
});

function resetBtn(btn, lbl, filename, count, error){
  btn.disabled = false;
  btn.style.opacity = '1';
  if(error){
    lbl.textContent = '❌ Error al cargar el archivo';
    lbl.style.color = '#b83232';
  } else {
    lbl.textContent = `✓ ${filename} — ${count} registros cargados`;
    lbl.style.color = '#2d7a4f';
  }
}

// ── TABLA DE REGISTROS ──
function mostrarTabla(datos){
  const tbody = document.getElementById('tbl-body');
  tbody.innerHTML = '';

  // Usar DocumentFragment para renderizar todo de una sola vez (mucho más rápido)
  const frag = document.createDocumentFragment();

  datos.forEach((r,i) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${r.institucion}</td>
      <td>${r.memo||'—'}</td>
      <td>${r.expediente}</td>
      <td>${r.no_res}</td>
    `;
    tr.style.cursor = 'pointer';
    // Guardar índice en el elemento para recuperar el registro rápido
    tr.dataset.idx = i;
    tr.addEventListener('click', function(){
      seleccionarRegistro(datos[this.dataset.idx]);
      resaltarFila(this);
    });
    frag.appendChild(tr);
  });

  tbody.appendChild(frag); // un solo repaint
}

function resaltarFila(tr){
  document.querySelectorAll('#tbl-body tr').forEach(t=>t.classList.remove('sel'));
  tr.classList.add('sel');
}

// ── BÚSQUEDA EN TABLA ──
document.getElementById('buscador').addEventListener('input', function(){
  const q = this.value.trim().toLowerCase();
  const filtrados = q.length < 2
    ? registros
    : registros.filter(r =>
        r.institucion.toLowerCase().includes(q) ||
        r.no_res.toLowerCase().includes(q) ||
        r.expediente.toLowerCase().includes(q) ||
        r.memo.toLowerCase().includes(q)
      );
  mostrarTabla(filtrados);
});

// ── SELECCIONAR REGISTRO → AUTOCOMPLETE EDITABLE ──
function seleccionarRegistro(r){
  selReg = r;

  // Autocompletar todos los campos (todos editables)
  setVal('inst-nombre', r.institucion);
  setVal('casfim',      r.casfim);
  setVal('no-res',      r.no_res);
  setVal('exp',         r.expediente);
  setVal('importe',     '');
  setVal('f-escrito',   r.f_escrito);
  setVal('f-recep',     r.f_recep);
  setVal('memo',        r.memo);
  setVal('f-memo',      r.f_memo);

  // Autogenerar TEXTO
  const tf = document.getElementById('texto');
  tf.removeAttribute('data-manual');
  autoTxt();

  // Mostrar sección de formulario
  document.getElementById('sec-form').style.display='block';
  document.getElementById('sec-preview').style.display='block';
  document.getElementById('sec-generar').style.display='block';
  document.getElementById('sec-form').scrollIntoView({behavior:'smooth', block:'start'});

  updImporte();
  updCC();
  updPreview();
}

function setVal(id, val){
  const el = document.getElementById(id);
  if(el) el.value = val || '';
}

// ── AUTO-TEXTO ──
document.getElementById('no-res').addEventListener('input', ()=>{ autoTxt(); updPreview(); });
document.getElementById('exp').addEventListener('input',    ()=>{ autoTxt(); updPreview(); });
function autoTxt(){
  const r  = document.getElementById('no-res').value.trim();
  const e  = document.getElementById('exp').value.trim();
  const tf = document.getElementById('texto');
  if(!tf.dataset.manual && r && e){
    tf.value = (r+'/'+e).slice(0,42);
    updCC();
  }
}
document.getElementById('texto').addEventListener('input', function(){ this.dataset.manual='1'; updCC(); updPreview(); });
function updCC(){
  const v=document.getElementById('texto').value, el=document.getElementById('cc');
  el.textContent=v.length+' / 42';
  el.className='cc'+(v.length>42?' over':v.length>35?' warn':'');
}

function updImporte(){
  updPreview();
}
document.getElementById('importe').addEventListener('input', updImporte);

['inst-nombre','casfim','no-res','exp','memo','f-memo','f-escrito','f-recep','operacion'].forEach(id=>{
  const el=document.getElementById(id);
  if(el){ el.addEventListener('input',updPreview); el.addEventListener('change',updPreview); }
});

// ── PREVIEW ──
function updPreview(){
  const inst   = document.getElementById('inst-nombre').value||'[Institución]';
  const casfim = document.getElementById('casfim').value||'—';
  const raw    = parseFloat(document.getElementById('importe').value)||0;
  const texto  = document.getElementById('texto').value||'—';
  const memo   = document.getElementById('memo').value||'—';
  const fm     = document.getElementById('f-memo').value;
  const fe     = document.getElementById('f-escrito').value;

  document.getElementById('preview').innerHTML=`
    <div style="margin-bottom:10px;padding-bottom:10px;border-bottom:1px solid #e8e4de">
      <strong style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#6b6b6b">Layout Excel (BBVA)</strong>
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px 16px;font-size:13px;margin-bottom:14px">
      <div><span class="tag">Institución</span><br>${inst}</div>
      <div><span class="tag">CASFIM</span><br>${casfim}</div>
      <div><span class="tag">Importe</span><br>${raw>0?fmtMoney(raw):'—'}</div>
      <div><span class="tag">Cuenta Resultados</span><br>385050311</div>
      <div><span class="tag">TEXTO</span><br>${texto}</div>
    </div>
    <div style="margin:10px 0;padding:10px 0;border-top:1px solid #e8e4de;border-bottom:1px solid #e8e4de">
      <strong style="font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:#6b6b6b">Memorando Word</strong>
    </div>
    <div style="font-size:13px;line-height:2">
      <strong>Ref.:</strong> ${memo}<br>
      <strong>Fecha:</strong> ${fm?fmtDate(fm):'—'}<br>
      <strong>Institución:</strong> ${inst}<br>
      <strong>Monto:</strong> ${raw>0?fmtMoney(raw):'—'}<br>
      <strong>Fecha escrito:</strong> ${fe?fmtDate(fe):'—'}
    </div>`;
}

// ── VALIDACIÓN ──
function validate(word=false){
  const e=[];
  if(!document.getElementById('inst-nombre').value.trim()) e.push('Nombre de institución es obligatorio');
  if(!document.getElementById('no-res').value.trim()) e.push('No. Resolución es obligatorio');
  if(!document.getElementById('exp').value.trim()) e.push('Expediente es obligatorio');
  if(!(parseFloat(document.getElementById('importe').value)>0)) e.push('Ingresa el importe');
  const t=document.getElementById('texto').value.trim();
  if(!t) e.push('TEXTO es obligatorio');
  if(t.length>42) e.push('TEXTO excede 42 caracteres');
  if(word){
    if(!document.getElementById('memo').value.trim()) e.push('No. Memorando es obligatorio');
    if(!document.getElementById('f-memo').value) e.push('Fecha del memorando es obligatoria');
    if(!document.getElementById('f-escrito').value) e.push('Fecha del escrito es obligatoria');
    if(!document.getElementById('f-recep').value) e.push('Fecha de recepción es obligatoria');
  }
  return e;
}
function showAlert(type,msg){
  document.getElementById('alert-area').innerHTML=`<div class="alert al-${type}">${msg}</div>`;
  window.scrollTo({top:0,behavior:'smooth'});
  setTimeout(()=>document.getElementById('alert-area').innerHTML='',7000);
}
function download(blob,name){ const u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=name;a.click();URL.revokeObjectURL(u); }

// ── GENERAR EXCEL (ExcelJS — soporta estilos completos) ──
async function generateExcel(){
  const errs=validate(false);if(errs.length){showAlert('err','⚠ '+errs.join(' · '));return;}
  const inst     = document.getElementById('inst-nombre').value.trim();
  const casfim   = document.getElementById('casfim').value.trim();
  const importe  = parseFloat(document.getElementById('importe').value);
  const operacion= document.getElementById('operacion').value[0];
  const texto    = document.getElementById('texto').value.trim();
  const res      = document.getElementById('no-res').value.trim();
  const CUENTA   = '385050311';

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('LAYOUT', {
    properties: { tabColor: {argb:'FFFFFFFF'} }
  });

  // Fondo blanco para toda la hoja
  ws.views = [{showGridLines: false}];

  // Anchos de columna
  ws.columns = [
    {width: 6},  // A vacía
    {width: 12},   // B Número sanción
    {width: 11},   // C *Operación
    {width: 9},    // D CASFIM
    {width: 88},   // E Nombre Institución (más ancho)
    {width: 26},   // F IMPORTE
    {width: 26},   // G CUENTA DE RESULTADOS
    {width: 36},   // H TEXTO
  ];

  // ── Alturas de fila ──
  // Filas vacías 1,2,4,7 con contenido de espacio y fuente visible para forzar altura
  [1, 2, 4, 7].forEach(n => {
    const row  = ws.getRow(n);
    row.height = 15;
    const cell = ws.getCell(`A${n}`);
    cell.value = ' ';
    cell.font  = {name:'Calibri', size:11};
  });
  ws.getRow(3).height  = 22;  // título
  ws.getRow(5).height  = 42;  // headers
  ws.getRow(6).height  = 16;  // datos
  ws.getRow(8).height  = 14;  // Notas:
  ws.getRow(9).height  = 14;
  ws.getRow(10).height = 14;
  ws.getRow(11).height = 14;

  // ── Estilos ──
  const borderMedium = {
    top   :{style:'medium',color:{argb:'FF000000'}},
    bottom:{style:'medium',color:{argb:'FF000000'}},
    left  :{style:'medium',color:{argb:'FF000000'}},
    right :{style:'medium',color:{argb:'FF000000'}},
  };
  const borderThin = {
    top   :{style:'thin',color:{argb:'FF000000'}},
    bottom:{style:'thin',color:{argb:'FF000000'}},
    left  :{style:'thin',color:{argb:'FF000000'}},
    right :{style:'thin',color:{argb:'FF000000'}},
  };
  const fillGray  = {type:'pattern',pattern:'solid',fgColor:{argb:'FFD9D9D9'}};
  const fillWhite = {type:'pattern',pattern:'solid',fgColor:{argb:'FFFFFFFF'}};
  const alignCC  = {horizontal:'center', vertical:'middle', wrapText:false};
  const alignCCW = {horizontal:'center', vertical:'middle', wrapText:true};
  const alignLC  = {horizontal:'left',   vertical:'middle', wrapText:false};

  // ── Fila 3: Título fusionado B3:H3 ──
  ws.mergeCells('B3:H3');
  const titleCell  = ws.getCell('B3');
  titleCell.value  = 'Instrucciones de sanciones a aplicar';
  titleCell.font   = {name:'Calibri', size:14, bold:true};
  titleCell.alignment = alignCC;
  titleCell.border = borderMedium;
  titleCell.fill   = fillWhite;

  // ── Fila 5: Headers ──
  const hdrDefs = [
    {col:'B', v:'Número de\nsanción'},
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
    c.alignment = alignCCW;
    c.border    = borderThin;
    c.fill      = fillGray;  // Blanco, Fondo 1, Oscuro 15%
  });

  // ── Fila 6: Datos ──
  const fontData = {name:'Calibri', size:11, bold:false};

  const setData = (addr, val, align, numFmt) => {
    const c = ws.getCell(addr);
    c.value     = val;
    c.font      = fontData;
    c.alignment = align;
    c.border    = borderThin;
    c.fill      = fillWhite;
    if(numFmt) c.numFmt = numFmt;
  };

  setData('B6', 1,                          alignCC);
  setData('C6', operacion,                   alignCC);
  setData('D6', parseInt(casfim)||casfim,    alignCC);
  setData('E6', inst,                        alignCC);
  setData('F6', importe,                     alignCC, '#,##0.00');  // tipo número
  setData('G6', CUENTA, alignCC);  // formato General — sin numFmt
  setData('H6', texto,                       alignCC);

  // ── Notas ──
  const setNota = (addr, val, bold) => {
    const c = ws.getCell(addr);
    c.value     = val;
    c.font      = {name:'Calibri', size:11, bold:false};
    c.alignment = alignLC;
  };
  setNota('B8',  'Notas:',                                                                                         false);
  setNota('B9',  '* Las aplicaciones son Valor Día Siguiente.',                                                    false);
  setNota('B10', '* En el campo "Operación" indicar si se trata de una Sanción (S) o Devolución (D).',            false);
  setNota('B11', '* El campo de texto tiene un máximo de 42 caracteres',                                           false);

  // ── Descargar ──
  const buf  = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  download(blob, `UL_LAYOUT_COBRO_${res.replace(/\//g,'_')}_${casfim}.xlsx`);
  showAlert('ok','✅ Excel generado correctamente');
}

// ── GENERAR WORD ──
async function generateWord(){
  const errs=validate(true);if(errs.length){showAlert('err','⚠ '+errs.join(' · '));return;}
  const inst      = document.getElementById('inst-nombre').value.trim();
  const imp       = parseFloat(document.getElementById('importe').value);
  const impFmt    = fmtMoney(imp);
  const impWords  = numToWords(imp);
  const memo      = document.getElementById('memo').value.trim();
  const fmemoStr  = fmtDate(document.getElementById('f-memo').value);
  const fescritoStr = fmtDate(document.getElementById('f-escrito').value);
  const diaRecep  = parseInt(document.getElementById('f-recep').value.split('-')[2]);
  const res       = document.getElementById('no-res').value.trim();
  const casfim    = document.getElementById('casfim').value.trim();

  function x(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
  function rPr(opts){
    const font=opts.font||'Calibri',sz=(opts.sz||11)*2;
    let s=`<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>`;
    s+=`<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>`;
    if(opts.bold)s+='<w:b/><w:bCs/>';
    if(opts.sup)s+='<w:vertAlign w:val="superscript"/>';
    if(opts.color)s+=`<w:color w:val="${opts.color}"/>`;
    s+='</w:rPr>';return s;
  }
  function pPr(align,ind){
    let s=`<w:pPr><w:jc w:val="${align||'both'}"/><w:spacing w:before="0" w:after="0"/>`;
    if(ind)s+=`<w:ind w:left="${ind}"/>`;
    s+='</w:pPr>';return s;
  }
  function run(text,opts){return`<w:r>${rPr(opts||{})}<w:t xml:space="preserve">${x(text)}</w:t></w:r>`;}
  function para(content,align,ind){return`<w:p>${pPr(align,ind)}${content}</w:p>`;}
  function empty(align){return para('',align||'both');}

  const oAnio ={font:'Calibri',sz:10,bold:true};
  const oMemo ={font:'Calibri',sz:11,bold:true};
  const oCiud ={font:'Calibri',sz:11};
  const oNomb ={font:'Calibri',sz:11,bold:true};
  const oGte  ={font:'Calibri',sz:11};
  const oC    ={font:'Calibri',sz:11};
  const oCB   ={font:'Calibri',sz:11,bold:true};
  const oCSup ={font:'Calibri',sz:11,bold:true,sup:true};
  const oNota ={font:'Calibri',sz:9};
  const oNSup ={font:'Times New Roman',sz:12,sup:true};

  const NOTA_TEXT='Monto determinado al aplicar el descuento de 20% por pronto pago al monto original, de conformidad con el art\u00edculo 67, segundo p\u00e1rrafo, de la Ley del Banco de M\u00e9xico, en virtud de que el escrito fue recibido dentro del plazo de quince d\u00edas h\u00e1biles siguientes a la fecha de notificaci\u00f3n de la resoluci\u00f3n.';

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
    ${para(run('Nos referimos al escrito de fecha '+fescritoStr+', recibido por este Banco Central el d\u00eda '+diaRecep+' del mismo mes y a\u00f1o, mediante el cual ',oC)+run(inst,oCB)+run(', autoriz\u00f3 el cargo en la cuenta que le lleva Banco de M\u00e9xico, para cubrir el importe de la multa que le fue impuesta por este Instituto Central.',oC))}
    ${empty()}
    ${para(run('Sobre el particular, y con la finalidad de concluir con el procedimiento de imposici\u00f3n de sanci\u00f3n, les solicitamos que efect\u00faen el cargo por ',oC)+run(impFmt+' ('+impWords+')',oCB)+run('1',oCSup)+run(', monto actualizado de conformidad con el art\u00edculo 67, de la Ley del Banco de M\u00e9xico, a la cuenta que le lleva Banco de M\u00e9xico a ',oC)+run(inst+',',oCB)+run(' y los fondos sean acreditados en la cuenta ',oC)+run('385.05.03.11-5',oCB)+run(' \u201cMultas impuestas conforme al Art. 36 Bis de la L.B.M. \u2013 DGSPIM.\u201d',oCB))}
    ${empty()}
    ${para(run('Agradeceremos a ustedes realizar el tr\u00e1mite respectivo tomando en cuenta el archivo que se adjunta y les solicitamos nos env\u00eden copia del documento contable que se genere.',oC))}
    ${empty()}${empty()}
    ${para(run('A t e n t a m e n t e ,',oCB),'center')}
    ${empty('center')}${empty('center')}${empty('center')}${empty('center')}
    ${para(run('Gerencia de Supervisi\u00f3n de Sistemas de Pagos',oCB),'center')}
    ${para(run('e Infraestructuras de Mercado',oCB),'center')}
    <w:p>
      <w:pPr>
        <w:jc w:val="left"/>
        <w:spacing w:before="1200" w:after="0"/>
        <w:keepWithNext/>
        <w:ind w:right="6000"/>
        <w:pBdr><w:top w:val="single" w:sz="6" w:space="4" w:color="000000"/></w:pBdr>
      </w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="18"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:jc w:val="left"/><w:spacing w:before="0" w:after="0"/><w:keepLines/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>1</w:t></w:r>
      <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr><w:t xml:space="preserve"> ${x(NOTA_TEXT)}</w:t></w:r>
    </w:p>
    <w:p>
      <w:pPr><w:pageBreakBefore/><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
        <w:t>Documento firmado digitalmente, su validaci\u00f3n requiere hacerse electr\u00f3nicamente.</w:t>
      </w:r>
    </w:p>
  `;

  const footerXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/><w:color w:val="C00000"/></w:rPr><w:t>Uso Limitado</w:t></w:r></w:p>
  <w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>
    <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="12"/><w:szCs w:val="12"/></w:rPr><w:t>Informaci\u00f3n cuyo acceso est\u00e1 restringido a un grupo limitado de personas empleadas por el Banco de M\u00e9xico y, en su caso, personas ajenas al mismo.</w:t></w:r></w:p>
</w:ftr>`;

  const docXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<w:body>${body}
<w:sectPr>
  <w:footerReference w:type="default" r:id="rId1"/>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1701" w:bottom="720" w:left="1701" w:header="709" w:footer="709" w:gutter="0"/>
</w:sectPr>
</w:body></w:document>`;

  const relsXml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;
  const contentTypes=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
</Types>`;
  const mainRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  const zip=new JSZip();
  zip.file('[Content_Types].xml',contentTypes);
  zip.file('_rels/.rels',mainRels);
  zip.file('word/document.xml',docXml);
  zip.file('word/footer1.xml',footerXml);
  zip.file('word/_rels/document.xml.rels',relsXml);
  const blob=await zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document'});
  download(blob,`UL_${memo}.docx`);
  showAlert('ok',`\u2705 Memorando generado: UL_${memo}.docx`);
}

function generateAll(){
  const errs=[...new Set([...validate(false),...validate(true)])];
  if(errs.length){showAlert('err','\u26a0 '+errs.join(' \u00b7 '));return;}
  generateExcel().then(()=>setTimeout(()=>generateWord(),600));
}

updPreview();