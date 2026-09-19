/* ================================================================
   Reporte de Dispensación — app local unificada
   Fuente única: archivo "Reporte de Dispensación" (.xlsx / .csv)
   100% en el navegador. Sin nube.
   ================================================================ */
'use strict';

/* ---------- Estado global ---------- */
var RAW = [];        // filas enriquecidas (nivel línea)
var META = { archivo:'', fecha:'', filasCrudas:0, filasUsadas:0 };
var FILTROS = { bodega:'', contrato:'', eps:'', mes:'' };

/* ---------- Utilidades DOM ---------- */
function $(s,c){ return (c||document).querySelector(s); }
function $$(s,c){ return Array.prototype.slice.call((c||document).querySelectorAll(s)); }
function el(tag,cls,html){ var e=document.createElement(tag); if(cls)e.className=cls; if(html!=null)e.innerHTML=html; return e; }
function toast(msg,err){ var t=$('#toast'); t.textContent=msg; t.className='toast show'+(err?' error':''); clearTimeout(t._t); t._t=setTimeout(function(){t.className='toast';},2600); }

/* ---------- Normalización de texto ---------- */
function norm(v){ if(v==null) return ''; return String(v).replace(/\s+/g,' ').trim(); }
function up(v){ return norm(v).toUpperCase(); }
function sinAcentos(v){ return up(v).normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function toNum(v){ if(v==null||v==='') return 0; var n=parseFloat(String(v).replace(/[^0-9.\-]/g,'')); return isNaN(n)?0:n; }
function fmt(n){ return (n||0).toLocaleString('es-CO'); }
function pct(a,b){ return b>0 ? (a/b*100) : 0; }
function pct1(a,b){ return (Math.round(pct(a,b)*10)/10); }
function pctCls(p){ return p>=90?'pct-good':(p>=70?'pct-mid':'pct-bad'); }

/* ---------- Alias de columnas del Reporte ---------- */
/* Cada campo lógico -> lista de posibles encabezados (normalizados sin acentos) */
var ALIASES = {
  documento:      ['NUMERO DOCUMENTO','NRO DOCUMENTO','NUM DOCUMENTO','DOCUMENTO','NO DOCUMENTO','NUMERO DE DOCUMENTO','CONSECUTIVO'],
  bodega:         ['BODEGA','PUNTO','PUNTO DISPENSACION','PUNTO DE DISPENSACION','SUCURSAL','FARMACIA','SEDE'],
  estado:         ['ESTADO','ESTADO DISPENSACION','ESTADO REGISTRO'],
  contrato:       ['CONTRATO','TIPO CONTRATO','MODALIDAD','REGIMEN CONTRATO'],
  eps:            ['EPS','ENTIDAD','ASEGURADOR','ASEGURADORA','ENTIDAD RESPONSABLE','PAGADOR'],
  entregado:      ['CANTIDAD ENTREGADA','CANT ENTREGADA','ENTREGADO','UNIDADES ENTREGADAS','CANTIDAD DISPENSADA','DISPENSADO'],
  formulado:      ['CANTIDAD FORMULADA','CANT FORMULADA','FORMULADO','UNIDADES FORMULADAS','CANTIDAD ORDENADA','PRESCRITO','CANTIDAD PRESCRITA'],
  pendiente:      ['CANTIDAD PENDIENTE','CANT PENDIENTE','PENDIENTE','SALDO','UNIDADES PENDIENTES'],
  soporte:        ['SOPORTE','TIENE SOPORTE','SOPORTE EVENTO','NRO SOPORTE','NUMERO SOPORTE','NO SOPORTE','FACTURA','AUTORIZACION'],
  codigo:         ['CODIGO','CODIGO MEDICAMENTO','COD MEDICAMENTO','CUM','COD PRODUCTO','CODIGO PRODUCTO','COD'],
  descripcion:    ['DESCRIPCION','DESCRIPCION MEDICAMENTO','MEDICAMENTO','PRODUCTO','DESCRIPCION PRODUCTO','NOMBRE MEDICAMENTO','ARTICULO'],
  cie10:          ['DESCRIPCION CIE 10','DESCRIPCION CIE10','CIE 10','CIE10','DIAGNOSTICO','DX','DESCRIPCION DIAGNOSTICO','DIAGNOSTICO PRINCIPAL'],
  codcie:         ['CODIGO CIE 10','CODIGO CIE10','COD CIE 10','COD CIE10','CIE'],
  paciente:       ['DOCUMENTO PACIENTE','IDENTIFICACION','IDENTIFICACION PACIENTE','NRO IDENTIFICACION','CEDULA','DOC PACIENTE','ID PACIENTE','NUMERO IDENTIFICACION'],
  paciNombre:     ['NOMBRE PACIENTE','PACIENTE','NOMBRE DEL PACIENTE','NOMBRE'],
  usuarioCrea:    ['USUARIO CREACION','USUARIO CREA','USUARIO','USUARIO DISPENSA','CREADO POR','DISPENSADO POR','USUARIO REGISTRO'],
  fecha:          ['FECHA','FECHA DISPENSACION','FECHA DISPENSA','FECHA CREACION','FECHA ENTREGA','FECHA REGISTRO','FECHA DOCUMENTO']
};

/* Placeholder para siguientes bloques */
/* ================================================================
   PARSEO Y LIMPIEZA DEL ARCHIVO
   Reglas:
   - Omitir fila 1 (metadata MEDISFARMA).
   - Usar fila 2 como encabezados.
   - Eliminar filas incompletas.
   - Reconstruir filas divididas / corridas.
   ================================================================ */

/* Lee un archivo local a matriz (AoA) con SheetJS */
function leerArchivo(file, cb){
  var reader = new FileReader();
  reader.onload = function(e){
    try{
      var data = new Uint8Array(e.target.result);
      var wb = XLSX.read(data, { type:'array', cellDates:false, raw:false });
      var ws = wb.Sheets[wb.SheetNames[0]];
      var aoa = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', raw:false, blankrows:false });
      cb(null, aoa);
    }catch(err){ cb(err); }
  };
  reader.onerror = function(){ cb(new Error('No se pudo leer el archivo.')); };
  reader.readAsArrayBuffer(file);
}

/* Detecta la fila de encabezados probando las primeras filas.
   Puntaje = cuántos alias reconoce. Por regla priorizamos la fila 2 (índice 1),
   pero si la 1ª fila resultara ser el encabezado real, la tomamos. */
function puntuarComoHeader(fila){
  var celdas = fila.map(sinAcentos);
  var vistos = {}, score = 0;
  Object.keys(ALIASES).forEach(function(campo){
    var found = celdas.some(function(c){
      return ALIASES[campo].some(function(a){ return c===a || (c.length>3 && c.indexOf(a)>=0); });
    });
    if(found && !vistos[campo]){ vistos[campo]=1; score++; }
  });
  return score;
}

function detectarHeaderRow(aoa){
  var limite = Math.min(aoa.length, 8);
  var mejor = { idx:1, score:-1 };
  for(var i=0;i<limite;i++){
    var s = puntuarComoHeader(aoa[i]||[]);
    // Preferimos fila 2 (índice 1): le damos un pequeño bonus
    if(i===1) s += 0.5;
    if(s > mejor.score){ mejor = { idx:i, score:s }; }
  }
  return mejor.idx;
}

/* Mapea cada campo lógico -> índice de columna real */
function mapearColumnas(headerRow){
  var celdas = headerRow.map(sinAcentos);
  var mapa = {};
  Object.keys(ALIASES).forEach(function(campo){
    var idx = -1, bestLen = 0;
    for(var c=0;c<celdas.length;c++){
      var cel = celdas[c];
      if(!cel) continue;
      for(var a=0;a<ALIASES[campo].length;a++){
        var al = ALIASES[campo][a];
        // coincidencia exacta primero
        if(cel===al){ idx=c; bestLen=999; break; }
        if(cel.indexOf(al)>=0 && al.length>bestLen){ idx=c; bestLen=al.length; }
      }
      if(bestLen===999) break;
    }
    mapa[campo] = idx;
  });
  return mapa;
}

/* Cuenta de columnas útiles (para saber cuándo una fila está «corrida») */
function anchoEsperado(headerRow){
  var w=0; for(var i=0;i<headerRow.length;i++){ if(norm(headerRow[i])!=='') w=i+1; } return w;
}

/* Una fila se considera «fragmento/continuación» si NO tiene documento ni bodega
   pero sí trae algo de contenido (p.ej. descripción partida en 2 filas). */
function esFragmento(fila, mapa){
  var doc = mapa.documento>=0 ? norm(fila[mapa.documento]) : '';
  var bod = mapa.bodega>=0 ? norm(fila[mapa.bodega]) : '';
  if(doc || bod) return false;
  return fila.some(function(c){ return norm(c)!==''; });
}

/* Reconstruye filas divididas: fusiona un fragmento con la fila previa,
   concatenando texto en las columnas de texto vacías. */
function fusionarFragmentos(rows, mapa){
  var out = [], dropFrag = 0;
  var textCols = {}; [mapa.descripcion,mapa.cie10,mapa.paciNombre].forEach(function(c){ if(c>=0) textCols[c]=1; });
  for(var i=0;i<rows.length;i++){
    var f = rows[i];
    if(out.length && esFragmento(f, mapa)){
      var prev = out[out.length-1];
      for(var c=0;c<f.length;c++){
        var val = norm(f[c]); if(!val) continue;
        if(textCols[c]){ prev[c] = norm(prev[c]) ? (norm(prev[c])+' '+val) : val; }
        else if(!norm(prev[c])){ prev[c] = f[c]; }
      }
      dropFrag++;
    }else{
      out.push(f.slice());
    }
  }
  return { rows: out, fragmentos: dropFrag };
}

/* Filtra filas incompletas: exige al menos documento O bodega + alguna cantidad/desc */
function esCompleta(fila, mapa){
  var doc = mapa.documento>=0 ? norm(fila[mapa.documento]) : '';
  var bod = mapa.bodega>=0 ? norm(fila[mapa.bodega]) : '';
  var desc = mapa.descripcion>=0 ? norm(fila[mapa.descripcion]) : '';
  var tieneCant = (mapa.entregado>=0 && norm(fila[mapa.entregado])!=='') ||
                  (mapa.formulado>=0 && norm(fila[mapa.formulado])!=='') ||
                  (mapa.pendiente>=0 && norm(fila[mapa.pendiente])!=='');
  if(!doc && !bod) return false;
  if(!desc && !tieneCant && mapa.codigo>=0 && !norm(fila[mapa.codigo])) return false;
  return true;
}

/* Pipeline completo: AoA cruda -> objetos de línea + reporte de limpieza */
function procesarAoA(aoa){
  var rep = { totalCrudas:0, headerRow:0, sinContenido:0, fragmentos:0, incompletas:0, usadas:0, columnas:{} };
  // 1) quitar filas totalmente vacías al inicio/dispersas
  var filas = aoa.filter(function(r){ return r && r.some(function(c){ return norm(c)!==''; }); });
  rep.totalCrudas = filas.length;
  if(filas.length < 2) throw new Error('El archivo no contiene datos suficientes.');

  // 2) detectar header (regla: fila 2, salvo evidencia clara de otra)
  var hIdx = detectarHeaderRow(filas);
  rep.headerRow = hIdx + 1; // 1-based para el usuario
  var headerRow = filas[hIdx];
  var mapa = mapearColumnas(headerRow);
  rep.columnas = mapa;

  // 3) datos = todo lo posterior al header
  var datos = filas.slice(hIdx+1);

  // 4) reconstruir fragmentos / filas corridas
  var fus = fusionarFragmentos(datos, mapa);
  rep.fragmentos = fus.fragmentos;

  // 5) eliminar incompletas
  var limpias = [];
  fus.rows.forEach(function(f){ if(esCompleta(f, mapa)) limpias.push(f); else rep.incompletas++; });
  rep.usadas = limpias.length;

  // 6) a objetos
  var objs = limpias.map(function(f){ return filaAObjeto(f, mapa); });
  return { objs: objs, rep: rep, mapa: mapa };
}

function cel(f, idx){ return idx>=0 ? f[idx] : ''; }
function filaAObjeto(f, m){
  return {
    documento:   norm(cel(f,m.documento)),
    bodega:      norm(cel(f,m.bodega)) || 'SIN BODEGA',
    estado:      up(cel(f,m.estado)),
    contrato:    up(cel(f,m.contrato)),
    epsRaw:      norm(cel(f,m.eps)),
    entregado:   toNum(cel(f,m.entregado)),
    formulado:   toNum(cel(f,m.formulado)),
    pendienteRaw:toNum(cel(f,m.pendiente)),
    soporteRaw:  norm(cel(f,m.soporte)),
    codigo:      norm(cel(f,m.codigo)),
    descripcion: norm(cel(f,m.descripcion)),
    cie10:       norm(cel(f,m.cie10)),
    codcie:      norm(cel(f,m.codcie)),
    paciente:    norm(cel(f,m.paciente)),
    paciNombre:  norm(cel(f,m.paciNombre)),
    usuarioCrea: norm(cel(f,m.usuarioCrea)) || 'SIN USUARIO',
    fechaRaw:    norm(cel(f,m.fecha))
  };
}

/* ================================================================
   ENRIQUECIMIENTO (nivel línea)
   ================================================================ */

/* Consolidación de EPS: agrupa variantes bajo un nombre común */
var EPS_MAP = [
  { re:/NUEVA\s*EPS|NUEVAEPS/,        g:'NUEVA EPS' },
  { re:/SANITAS|EPS\s*SANITAS/,       g:'SANITAS' },
  { re:/SURA|EPS\s*SURA/,             g:'SURA' },
  { re:/SALUD\s*TOTAL/,               g:'SALUD TOTAL' },
  { re:/COMPENSAR/,                    g:'COMPENSAR' },
  { re:/FAMISANAR/,                    g:'FAMISANAR' },
  { re:/COOSALUD/,                     g:'COOSALUD' },
  { re:/MUTUAL\s*SER|MUTUALSER/,      g:'MUTUAL SER' },
  { re:/CAJACOPI/,                     g:'CAJACOPI' },
  { re:/ASMET/,                        g:'ASMET SALUD' },
  { re:/EMSSANAR|EMSANAR/,            g:'EMSSANAR' },
  { re:/CAPITAL\s*SALUD/,             g:'CAPITAL SALUD' },
  { re:/CAPRESOCA/,                    g:'CAPRESOCA' },
  { re:/COMFA/,                        g:'COMFAMILIAR' },
  { re:/MALLAMAS|MALLAM/,             g:'MALLAMAS' }
];
function consolidarEps(raw){
  var s = sinAcentos(raw);
  if(!s) return 'SIN EPS';
  for(var i=0;i<EPS_MAP.length;i++){ if(EPS_MAP[i].re.test(s)) return EPS_MAP[i].g; }
  return up(raw);
}

/* Estado inactivo */
function esInactivo(estado){
  var s = sinAcentos(estado);
  return s.indexOf('INACTIV')>=0 || s==='ANULAD' || s.indexOf('ANULAD')>=0;
}

/* Contrato evento */
function esEvento(contrato){ return sinAcentos(contrato).indexOf('EVENTO')>=0; }

/* ¿Tiene soporte esta línea? */
function lineaTieneSoporte(o){
  var s = sinAcentos(o.soporteRaw);
  if(!s) return false;
  if(/^(NO|SIN|N\/A|NA|0|FALSE|-)$/.test(s)) return false;
  return true;
}

/* Pendiente efectivo por línea */
function pendienteLinea(o){
  if(o.pendienteRaw>0) return o.pendienteRaw;
  var d = o.formulado - o.entregado;
  return d>0 ? d : 0;
}

/* Mes a partir de la fecha (varios formatos) */
function extraerMes(raw){
  if(!raw) return '';
  var s = String(raw).trim();
  var m;
  if((m=s.match(/^(\d{4})[\-\/](\d{1,2})/))) return m[1]+'-'+('0'+m[2]).slice(-2);
  if((m=s.match(/^(\d{1,2})[\-\/](\d{1,2})[\-\/](\d{2,4})/))){
    var y=m[3].length===2?'20'+m[3]:m[3]; return y+'-'+('0'+m[2]).slice(-2);
  }
  // serial de Excel
  var n = parseFloat(s);
  if(!isNaN(n) && n>20000 && n<80000){
    var d = new Date(Math.round((n-25569)*86400*1000));
    return d.getUTCFullYear()+'-'+('0'+(d.getUTCMonth()+1)).slice(-2);
  }
  return '';
}

function claveDoc(o){ return o.documento+'\u241f'+o.bodega; }

function enriquecer(objs){
  return objs.map(function(o){
    o.eps = consolidarEps(o.epsRaw);
    o.inactivo = esInactivo(o.estado);
    o.activo = !o.inactivo;
    o.evento = esEvento(o.contrato);
    o.tieneSoporte = lineaTieneSoporte(o);
    o.pendiente = pendienteLinea(o);
    o.entregadaLinea = o.pendiente<=0 && (o.entregado>0 || o.formulado>0 || o.entregado===o.formulado);
    if(o.formulado>0) o.entregadaLinea = o.entregado>=o.formulado;
    o.mes = extraerMes(o.fechaRaw);
    o.clave = claveDoc(o);
    return o;
  });
}

/* ---------- Agrupaciones ---------- */
/* Aplica filtros globales a un conjunto de líneas */
function aplicarFiltros(lineas){
  return lineas.filter(function(o){
    if(FILTROS.bodega && sinAcentos(o.bodega).indexOf(sinAcentos(FILTROS.bodega))<0) return false;
    if(FILTROS.contrato && o.contrato!==FILTROS.contrato) return false;
    if(FILTROS.eps && o.eps!==FILTROS.eps) return false;
    if(FILTROS.mes && o.mes!==FILTROS.mes) return false;
    return true;
  });
}
function soloActivas(lineas){ return lineas.filter(function(o){ return o.activo; }); }

/* Agrupa líneas en dispensas (Documento+Bodega). Una dispensa entregada
   solo si TODAS sus líneas quedaron entregadas. */
function agruparDispensas(lineas){
  var map = {};
  lineas.forEach(function(o){
    var k = o.clave;
    if(!map[k]) map[k] = { clave:k, documento:o.documento, bodega:o.bodega, contrato:o.contrato,
      eps:o.eps, mes:o.mes, usuarioCrea:o.usuarioCrea, lineas:[], evento:o.evento, algunSoporte:false, fechaRaw:o.fechaRaw };
    map[k].lineas.push(o);
    if(o.tieneSoporte) map[k].algunSoporte = true;
    if(o.evento) map[k].evento = true;
  });
  return Object.keys(map).map(function(k){
    var d = map[k];
    d.entregada = d.lineas.every(function(l){ return l.entregadaLinea; });
    d.pendiente = !d.entregada;
    return d;
  });
}

/* ================================================================
   RENDER — helpers de UI
   ================================================================ */
function statCard(label,value,sub,warn,barPct){
  var bar = barPct!=null ? '<div class="bar"><i style="width:'+Math.min(100,barPct)+'%"></i></div>' : '';
  return '<div class="stat'+(warn?' warn':'')+'"><div class="label">'+label+'</div>'+
    '<div class="value">'+value+'</div>'+(sub?'<div class="sub">'+sub+'</div>':'')+bar+'</div>';
}

/* Donut SVG de 2 segmentos (bueno vs resto) */
function donut(svg, a, b, colA, colB, centerTxt){
  svg.innerHTML='';
  var total=a+b, cx=100, cy=100, r=70, sw=30;
  var ns='http://www.w3.org/2000/svg';
  function arc(frac, offset, color){
    var C=2*Math.PI*r;
    var c=document.createElementNS(ns,'circle');
    c.setAttribute('cx',cx); c.setAttribute('cy',cy); c.setAttribute('r',r);
    c.setAttribute('fill','none'); c.setAttribute('stroke',color); c.setAttribute('stroke-width',sw);
    c.setAttribute('stroke-dasharray',(C*frac)+' '+(C*(1-frac)));
    c.setAttribute('stroke-dashoffset',(-C*offset));
    c.setAttribute('transform','rotate(-90 '+cx+' '+cy+')');
    svg.appendChild(c);
  }
  if(total<=0){ arc(1,0,'#e2e8f0'); }
  else { arc(a/total,0,colA); arc(b/total,a/total,colB); }
  var t=document.createElementNS(ns,'text');
  t.setAttribute('x',cx); t.setAttribute('y',cy+9); t.setAttribute('text-anchor','middle');
  t.setAttribute('class','pie-center-label'); t.textContent=centerTxt;
  svg.appendChild(t);
}
function legend(box, items){
  box.innerHTML = items.map(function(it){
    return '<div class="item"><span class="sw" style="background:'+it.c+'"></span>'+it.l+
      '<span class="val">'+it.v+'</span></div>';
  }).join('');
}
function opts(sel, values, keep){
  var cur = keep?sel.value:'';
  var first = sel.options.length?sel.options[0].outerHTML:'';
  sel.innerHTML = first + values.map(function(v){ return '<option value="'+v+'">'+v+'</option>'; }).join('');
  if(keep) sel.value=cur;
}

/* ================================================================
   RENDER 1 — INDICADOR DE DISPENSA
   ================================================================ */
function renderDispensa(){
  var lineas = soloActivas(aplicarFiltros(RAW));
  var disp = agruparDispensas(lineas);
  var byBod = {};
  disp.forEach(function(d){
    if(!byBod[d.bodega]) byBod[d.bodega]={ bodega:d.bodega, total:0, entregadas:0 };
    byBod[d.bodega].total++;
    if(d.entregada) byBod[d.bodega].entregadas++;
  });
  var rows = Object.keys(byBod).map(function(k){ return byBod[k]; })
    .sort(function(a,b){ return b.total-a.total; });
  var totD=disp.length, totE=disp.filter(function(d){return d.entregada;}).length, totP=totD-totE;

  $('#statsDispensa').innerHTML =
    statCard('Dispensas activas', fmt(totD), 'únicas (doc+bodega)') +
    statCard('Entregadas', fmt(totE), pct1(totE,totD)+'% del total', false, pct(totE,totD)) +
    statCard('Pendientes', fmt(totP), pct1(totP,totD)+'% del total', totP>0) +
    statCard('Bodegas', fmt(rows.length),'con dispensas activas');

  var tb = $('#tblDispensa tbody'); tb.innerHTML='';
  rows.forEach(function(r){
    var p=r.total-r.entregadas, ef=pct1(r.entregadas,r.total), ip=pct1(p,r.total);
    tb.insertAdjacentHTML('beforeend','<tr><td class="txt wrapcell">'+r.bodega+'</td><td>'+fmt(r.total)+'</td><td>'+
      fmt(r.entregadas)+'</td><td>'+fmt(p)+'</td><td class="'+pctCls(ef)+'">'+ef+'%</td><td>'+ip+'%</td></tr>');
  });
  tb.insertAdjacentHTML('beforeend','<tr class="total-row"><td class="txt">TOTAL</td><td>'+fmt(totD)+'</td><td>'+
    fmt(totE)+'</td><td>'+fmt(totP)+'</td><td>'+pct1(totE,totD)+'%</td><td>'+pct1(totP,totD)+'%</td></tr>');

  // Pie por bodega
  var sel=$('#pieBodegaSelect');
  var vals=['(Todas)'].concat(rows.map(function(r){return r.bodega;}));
  sel.innerHTML = vals.map(function(v){return '<option>'+v+'</option>';}).join('');
  function pintarPie(){
    var e,p;
    if(sel.value==='(Todas)'){ e=totE; p=totP; }
    else { var r=byBod[sel.value]; e=r.entregadas; p=r.total-r.entregadas; }
    donut($('#pieDispensa'), e, p, '#10b981', '#dc2626', pct1(e,e+p)+'%');
    legend($('#pieLegend'), [
      { c:'#10b981', l:'Entregadas', v:fmt(e) },
      { c:'#dc2626', l:'Pendientes', v:fmt(p) }
    ]);
  }
  sel.onchange=pintarPie; pintarPie();
  DISP_CACHE = rows; DISP_TOT={totD:totD,totE:totE,totP:totP};
}
var DISP_CACHE=[], DISP_TOT={};

/* ================================================================
   RENDER 2 — INDICADOR SOPORTE EVENTO
   ================================================================ */
function renderSoporte(){
  var lineas = soloActivas(aplicarFiltros(RAW));
  var disp = agruparDispensas(lineas).filter(function(d){ return d.evento && d.entregada; });
  var byBod={};
  disp.forEach(function(d){
    if(!byBod[d.bodega]) byBod[d.bodega]={ bodega:d.bodega, total:0, con:0 };
    byBod[d.bodega].total++;
    if(d.algunSoporte) byBod[d.bodega].con++;
  });
  var rows=Object.keys(byBod).map(function(k){return byBod[k];}).sort(function(a,b){return b.total-a.total;});
  var totT=disp.length, totC=disp.filter(function(d){return d.algunSoporte;}).length, totS=totT-totC;

  $('#statsSoporte').innerHTML =
    statCard('Entregadas (Evento)', fmt(totT), 'dispensas activas de EVENTO') +
    statCard('Con soporte', fmt(totC), pct1(totC,totT)+'%', false, pct(totC,totT)) +
    statCard('Sin soporte', fmt(totS), pct1(totS,totT)+'%', totS>0) +
    statCard('Bodegas', fmt(rows.length),'con evento entregado');

  var tb=$('#tblSoporte tbody'); tb.innerHTML='';
  rows.forEach(function(r){
    var s=r.total-r.con, pc=pct1(r.con,r.total);
    tb.insertAdjacentHTML('beforeend','<tr><td class="txt wrapcell">'+r.bodega+'</td><td>'+fmt(r.total)+'</td><td>'+
      fmt(r.con)+'</td><td>'+fmt(s)+'</td><td class="'+pctCls(pc)+'">'+pc+'%</td></tr>');
  });
  tb.insertAdjacentHTML('beforeend','<tr class="total-row"><td class="txt">TOTAL</td><td>'+fmt(totT)+'</td><td>'+
    fmt(totC)+'</td><td>'+fmt(totS)+'</td><td>'+pct1(totC,totT)+'%</td></tr>');

  var sel=$('#pieSoporteSelect');
  var vals=['(Todas)'].concat(rows.map(function(r){return r.bodega;}));
  sel.innerHTML=vals.map(function(v){return '<option>'+v+'</option>';}).join('');
  function pintar(){
    var c,s;
    if(sel.value==='(Todas)'){ c=totC; s=totS; }
    else { var r=byBod[sel.value]; c=r.con; s=r.total-r.con; }
    donut($('#pieSoporte'), c, s, '#0b5fa5', '#d98a2b', pct1(c,c+s)+'%');
    legend($('#pieSoporteLegend'), [
      { c:'#0b5fa5', l:'Con soporte', v:fmt(c) },
      { c:'#d98a2b', l:'Sin soporte', v:fmt(s) }
    ]);
  }
  sel.onchange=pintar; pintar();
  SOP_CACHE=rows; SOP_TOT={totT:totT,totC:totC,totS:totS}; SOP_DISP=disp;
}
var SOP_CACHE=[], SOP_TOT={}, SOP_DISP=[];

/* ================================================================
   RENDER 3 — COHORTES
   ================================================================ */
var COHORTES_DEF = [
  { key:'VIH',     label:'VIH',           dxs:['VIH','HIV','INMUNODEFICIENCIA HUMANA','SINDROME DE INMUNODEFICIENCIA'],meds:['TENOFOVIR','EMTRICITABINA','DOLUTEGRAVIR','EFAVIRENZ','LAMIVUDINA','ABACAVIR','RALTEGRAVIR','DARUNAVIR','RITONAVIR','LOPINAVIR','ZIDOVUDINA']},
  { key:'DIABETES',label:'Diabetes',       dxs:['DIABET','MELLITUS','HIPERGLUCEMIA'],meds:['METFORM','INSULINA','GLIMEPIRIDA','GLIBENCLAMIDA','GLICLAZIDA','EMPAGLIFLOZINA','SITAGLIPTINA','VILDAGLIPTINA','LINAGLIPTINA','CANAGLIFLOZINA','DAPAGLIFLOZINA']},
  { key:'HTA',     label:'Hipertensión',   dxs:['HIPERTENS','ENFERMEDAD HIPERTENSIVA'],meds:['LOSARTAN','ENALAPRIL','CAPTOPRIL','VALSARTAN','AMLODIPINO','HIDROCLOROTIAZIDA','METOPROLOL','CARVEDILOL','TELMISARTAN','RAMIPRIL']},
  { key:'ASMA',    label:'Asma / EPOC',    dxs:['ASMA','EPOC','ENFERMEDAD PULMONAR OBSTRUCTIVA','BRONQUITIS CRONICA','BRONCOESPASMO'],meds:['SALMETEROL','FLUTICASONA','BUDESONIDA','IPRATROPIO','TIOTROPIO','MONTELUKAST','SALBUTAMOL']},
  { key:'EPILEPSIA',label:'Epilepsia',     dxs:['EPILEPS','CONVULSIV','CRISIS EPIL'],meds:['CARBAMAZEPINA','FENITOINA','VALPROATO','LEVETIRACETAM','LAMOTRIGINA','TOPIRAMATO','GABAPENTINA','CLONAZEPAM']},
  { key:'CANCER',  label:'Oncología',      dxs:['CANCER','NEOPLAS','TUMOR','MALIGN','ONCOLOG','LEUCEMIA','LINFOMA','CARCINOMA','MELANOMA'],meds:['CICLOFOSFAMIDA','METOTREXATO','CISPLATINO','DOXORUBICINA','PACLITAXEL','RITUXIMAB','IMATINIB','TRASTUZUMAB','BEVACIZUMAB']},
  { key:'MENTAL',  label:'Salud mental',   dxs:['DEPRESI','ANSIEDAD','BIPOLAR','ESQUIZOFREN','PSICOSIS','TRASTORNO AFECTIVO','EPISODIO DEPRESIVO'],meds:['SERTRALINA','FLUOXETINA','CITALOPRAM','ESCITALOPRAM','PAROXETINA','VENLAFAXINA','DULOXETINA','QUETIAPINA','OLANZAPINA','RISPERIDONA','ARIPIPRAZOL','CLOZAPINA','HALOPERIDOL']},
  { key:'RENAL',   label:'Enfermedad renal',dxs:['ENFERMEDAD RENAL','INSUFICIENCIA RENAL','NEFROPATIA','DIALISIS','HEMODIALISIS','FALLA RENAL'],meds:['ERITROPOYETINA','DARBEPOETINA','SEVELAMER']},
  { key:'TIROIDES',label:'Tiroides',       dxs:['TIROID','HIPOTIROID','HIPERTIROID','BOCIO','HASHIMOTO'],meds:['LEVOTIROXINA','METIMAZOL']},
  { key:'ARTRITIS',label:'Artritis / Autoinmune',dxs:['ARTRITIS','LUPUS','AUTOINMUNE','REUMATO','ESCLERODER','POLIARTRITIS'],meds:['AZATIOPRINA','LEFLUNOMIDA','SULFASALAZINA','HIDROXICLOROQUINA','ADALIMUMAB','ETANERCEPT','INFLIXIMAB','TOCILIZUMAB']},
  { key:'CARDIO',  label:'Cardiovascular', dxs:['CARDIO','MIOCARD','INSUFICIENCIA CARDIACA','CORONARI','INFARTO','ANGINA','ARRITMIA','FIBRILACION AURICULAR'],meds:['ATORVASTATINA','ROSUVASTATINA','CLOPIDOGREL','ENOXAPARINA','WARFARINA','DIGOXINA','AMIODARONA']},
  { key:'HEPATIC', label:'Hepático',       dxs:['HEPAT','CIRROSIS','HEPATITIS','HIGADO GRASO','ESTEATOSIS HEPATICA'],meds:['ENTECAVIR','RIBAVIRINA','SOFOSBUVIR','URSODIOL']},
  { key:'OFTALMO', label:'Oftalmología',   dxs:['OFTALM','GLAUCOMA','CATARATA','RETIN','MACUL','CONJUNTIV'],meds:['LATANOPROST','TIMOLOL','DORZOLAMIDA','BRINZOLAMIDA','TRAVOPROST','BIMATOPROST']}
];

function perteneceCohorte(linea, def){
  var cie = sinAcentos(linea.cie10);
  var des = sinAcentos(linea.descripcion);
  if(cie && def.dxs.some(function(d){ return cie.indexOf(d)>=0; })) return true;
  if(des && def.meds.some(function(m){ return des.indexOf(m)>=0; })) return true;
  return false;
}

function renderCohortes(){
  var lineas = soloActivas(aplicarFiltros(RAW));
  var cohortes = COHORTES_DEF.map(function(def){
    var match = lineas.filter(function(l){ return perteneceCohorte(l,def); });
    var pacSet={},codSet={},bodSet={},ent=0,pen=0,uni=0;
    match.forEach(function(l){
      pacSet[l.paciente]=1; codSet[l.codigo]=1; bodSet[l.bodega]=1;
      uni+=l.entregado; pen+=l.pendiente; if(l.entregadaLinea) ent++;
    });
    return { key:def.key, label:def.label, lineas:match.length, pacients:Object.keys(pacSet).length,
      entregadas:ent, pendientes:pen, pctC:pct1(ent,match.length),
      unidades:Math.round(uni), codigos:Object.keys(codSet).length, bodegas:Object.keys(bodSet).length };
  }).filter(function(c){ return c.lineas>0; });

  var totL=0,totP=0,totE=0,totPn=0,totU=0;
  cohortes.forEach(function(c){ totL+=c.lineas; totP+=c.pacients; totE+=c.entregadas; totPn+=c.pendientes; totU+=c.unidades; });

  $('#statsCohortes').innerHTML =
    statCard('Cohortes activas', fmt(cohortes.length), 'con al menos 1 línea') +
    statCard('Líneas en cohortes', fmt(totL), fmt(totP)+' pacientes') +
    statCard('Entregadas', fmt(totE), pct1(totE,totL)+'%') +
    statCard('Unidades', fmt(totU), 'entregadas total');

  var tb=$('#tblCohortes tbody'); tb.innerHTML='';
  cohortes.forEach(function(c){
    tb.insertAdjacentHTML('beforeend','<tr><td class="txt wrapcell">'+c.label+'</td><td>'+fmt(c.pacients)+'</td><td>'+fmt(c.lineas)+'</td><td>'+
      fmt(c.entregadas)+'</td><td>'+fmt(c.pendientes)+'</td><td class="'+pctCls(c.pctC)+'">'+c.pctC+'%</td><td>'+fmt(c.unidades)+'</td><td>'+fmt(c.codigos)+'</td><td>'+c.bodegas+'</td></tr>');
  });
  if(cohortes.length){
    tb.insertAdjacentHTML('beforeend','<tr class="total-row"><td class="txt">TOTAL</td><td>-</td><td>'+fmt(totL)+'</td><td>'+
      fmt(totE)+'</td><td>'+fmt(totPn)+'</td><td>'+pct1(totE,totL)+'%</td><td>'+fmt(totU)+'</td><td>-</td><td>-</td></tr>');
  }

  opts($('#fCohorte'), cohortes.map(function(c){return c.label;}), true);
  var bods=[]; lineas.forEach(function(l){ if(bods.indexOf(l.bodega)<0) bods.push(l.bodega); });
  opts($('#fCohorteBodega'), bods.sort(), true);

  COH_DATA = { lineas:lineas, cohortes:cohortes, defs:COHORTES_DEF };
  renderCohortesTop();

  var diagBox=$('#cohortesDiag');
  var unMatched = lineas.filter(function(l){ return !COHORTES_DEF.some(function(d){ return perteneceCohorte(l,d); }); });
  if(unMatched.length>0){
    diagBox.style.display='block';
    diagBox.innerHTML='<b>'+fmt(unMatched.length)+'</b> líneas activas no clasificaron en ninguna cohorte ('+pct1(unMatched.length,lineas.length)+'% del filtrado).';
  } else { diagBox.style.display='none'; }
}
var COH_DATA={};

function renderCohortesTop(){
  var cLabel=$('#fCohorte').value, bLabel=$('#fCohorteBodega').value;
  var def = COH_DATA.defs.filter(function(d){return d.label===cLabel;})[0] || null;
  var lineas = COH_DATA.lineas;
  if(def) lineas = lineas.filter(function(l){ return perteneceCohorte(l,def); });
  if(bLabel) lineas = lineas.filter(function(l){ return l.bodega===bLabel; });
  var codMap={};
  lineas.forEach(function(l){
    if(!codMap[l.codigo]) codMap[l.codigo]={ codigo:l.codigo||'(sin código)', desc:l.descripcion, lineas:0, pacs:{}, pen:0 };
    codMap[l.codigo].lineas++; codMap[l.codigo].pacs[l.paciente]=1; codMap[l.codigo].pen+=l.pendiente;
  });
  var top=Object.keys(codMap).map(function(k){ var c=codMap[k]; c.pacients=Object.keys(c.pacs).length; return c; })
    .sort(function(a,b){return b.lineas-a.lineas;}).slice(0,10);
  var tb=$('#tblCohortesTop tbody'); tb.innerHTML='';
  top.forEach(function(t,i){
    tb.insertAdjacentHTML('beforeend','<tr><td>'+(i+1)+'</td><td class="txt">'+t.codigo+'</td><td class="txt wrapcell">'+t.desc+'</td><td>'+fmt(t.lineas)+'</td><td>'+fmt(t.pacients)+'</td><td>'+fmt(t.pen)+'</td></tr>');
  });
}

/* ================================================================
   RENDER 4 — DISPENSAS INACTIVAS
   ================================================================ */
function renderInactivas(){
  var lineasAll = aplicarFiltros(RAW).filter(function(o){ return o.inactivo; });
  var dispAll = agruparDispensas(lineasAll);
  // lista de usuarios (del conjunto completo) para el select
  var usuariosAll = {}; dispAll.forEach(function(d){ usuariosAll[d.usuarioCrea]=1; });
  opts($('#fInactivasUsuario'), Object.keys(usuariosAll).sort(), true);
  var fu = $('#fInactivasUsuario').value;
  var disp = fu ? dispAll.filter(function(d){ return d.usuarioCrea===fu; }) : dispAll;
  var lineas = fu ? lineasAll.filter(function(o){ return o.usuarioCrea===fu; }) : lineasAll;

  var byUsu={}, byBod={};
  disp.forEach(function(d){
    if(!byUsu[d.usuarioCrea]) byUsu[d.usuarioCrea]={usuario:d.usuarioCrea,total:0};
    byUsu[d.usuarioCrea].total++;
    if(!byBod[d.bodega]) byBod[d.bodega]={bodega:d.bodega,total:0};
    byBod[d.bodega].total++;
  });
  var rowsU=Object.keys(byUsu).map(function(k){return byUsu[k];}).sort(function(a,b){return b.total-a.total;});
  var rowsB=Object.keys(byBod).map(function(k){return byBod[k];}).sort(function(a,b){return b.total-a.total;});
  var tot=disp.length;

  $('#statsInactivas').innerHTML =
    statCard('Dispensas inactivas', fmt(tot), 'únicas (doc+bodega)', tot>0) +
    statCard('Líneas inactivas', fmt(lineas.length), 'filas del reporte') +
    statCard('Usuarios', fmt(rowsU.length), 'con dispensas inactivas') +
    statCard('Bodegas', fmt(rowsB.length), 'afectadas');

  var tbU=$('#tblInactivasUsuario tbody'); tbU.innerHTML='';
  rowsU.forEach(function(r){ tbU.insertAdjacentHTML('beforeend','<tr><td class="txt">'+r.usuario+'</td><td>'+fmt(r.total)+'</td><td>'+pct1(r.total,tot)+'%</td></tr>'); });
  var tbB=$('#tblInactivasBodega tbody'); tbB.innerHTML='';
  rowsB.forEach(function(r){ tbB.insertAdjacentHTML('beforeend','<tr><td class="txt wrapcell">'+r.bodega+'</td><td>'+fmt(r.total)+'</td><td>'+pct1(r.total,tot)+'%</td></tr>'); });

  INAC_CACHE = { disp:disp, lineas:lineas };

  var diagBox=$('#inactivasDiag');
  if(tot>0){ diagBox.style.display='block'; diagBox.innerHTML='Usuarios con más inactivas: <b>'+rowsU.slice(0,3).map(function(r){return r.usuario+' ('+r.total+')';}).join(', ')+'</b>'; }
  else { diagBox.style.display='none'; }
}
var INAC_CACHE={};

/* ================================================================
   EXPORTACIÓN A EXCEL (XLSX en el navegador)
   ================================================================ */
function exportar(nombre, sheets){
  var wb = XLSX.utils.book_new();
  sheets.forEach(function(s){
    var ws = XLSX.utils.aoa_to_sheet(s.data);
    XLSX.utils.book_append_sheet(wb, ws, s.name.slice(0,31));
  });
  XLSX.writeFile(wb, nombre);
  toast('Archivo generado: '+nombre);
}

function expDispensa(){
  var head=['Bodega','Dispensas','Entregadas','Pendientes','Eficiencia %','Índice pendiente %'];
  var rows=DISP_CACHE.map(function(r){ var p=r.total-r.entregadas;
    return [r.bodega,r.total,r.entregadas,p,pct1(r.entregadas,r.total),pct1(p,r.total)]; });
  rows.push(['TOTAL',DISP_TOT.totD,DISP_TOT.totE,DISP_TOT.totP,pct1(DISP_TOT.totE,DISP_TOT.totD),pct1(DISP_TOT.totP,DISP_TOT.totD)]);
  exportar('indicador_dispensa.xlsx',[{name:'Dispensa',data:[head].concat(rows)}]);
}

function expSoporte(){
  var head=['Bodega','Entregadas Evento','Con soporte','Sin soporte','% Con soporte'];
  var rows=SOP_CACHE.map(function(r){ var s=r.total-r.con; return [r.bodega,r.total,r.con,s,pct1(r.con,r.total)]; });
  rows.push(['TOTAL',SOP_TOT.totT,SOP_TOT.totC,SOP_TOT.totS,pct1(SOP_TOT.totC,SOP_TOT.totT)]);
  var det=[['Documento','Bodega','Contrato','EPS','¿Con soporte?','Líneas','Fecha']];
  SOP_DISP.forEach(function(d){ det.push([d.documento,d.bodega,d.contrato,d.eps,d.algunSoporte?'SÍ':'NO',d.lineas.length,d.fechaRaw]); });
  exportar('indicador_soporte_evento.xlsx',[{name:'Resumen',data:[head].concat(rows)},{name:'Detalle',data:det}]);
}

function expCohortes(){
  var head=['Cohorte','Pacientes','Líneas','Entregadas','Pendientes','% Cumpl.','Unidades','Códigos','Bodegas'];
  var rows=(COH_DATA.cohortes||[]).map(function(c){ return [c.label,c.pacients,c.lineas,c.entregadas,c.pendientes,c.pctC,c.unidades,c.codigos,c.bodegas]; });
  var det=[['Cohorte','Documento','Bodega','Paciente','Código','Descripción','CIE10','Entregado','Pendiente','Entregada']];
  (COH_DATA.lineas||[]).forEach(function(l){
    COHORTES_DEF.forEach(function(def){
      if(perteneceCohorte(l,def)) det.push([def.label,l.documento,l.bodega,l.paciente,l.codigo,l.descripcion,l.cie10,l.entregado,l.pendiente,l.entregadaLinea?'SÍ':'NO']);
    });
  });
  exportar('informe_cohortes.xlsx',[{name:'Resumen',data:[head].concat(rows)},{name:'Detalle',data:det}]);
}

function expInactivas(){
  var det=[['Documento','Bodega','Usuario creación','Contrato','EPS','Líneas','Fecha']];
  (INAC_CACHE.disp||[]).forEach(function(d){ det.push([d.documento,d.bodega,d.usuarioCrea,d.contrato,d.eps,d.lineas.length,d.fechaRaw]); });
  var lin=[['Documento','Bodega','Usuario','Estado','Código','Descripción','Entregado','Pendiente','Fecha']];
  (INAC_CACHE.lineas||[]).forEach(function(l){ lin.push([l.documento,l.bodega,l.usuarioCrea,l.estado,l.codigo,l.descripcion,l.entregado,l.pendiente,l.fechaRaw]); });
  exportar('dispensas_inactivas.xlsx',[{name:'Dispensas',data:det},{name:'Lineas',data:lin}]);
}

/* ================================================================
   ORQUESTACIÓN / EVENTOS
   ================================================================ */
var TAB_ACTUAL='dispensa';

function renderTodo(){
  renderDispensa();
  renderSoporte();
  renderCohortes();
  renderInactivas();
}

function poblarFiltrosGlobales(){
  var contratos={}, epss={}, meses={};
  RAW.forEach(function(o){
    if(o.contrato) contratos[o.contrato]=1;
    if(o.eps) epss[o.eps]=1;
    if(o.mes) meses[o.mes]=1;
  });
  function fill(sel, arr){ sel.innerHTML='<option value="">'+sel.options[0].text+'</option>'+
    arr.sort().map(function(v){return '<option value="'+v+'">'+v+'</option>';}).join(''); }
  fill($('#fContrato'), Object.keys(contratos));
  fill($('#fEps'), Object.keys(epss));
  fill($('#fMes'), Object.keys(meses));
}

function mostrarReporteLimpieza(rep, mapa){
  var faltan=[];
  ['documento','bodega','estado','contrato','descripcion','entregado'].forEach(function(k){
    if(mapa[k]<0) faltan.push(k);
  });
  var box=$('#cleanReport');
  box.style.display='block';
  box.innerHTML='<div class="cr-head">✓ Archivo procesado correctamente</div><ul>'+
    '<li>Filas leídas del archivo: <b>'+fmt(rep.totalCrudas)+'</b></li>'+
    '<li>Encabezados detectados en la fila <b>'+rep.headerRow+'</b> (la fila 1 de metadata se omitió)</li>'+
    '<li>Filas divididas reconstruidas: <b>'+fmt(rep.fragmentos)+'</b></li>'+
    '<li>Filas incompletas eliminadas: <b>'+fmt(rep.incompletas)+'</b></li>'+
    '<li>Filas de datos válidas usadas: <b>'+fmt(rep.usadas)+'</b></li>'+
    (faltan.length?'<li style="color:#b23a34">⚠ No se reconocieron columnas: <b>'+faltan.join(', ')+'</b> (revisa los encabezados del archivo)</li>':'')+
    '</ul>';
}

function cargarArchivo(file){
  if(!file) return;
  var ext=(file.name.split('.').pop()||'').toLowerCase();
  if(['xlsx','xls','csv'].indexOf(ext)<0){ toast('Formato no admitido. Usa .xlsx, .xls o .csv',true); return; }
  $('#dbStatusText').textContent='Procesando '+file.name+'…';
  leerArchivo(file, function(err, aoa){
    if(err){ toast('Error al leer el archivo: '+err.message,true); $('#dbStatusText').textContent='Error de lectura'; return; }
    try{
      var res = procesarAoA(aoa);
      if(!res.objs.length){ toast('No se encontraron filas válidas en el archivo.',true); return; }
      RAW = enriquecer(res.objs);
      META.archivo=file.name; META.fecha=new Date().toLocaleString('es-CO');
      META.filasCrudas=res.rep.totalCrudas; META.filasUsadas=res.rep.usadas;

      mostrarReporteLimpieza(res.rep, res.mapa);
      poblarFiltrosGlobales();
      FILTROS={bodega:'',contrato:'',eps:'',mes:''};
      $('#fBodega').value=''; $('#fContrato').value=''; $('#fEps').value=''; $('#fMes').value='';
      renderTodo();

      $('#filtersCard').style.display='block';
      $('#viewerCard').style.display='block';
      $('#dbDot').className='dot on';
      $('#dbStatusText').textContent=fmt(RAW.length)+' líneas · '+file.name;
      $('#fechaDatos').textContent='Cargado: '+META.fecha;
      toast('Datos cargados: '+fmt(RAW.length)+' líneas');
    }catch(e){ toast('Error al procesar: '+e.message,true); $('#dbStatusText').textContent='Error de proceso'; }
  });
}

function initEventos(){
  var dz=$('#dropzone'), fi=$('#fileInput');
  $('#btnBrowse').addEventListener('click', function(e){ e.stopPropagation(); fi.click(); });
  dz.addEventListener('click', function(){ fi.click(); });
  fi.addEventListener('change', function(){ if(fi.files[0]) cargarArchivo(fi.files[0]); fi.value=''; });
  ['dragenter','dragover'].forEach(function(ev){ dz.addEventListener(ev,function(e){ e.preventDefault(); dz.classList.add('drag'); }); });
  ['dragleave','drop'].forEach(function(ev){ dz.addEventListener(ev,function(e){ e.preventDefault(); dz.classList.remove('drag'); }); });
  dz.addEventListener('drop', function(e){ var f=e.dataTransfer.files[0]; if(f) cargarArchivo(f); });

  // Tabs
  $$('.result-tabs button').forEach(function(b){
    b.addEventListener('click', function(){
      $$('.result-tabs button').forEach(function(x){x.classList.remove('active');});
      b.classList.add('active');
      TAB_ACTUAL=b.getAttribute('data-sub');
      $$('.subview').forEach(function(s){ s.classList.remove('active'); });
      $('#sub-'+TAB_ACTUAL).classList.add('active');
    });
  });

  // Filtros globales
  function onFiltro(){ if(!RAW.length) return;
    FILTROS.bodega=$('#fBodega').value; FILTROS.contrato=$('#fContrato').value;
    FILTROS.eps=$('#fEps').value; FILTROS.mes=$('#fMes').value; renderTodo(); }
  $('#fBodega').addEventListener('input', onFiltro);
  ['#fContrato','#fEps','#fMes'].forEach(function(s){ $(s).addEventListener('change', onFiltro); });
  $('#btnLimpiar').addEventListener('click', function(){
    $('#fBodega').value=''; $('#fContrato').value=''; $('#fEps').value=''; $('#fMes').value='';
    FILTROS={bodega:'',contrato:'',eps:'',mes:''}; renderTodo();
  });

  // Filtros cohortes / inactivas
  $('#fCohorte').addEventListener('change', renderCohortesTop);
  $('#fCohorteBodega').addEventListener('change', renderCohortesTop);
  $('#fInactivasUsuario').addEventListener('change', renderInactivas);

  // Exportaciones
  $('#btnExpDispensa').addEventListener('click', function(){ if(RAW.length) expDispensa(); });
  $('#btnExpSoporte').addEventListener('click', function(){ if(RAW.length) expSoporte(); });
  $('#btnExpCohortes').addEventListener('click', function(){ if(RAW.length) expCohortes(); });
  $('#btnExpInactivas').addEventListener('click', function(){ if(RAW.length) expInactivas(); });
}

document.addEventListener('DOMContentLoaded', initEventos);

