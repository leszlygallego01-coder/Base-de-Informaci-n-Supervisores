/* ================================================================
   Reporte de Dispensación — app local unificada
   Fuente única: archivo "Reporte de Dispensación" (.xlsx / .csv)
   100% en el navegador. Sin nube.
   ================================================================ */
'use strict';

/* ---------- Estado global ---------- */
var RAW = [];        // filas enriquecidas (nivel línea)
var META = { archivo:'', fecha:'', filasCrudas:0, filasUsadas:0 };
var FILTROS = { bodega:'', eps:'', depto:'', zona:'', contrato:'' };

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
  tipoComp:       ['TIPO COMPROBANTE','TIPO DE COMPROBANTE','COMPROBANTE'],
  documento:      ['DOCUMENTO','NUMERO DOCUMENTO','NRO DOCUMENTO','NUM DOCUMENTO','NO DOCUMENTO','NUMERO DE DOCUMENTO','CONSECUTIVO'],
  bodega:         ['BODEGA DETALLE','BODEGA DE DETALLE','DETALLE BODEGA'],
  bodegaAlt:      ['SUCURSAL','BODEGA','PUNTO','PUNTO DISPENSACION','SEDE','FARMACIA'],
  estado:         ['ESTADO','ESTADO DISPENSACION','ESTADO REGISTRO'],
  contrato:       ['CONTRATO','TIPO CONTRATO','MODALIDAD','REGIMEN CONTRATO'],
  sigla:          ['SIGLA COMERCIAL','SIGLA','SIGLA EPS','SIGLA ENTIDAD'],
  eps:            ['EPS','ENTIDAD','ASEGURADOR','ASEGURADORA','ENTIDAD RESPONSABLE','PAGADOR'],
  diferencia:     ['DIFERENCIA','DIF','SALDO DIFERENCIA'],
  entregado:      ['CANTIDAD ENTREGADA','CANT ENTREGADA','ENTREGADO','UNIDADES ENTREGADAS','CANTIDAD DISPENSADA','DISPENSADO'],
  formulado:      ['CANTIDAD FORMULADA','CANT FORMULADA','FORMULADO','UNIDADES FORMULADAS','CANTIDAD ORDENADA','PRESCRITO','CANTIDAD PRESCRITA'],
  soporte:        ['SOPORTE','TIENE SOPORTE','SOPORTE EVENTO','NRO SOPORTE','NUMERO SOPORTE','NO SOPORTE','NUMERO SOPORTE EVENTO'],
  codigo:         ['CODIGO','CODIGO MEDICAMENTO','COD MEDICAMENTO','CUM','COD PRODUCTO','CODIGO PRODUCTO','COD','CODIGO ARTICULO'],
  descripcion:    ['DESCRIPCION','DESCRIPCION MEDICAMENTO','MEDICAMENTO','PRODUCTO','DESCRIPCION PRODUCTO','NOMBRE MEDICAMENTO','ARTICULO','DESCRIPCION ARTICULO'],
  cohorte:        ['DESCRIPCION COHORTE','COHORTE','DESC COHORTE','NOMBRE COHORTE','GRUPO COHORTE','PROGRAMA'],
  cie10:          ['DESCRIPCION CIE 10','DESCRIPCION CIE10','DESC CIE 10','DIAGNOSTICO CIE 10','DESCRIPCION DIAGNOSTICO','DIAGNOSTICO','CIE 10','CIE10'],
  paciente:       ['DOCUMENTO PACIENTE','IDENTIFICACION','IDENTIFICACION PACIENTE','NRO IDENTIFICACION','CEDULA','DOC PACIENTE','ID PACIENTE','NUMERO IDENTIFICACION','DOCUMENTO AFILIADO'],
  usuarioCrea:    ['USUARIO CREACION','USUARIO CREA','USUARIO','USUARIO DISPENSA','CREADO POR','DISPENSADO POR','USUARIO REGISTRO'],
  fecha:          ['FECHA','FECHA DISPENSACION','FECHA DISPENSA','FECHA CREACION','FECHA ENTREGA','FECHA REGISTRO','FECHA DOCUMENTO','FECHA COMPROBANTE']
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
  var textCols = {}; [mapa.descripcion,mapa.cohorte].forEach(function(c){ if(c>=0) textCols[c]=1; });
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
                  (mapa.diferencia>=0 && norm(fila[mapa.diferencia])!=='');
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
  // Bodega: usar EXCLUSIVAMENTE "Bodega detalle". Si no existe, caer a otra columna de bodega.
  if(mapa.bodega < 0 && mapa.bodegaAlt >= 0) mapa.bodega = mapa.bodegaAlt;
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
    tipoComp:    norm(cel(f,m.tipoComp)),
    documento:   norm(cel(f,m.documento)),
    bodega:      norm(cel(f,m.bodega)) || 'SIN BODEGA',
    estado:      up(cel(f,m.estado)),
    contrato:    up(cel(f,m.contrato)),
    sigla:       norm(cel(f,m.sigla)),
    epsRaw:      norm(cel(f,m.eps)),
    diferenciaRaw: norm(cel(f,m.diferencia)),
    entregado:   toNum(cel(f,m.entregado)),
    formulado:   toNum(cel(f,m.formulado)),
    soporteRaw:  norm(cel(f,m.soporte)),
    codigo:      norm(cel(f,m.codigo)),
    descripcion: norm(cel(f,m.descripcion)),
    cohorteRaw:  norm(cel(f,m.cohorte)),
    cie10Raw:    norm(cel(f,m.cie10)),
    paciente:    norm(cel(f,m.paciente)),
    usuarioCrea: norm(cel(f,m.usuarioCrea)) || 'SIN USUARIO',
    fechaRaw:    norm(cel(f,m.fecha))
  };
}

/* ================================================================
   ENRIQUECIMIENTO (nivel línea)
   ================================================================ */

/* Consolidación de EPS a partir de la columna SIGLA COMERCIAL (o sigla / EPS).
   Reglas del negocio (prioridad de arriba hacia abajo). */
var EPS_MAP = [
  { re:/NUEVA\s*EMPRESA\s*PROMOTORA|NUEVA\s*EPS|NUEVAEPS/,  g:'NUEVA EPS' },
  { re:/ASMET/,                                             g:'ASMET SALUD' },
  { re:/COOSALUD/,                                           g:'COOSALUD' },
  { re:/SANITAS/,                                            g:'SANITAS' },
  { re:/FAMILIAR\s*DE\s*COLOMBIA|EPS\s*FAMILIAR|^FAMILIAR/, g:'FAMILIAR' },
  { re:/FAMISANAR/,                                          g:'FAMISANAR' }
];
/* Consolida usando primero la SIGLA COMERCIAL; si no hay, usa EPS cruda. */
function consolidarEps(o){
  var base = o.sigla || o.epsRaw;
  var s = sinAcentos(base);
  if(!s) return 'SIN EPS';
  for(var i=0;i<EPS_MAP.length;i++){ if(EPS_MAP[i].re.test(s)) return EPS_MAP[i].g; }
  return 'SIN EPS';
}

/* ---------- Homologación geográfica por BODEGA / SUCURSAL ----------
   Cruza el código (p.ej. "M102") contra Zona y Departamento.
   Amplia libremente este diccionario con tus códigos reales. */
var ZONA_MAP = {
  'M15':  { zona:'ZONA TOLIMA',       depto:'TOLIMA' },
  'M102': { zona:'ZONA NARI\u00d1O',       depto:'NARI\u00d1O' },
  'M107': { zona:'ZONA EJE CAFETERO', depto:'RISARALDA' },
  'M108': { zona:'ZONA EJE CAFETERO', depto:'CALDAS' },
  'M109': { zona:'ZONA EJE CAFETERO', depto:'QUINDIO' },
  'M110': { zona:'ZONA VALLE',        depto:'VALLE DEL CAUCA' },
  'M111': { zona:'ZONA CAUCA',        depto:'CAUCA' },
  'M112': { zona:'ZONA HUILA',        depto:'HUILA' },
  'M113': { zona:'ZONA CUNDINAMARCA', depto:'CUNDINAMARCA' },
  'M114': { zona:'ZONA ANTIOQUIA',    depto:'ANTIOQUIA' }
};
/* Extrae el código de bodega (algo tipo M seguido de dígitos) del texto. */
function codigoBodega(bodega){
  var s = up(bodega);
  var m = s.match(/\bM\s*-?\s*(\d{1,4})\b/);
  if(m) return 'M'+m[1];
  m = s.match(/^\s*(\d{1,4})\b/);
  if(m) return 'M'+m[1];
  return '';
}
function homologarGeo(bodega){
  var cod = codigoBodega(bodega);
  if(cod && ZONA_MAP[cod]) return { zona:ZONA_MAP[cod].zona, depto:ZONA_MAP[cod].depto, codigo:cod };
  return { zona:'SIN ZONA', depto:'SIN DEPARTAMENTO', codigo:cod };
}

/* Estado inactivo */
function esInactivo(estado){
  var s = sinAcentos(estado);
  return s.indexOf('INACTIV')>=0 || s.indexOf('ANULAD')>=0;
}

/* Contrato evento vs cápita */
function esEvento(contrato){ return sinAcentos(contrato).indexOf('EVENTO')>=0; }
function esCapita(contrato){ var s=sinAcentos(contrato); return s.indexOf('CAPITA')>=0 || s.indexOf('CAPITACION')>=0; }

/* ¿Tiene soporte esta línea? */
function lineaTieneSoporte(o){
  var s = sinAcentos(o.soporteRaw);
  if(!s) return false;
  if(/^(NO|SIN|N\/A|NA|0|FALSE|-|0\.0)$/.test(s)) return false;
  return true;
}

/* Regla global: una línea está ENTREGADA si Diferencia === 0.
   Diferencia < 0 => pendiente. Se convierte estrictamente a número. */
function difLinea(o){
  if(o.diferenciaRaw!==''){ return toNum(o.diferenciaRaw); }
  // respaldo: entregado - formulado (negativo = pendiente)
  return o.entregado - o.formulado;
}

/* Fecha -> Date (para rango) y mes YYYY-MM */
function parseFecha(raw){
  if(!raw) return null;
  var s = String(raw).trim(); var m;
  if((m=s.match(/^(\d{4})[\-\/](\d{1,2})[\-\/](\d{1,2})/))) return new Date(+m[1],+m[2]-1,+m[3]);
  if((m=s.match(/^(\d{1,2})[\-\/](\d{1,2})[\-\/](\d{2,4})/))){ var y=m[3].length===2?2000+ +m[3]:+m[3]; return new Date(y,+m[2]-1,+m[1]); }
  var n=parseFloat(s);
  if(!isNaN(n) && n>20000 && n<80000){ return new Date(Math.round((n-25569)*86400*1000)); }
  return null;
}
function fmtFecha(d){ if(!d) return '-'; return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2); }

/* ---------- Cohortes por DESCRIPCION CIE 10 ----------
   Consolida los diagnósticos en grandes grupos de salud.
   Orden = prioridad (primer match gana). Amplia libremente. */
var COHORTE_CIE = [
  { grupo:'Trasplantados',       kw:['TRASPLANTE','TRANSPLANTE','INJERTO DE ORGANO','RECEPTOR DE ORGANO'] },
  { grupo:'Maternas',            kw:['EMBARAZO','PARTO','GESTACION','GESTANTE','PUERPERIO','PRENATAL','OBSTETRIC','ABORTO','CESAREA'] },
  { grupo:'Diábeticos',          kw:['DIABETES','DIABET','MELLITUS','HIPERGLUCEMIA','GLUCEMIA'] },
  { grupo:'Cardiovascular / HTA',kw:['HIPERTENSION','HIPERTENSIVA','ENFERMEDAD CARDIACA','CARDIACA','CARDIOPATIA','CARDIO','INFARTO','ANGINA','ISQUEMIC','ARRITMIA','INSUFICIENCIA CARDIACA','ATEROSCLEROSIS','CORONARI','VALVULOPATIA'] },
  { grupo:'Antibióticos',        kw:['INFECCION','INFECCIOSA','BRONQUITIS','AMIGDALITIS','FARINGITIS','NEUMONIA','SINUSITIS','OTITIS','SEPSIS','BACTERI','ABSCESO','CELULITIS','URINARIA','GASTROENTERITIS'] },
  { grupo:'Nutrición',           kw:['DESNUTRICION','DEFICIENCIA','NUTRICIONAL','MALNUTRICION','ANEMIA','AVITAMINOSIS','CARENCIA'] }
];
function cohorteDesdeCie(cie10){
  var s = sinAcentos(cie10);
  if(!s) return 'SIN COHORTE';
  for(var i=0;i<COHORTE_CIE.length;i++){
    var def=COHORTE_CIE[i];
    for(var j=0;j<def.kw.length;j++){ if(s.indexOf(def.kw[j])>=0) return def.grupo; }
  }
  return 'OTROS DIAGN\u00d3STICOS';
}

function claveDoc(o){ return o.documento+'\u241f'+o.bodega; }

function enriquecer(objs){
  return objs.map(function(o){
    o.eps = consolidarEps(o);
    var geo = homologarGeo(o.bodega);
    o.zona = geo.zona; o.depto = geo.depto; o.codBodega = geo.codigo;
    o.inactivo = esInactivo(o.estado);
    o.activo = !o.inactivo;
    o.evento = esEvento(o.contrato);
    o.capita = esCapita(o.contrato);
    o.tieneSoporte = lineaTieneSoporte(o);
    o.dif = difLinea(o);
    o.entregadaLinea = (o.dif === 0);
    o.pendiente = o.dif < 0 ? Math.abs(o.dif) : 0;
    o.cohorte = cohorteDesdeCie(o.cie10Raw);
    o.fechaDate = parseFecha(o.fechaRaw);
    o.clave = claveDoc(o);
    return o;
  });
}

/* ---------- Agrupaciones ---------- */
function aplicarFiltros(lineas){
  return lineas.filter(function(o){
    if(FILTROS.bodega && sinAcentos(o.bodega).indexOf(sinAcentos(FILTROS.bodega))<0) return false;
    if(FILTROS.eps && o.eps!==FILTROS.eps) return false;
    if(FILTROS.depto && o.depto!==FILTROS.depto) return false;
    if(FILTROS.zona && o.zona!==FILTROS.zona) return false;
    if(FILTROS.contrato && o.contrato!==FILTROS.contrato) return false;
    return true;
  });
}
function soloActivas(lineas){ return lineas.filter(function(o){ return o.activo; }); }

/* Agrupa líneas en dispensas (Documento+Bodega).
   Entregada solo si TODAS sus líneas tienen Diferencia === 0. */
function agruparDispensas(lineas){
  var map = {};
  lineas.forEach(function(o){
    var k = o.clave;
    if(!map[k]) map[k] = { clave:k, documento:o.documento, bodega:o.bodega, contrato:o.contrato,
      eps:o.eps, zona:o.zona, depto:o.depto, usuarioCrea:o.usuarioCrea, lineas:[],
      evento:o.evento, capita:o.capita, algunSoporte:false, fechaRaw:o.fechaRaw, fechaDate:o.fechaDate };
    map[k].lineas.push(o);
    if(o.tieneSoporte) map[k].algunSoporte = true;
    if(o.evento) map[k].evento = true;
    if(o.capita) map[k].capita = true;
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
    if(!byBod[d.bodega]) byBod[d.bodega]={ bodega:d.bodega, zona:d.zona, depto:d.depto, total:0, entregadas:0 };
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
  // TODAS las dispensas de contrato EVENTO (entregadas + pendientes)
  var dispEvento = agruparDispensas(lineas).filter(function(d){ return d.evento; });
  var totEvento = dispEvento.length;
  var eventoEntregadas = dispEvento.filter(function(d){ return d.entregada; });
  var totEnt = eventoEntregadas.length, totPen = totEvento - totEnt;

  // Análisis de soporte SOLO sobre las de evento entregadas
  var disp = eventoEntregadas;
  var byBod={};
  disp.forEach(function(d){
    if(!byBod[d.bodega]) byBod[d.bodega]={ bodega:d.bodega, total:0, con:0 };
    byBod[d.bodega].total++;
    if(d.algunSoporte) byBod[d.bodega].con++;
  });
  var rows=Object.keys(byBod).map(function(k){return byBod[k];}).sort(function(a,b){return b.total-a.total;});
  var totT=disp.length, totC=disp.filter(function(d){return d.algunSoporte;}).length, totS=totT-totC;

  // Totalizador global de EVENTO + subgrupos
  $('#statsSoporte').innerHTML =
    statCard('Total dispensas Evento', fmt(totEvento), 'entregadas + pendientes') +
    statCard('Evento entregadas', fmt(totEnt), pct1(totEnt,totEvento)+'% (100% líneas dif=0)', false, pct(totEnt,totEvento)) +
    statCard('Evento pendientes', fmt(totPen), pct1(totPen,totEvento)+'% (alguna dif<0)', totPen>0) +
    statCard('Con soporte', fmt(totC), pct1(totC,totT)+'% de entregadas', false, pct(totC,totT)) +
    statCard('Sin soporte', fmt(totS), pct1(totS,totT)+'% de entregadas', totS>0);

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
  SOP_CACHE=rows; SOP_TOT={totEvento:totEvento,totEnt:totEnt,totPen:totPen,totT:totT,totC:totC,totS:totS}; SOP_DISP=dispEvento;
}
var SOP_CACHE=[], SOP_TOT={}, SOP_DISP=[];

/* ================================================================
   RENDER 3 — COHORTES
   ================================================================ */
/* Cohortes ahora se toman de la columna DESCRIPCION COHORTE del reporte.
   Ya NO se clasifica por CIE-10 ni por nombre de medicamento.
   Se usa estrictamente el valor de la columna COHORTE. */

function renderCohortes(){
  var lineas = soloActivas(aplicarFiltros(RAW));
  var totLineasGlobal = lineas.length;
  // Agrupar por cohorte (CIE-10)
  var cohMap={};
  lineas.forEach(function(l){
    var c = l.cohorte || 'SIN COHORTE';
    if(!cohMap[c]) cohMap[c]={ cohorte:c, pacSet:{}, codSet:{}, bodSet:{}, lineas:0, entregadas:0, pendientes:0, undEnt:0, undPen:0 };
    var g=cohMap[c];
    g.lineas++;
    g.pacSet[l.paciente]=1; g.codSet[l.codigo]=1; g.bodSet[l.bodega]=1;
    // Conversión estricta a número antes de sumar
    if(l.entregadaLinea){ g.entregadas++; g.undEnt += Number(l.entregado)||0; }
    else { g.pendientes++; g.undPen += Number(l.pendiente)||0; }
  });
  var cohortes = Object.keys(cohMap).map(function(k){
    var g=cohMap[k];
    return { cohorte:g.cohorte, lineas:g.lineas, pacients:Object.keys(g.pacSet).length,
      partPct:pct1(g.lineas,totLineasGlobal),
      entregadas:g.entregadas, pendientes:g.pendientes, pctC:pct1(g.entregadas,g.lineas),
      undEnt:Math.round(g.undEnt), undPen:Math.round(g.undPen),
      codigos:Object.keys(g.codSet).length, bodegas:Object.keys(g.bodSet).length };
  }).sort(function(a,b){ return b.lineas-a.lineas; });

  var totL=0,totP=0,totE=0,totPn=0,totUE=0,totUP=0;
  cohortes.forEach(function(c){ totL+=c.lineas; totP+=c.pacients; totE+=c.entregadas; totPn+=c.pendientes; totUE+=c.undEnt; totUP+=c.undPen; });

  $('#statsCohortes').innerHTML =
    statCard('Cohortes (CIE-10)', fmt(cohortes.length), 'grupos de diagnóstico') +
    statCard('Líneas clasificadas', fmt(totL), fmt(totP)+' pacientes únicos') +
    statCard('Líneas entregadas', fmt(totE), pct1(totE,totL)+'% cumplimiento', false, pct(totE,totL)) +
    statCard('Líneas pendientes', fmt(totPn), pct1(totPn,totL)+'%', totPn>0);

  var tb=$('#tblCohortes tbody'); tb.innerHTML='';
  cohortes.forEach(function(c){
    tb.insertAdjacentHTML('beforeend','<tr><td class="txt wrapcell">'+c.cohorte+'</td><td>'+fmt(c.pacients)+'</td><td>'+fmt(c.lineas)+'</td><td>'+
      c.partPct+'%</td><td>'+fmt(c.entregadas)+'</td><td>'+fmt(c.pendientes)+'</td><td class="'+pctCls(c.pctC)+'">'+c.pctC+'%</td><td>'+fmt(c.undEnt)+'</td><td>'+fmt(c.undPen)+'</td></tr>');
  });
  if(cohortes.length){
    tb.insertAdjacentHTML('beforeend','<tr class="total-row"><td class="txt">TOTAL</td><td>-</td><td>'+fmt(totL)+'</td><td>100%</td><td>'+
      fmt(totE)+'</td><td>'+fmt(totPn)+'</td><td>'+pct1(totE,totL)+'%</td><td>'+fmt(totUE)+'</td><td>'+fmt(totUP)+'</td></tr>');
  }

  // Filtro cohorte / bodega
  opts($('#fCohorte'), cohortes.map(function(c){return c.cohorte;}), true);
  var bods=[]; lineas.forEach(function(l){ if(bods.indexOf(l.bodega)<0) bods.push(l.bodega); });
  opts($('#fCohorteBodega'), bods.sort(), true);

  COH_DATA = { lineas:lineas, cohortes:cohortes };
  renderCohortesTop();

  var diagBox=$('#cohortesDiag');
  var sinCoh = lineas.filter(function(l){ return l.cohorte==='SIN COHORTE'; });
  if(sinCoh.length>0){
    diagBox.style.display='block';
    diagBox.innerHTML='<b>'+fmt(sinCoh.length)+'</b> líneas sin diagnóstico CIE-10 ('+pct1(sinCoh.length,lineas.length)+'% del filtrado).';
  } else { diagBox.style.display='none'; }
}
var COH_DATA={};

function renderCohortesTop(){
  var cLabel=$('#fCohorte').value, bLabel=$('#fCohorteBodega').value;
  var lineas = COH_DATA.lineas || [];
  if(cLabel) lineas = lineas.filter(function(l){ return l.cohorte===cLabel; });
  if(bLabel) lineas = lineas.filter(function(l){ return l.bodega===bLabel; });
  var codMap={};
  lineas.forEach(function(l){
    if(!codMap[l.codigo]) codMap[l.codigo]={ codigo:l.codigo||'(sin código)', desc:l.descripcion, lineas:0, pacs:{}, pen:0 };
    codMap[l.codigo].lineas++; codMap[l.codigo].pacs[l.paciente]=1; codMap[l.codigo].pen+=Number(l.pendiente)||0;
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

  // Rango de fechas de las dispensas inactivas
  var fechas = lineas.map(function(o){ return o.fechaDate; }).filter(function(d){ return d; });
  var fMin=null,fMax=null;
  fechas.forEach(function(d){ if(!fMin||d<fMin) fMin=d; if(!fMax||d>fMax) fMax=d; });
  var rango = fechas.length ? (fmtFecha(fMin)+'  a  '+fmtFecha(fMax)) : 'Sin fecha';

  $('#statsInactivas').innerHTML =
    statCard('Dispensas inactivas', fmt(tot), 'únicas (doc+bodega)', tot>0) +
    statCard('Líneas involucradas', fmt(lineas.length), 'filas del reporte') +
    statCard('Usuarios creación', fmt(rowsU.length), 'distintos') +
    statCard('Bodegas', fmt(rowsB.length), 'involucradas') +
    '<div class="stat"><div class="label">Rango de fechas</div><div class="value" style="font-size:15px;line-height:1.35;margin-top:6px;">'+rango+'</div></div>';

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
  var head=['Bodega','Zona','Departamento','Dispensas','Entregadas','Pendientes','Eficiencia %','Índice pendiente %'];
  var rows=DISP_CACHE.map(function(r){ var p=r.total-r.entregadas;
    return [r.bodega,r.zona||'',r.depto||'',r.total,r.entregadas,p,pct1(r.entregadas,r.total),pct1(p,r.total)]; });
  rows.push(['TOTAL','','',DISP_TOT.totD,DISP_TOT.totE,DISP_TOT.totP,pct1(DISP_TOT.totE,DISP_TOT.totD),pct1(DISP_TOT.totP,DISP_TOT.totD)]);
  exportar('indicador_dispensa.xlsx',[{name:'Dispensa',data:[head].concat(rows)}]);
}

function expSoporte(){
  var resumen=[
    ['Métrica','Cantidad','%'],
    ['Total dispensas Evento', SOP_TOT.totEvento, '100%'],
    ['Evento entregadas', SOP_TOT.totEnt, pct1(SOP_TOT.totEnt,SOP_TOT.totEvento)+'%'],
    ['Evento pendientes', SOP_TOT.totPen, pct1(SOP_TOT.totPen,SOP_TOT.totEvento)+'%'],
    ['Con soporte (de entregadas)', SOP_TOT.totC, pct1(SOP_TOT.totC,SOP_TOT.totT)+'%'],
    ['Sin soporte (de entregadas)', SOP_TOT.totS, pct1(SOP_TOT.totS,SOP_TOT.totT)+'%']
  ];
  var head=['Bodega','Entregadas Evento','Con soporte','Sin soporte','% Con soporte'];
  var rows=SOP_CACHE.map(function(r){ var s=r.total-r.con; return [r.bodega,r.total,r.con,s,pct1(r.con,r.total)]; });
  rows.push(['TOTAL',SOP_TOT.totT,SOP_TOT.totC,SOP_TOT.totS,pct1(SOP_TOT.totC,SOP_TOT.totT)]);
  var det=[['Documento','Bodega','Contrato','EPS','Estado dispensa','¿Con soporte?','Líneas','Fecha']];
  SOP_DISP.forEach(function(d){ det.push([d.documento,d.bodega,d.contrato,d.eps,d.entregada?'ENTREGADA':'PENDIENTE',d.algunSoporte?'SÍ':'NO',d.lineas.length,d.fechaRaw]); });
  exportar('indicador_soporte_evento.xlsx',[{name:'Totalizador',data:resumen},{name:'Por bodega',data:[head].concat(rows)},{name:'Detalle',data:det}]);
}

function expCohortes(){
  var head=['Cohorte (CIE-10)','Pacientes','Líneas','% Part.','Entregadas','Pendientes','% Cumpl.','Und. entregadas','Und. pendientes'];
  var rows=(COH_DATA.cohortes||[]).map(function(c){ return [c.cohorte,c.pacients,c.lineas,c.partPct,c.entregadas,c.pendientes,c.pctC,c.undEnt,c.undPen]; });
  var det=[['Cohorte','Documento','Bodega','Paciente','Código','Descripción','CIE 10','Entregado','Pendiente','Diferencia','Entregada']];
  (COH_DATA.lineas||[]).forEach(function(l){
    det.push([l.cohorte,l.documento,l.bodega,l.paciente,l.codigo,l.descripcion,l.cie10Raw,l.entregado,l.pendiente,l.dif,l.entregadaLinea?'SÍ':'NO']);
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
  var contratos={}, epss={}, deptos={}, zonas={};
  RAW.forEach(function(o){
    if(o.contrato) contratos[o.contrato]=1;
    if(o.eps) epss[o.eps]=1;
    if(o.depto) deptos[o.depto]=1;
    if(o.zona) zonas[o.zona]=1;
  });
  function fill(sel, arr){ sel.innerHTML='<option value="">'+sel.options[0].text+'</option>'+
    arr.sort().map(function(v){return '<option value="'+v+'">'+v+'</option>';}).join(''); }
  fill($('#fEps'), Object.keys(epss));
  fill($('#fDepto'), Object.keys(deptos));
  fill($('#fZona'), Object.keys(zonas));
  fill($('#fContrato'), Object.keys(contratos));
}

function mostrarReporteLimpieza(rep, mapa){
  var faltan=[];
  ['documento','bodega','estado','contrato','diferencia','cohorte'].forEach(function(k){
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
      FILTROS={bodega:'',eps:'',depto:'',zona:'',contrato:''};
      $('#fBodega').value=''; $('#fEps').value=''; $('#fDepto').value=''; $('#fZona').value=''; $('#fContrato').value='';
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
    FILTROS.bodega=$('#fBodega').value; FILTROS.eps=$('#fEps').value;
    FILTROS.depto=$('#fDepto').value; FILTROS.zona=$('#fZona').value;
    FILTROS.contrato=$('#fContrato').value; renderTodo(); }
  $('#fBodega').addEventListener('input', onFiltro);
  ['#fEps','#fDepto','#fZona','#fContrato'].forEach(function(s){ $(s).addEventListener('change', onFiltro); });
  $('#btnLimpiar').addEventListener('click', function(){
    $('#fBodega').value=''; $('#fEps').value=''; $('#fDepto').value=''; $('#fZona').value=''; $('#fContrato').value='';
    FILTROS={bodega:'',eps:'',depto:'',zona:'',contrato:''}; renderTodo();
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

