/* ================================================================
   Reporte de Dispensación — app local unificada
   Fuente única: archivo "Reporte de Dispensación" (.xlsx / .csv)
   100% en el navegador. Sin nube.
   ================================================================ */
'use strict';

/* ---------- Estado global ---------- */
var RAW = [];        // filas enriquecidas (nivel línea) — acumuladas de todos los archivos
var META = { archivo:'', fecha:'', filasCrudas:0, filasUsadas:0 };
var ARCHIVOS = [];   // {nombre, lineas, disp, bodegas} por archivo cargado
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
  renderCapitaSoporte(lineas);
}
var DISP_CACHE=[], DISP_TOT={};

/* ---------- CAPITA con/sin soporte por bodega (dentro de Indicador de dispensa) ----------
   Toma las dispensas de contrato CAPITA (activas) y las separa por soporte,
   agrupadas por bodega. "Con soporte" = al menos una linea con registro de soporte. */
var CAP_CACHE=[], CAP_TOT={};
function renderCapitaSoporte(lineas){
  var dispCapita = agruparDispensas(lineas).filter(function(d){ return d.capita; });
  var byBod={};
  dispCapita.forEach(function(d){
    if(!byBod[d.bodega]) byBod[d.bodega]={ bodega:d.bodega, total:0, con:0 };
    byBod[d.bodega].total++;
    if(d.algunSoporte) byBod[d.bodega].con++;
  });
  var rows=Object.keys(byBod).map(function(k){return byBod[k];}).sort(function(a,b){return b.total-a.total;});
  var totT=dispCapita.length, totC=dispCapita.filter(function(d){return d.algunSoporte;}).length, totS=totT-totC;

  $('#statsCapita').innerHTML =
    statCard('Total dispensas Cápita', fmt(totT), 'activas (doc+bodega)') +
    statCard('Con soporte', fmt(totC), pct1(totC,totT)+'% del total', false, pct(totC,totT)) +
    statCard('Sin soporte', fmt(totS), pct1(totS,totT)+'% del total', totS>0) +
    statCard('Bodegas', fmt(rows.length), 'con cápita');

  var tb=$('#tblCapita tbody'); tb.innerHTML='';
  if(!rows.length){
    tb.innerHTML='<tr><td colspan="5" class="txt" style="text-align:center;color:#94a3b8">Sin dispensas de contrato CÁPITA en el filtro actual.</td></tr>';
  } else {
    rows.forEach(function(r){
      var s=r.total-r.con, pc=pct1(r.con,r.total);
      tb.insertAdjacentHTML('beforeend','<tr><td class="txt wrapcell">'+r.bodega+'</td><td>'+fmt(r.total)+'</td><td>'+
        fmt(r.con)+'</td><td>'+fmt(s)+'</td><td class="'+pctCls(pc)+'">'+pc+'%</td></tr>');
    });
    tb.insertAdjacentHTML('beforeend','<tr class="total-row"><td class="txt">TOTAL</td><td>'+fmt(totT)+'</td><td>'+
      fmt(totC)+'</td><td>'+fmt(totS)+'</td><td>'+pct1(totC,totT)+'%</td></tr>');
  }

  var sel=$('#pieCapitaSelect');
  var vals=['(Todas)'].concat(rows.map(function(r){return r.bodega;}));
  sel.innerHTML=vals.map(function(v){return '<option>'+v+'</option>';}).join('');
  function pintar(){
    var c,s;
    if(sel.value==='(Todas)'||!byBod[sel.value]){ c=totC; s=totS; }
    else { var r=byBod[sel.value]; c=r.con; s=r.total-r.con; }
    donut($('#pieCapita'), c, s, '#0b5fa5', '#d98a2b', pct1(c,c+s)+'%');
    legend($('#pieCapitaLegend'), [
      { c:'#0b5fa5', l:'Con soporte', v:fmt(c) },
      { c:'#d98a2b', l:'Sin soporte', v:fmt(s) }
    ]);
  }
  sel.onchange=pintar; pintar();
  CAP_CACHE=rows; CAP_TOT={totT:totT,totC:totC,totS:totS};
}

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
   RENDER 5 — DISPENSACIONES POR USUARIO
   Agrupa por "Usuario Creación". Cada dispensa = Documento + Bodega.
   Entregada = todas sus líneas con Diferencia = 0.
   Responde a los filtros globales superiores.
   ================================================================ */
var USU_CACHE=[], USU_TOT={}, USU_DISP=[];
function mesKey(d){ return d ? d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2) : ''; }
var MESES_ES=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
function mesLabel(k){ if(!k) return ''; var p=k.split('-'); return MESES_ES[(+p[1])-1]+' '+p[0]; }

function renderUsuarios(){
  var lineas = soloActivas(aplicarFiltros(RAW));
  var dispAll = agruparDispensas(lineas);
  USU_DISP = dispAll;

  // Opciones de mes (a partir de las fechas de las dispensas)
  var mesesSet = {};
  dispAll.forEach(function(d){ var k=mesKey(d.fechaDate); if(k) mesesSet[k]=1; });
  var meses = Object.keys(mesesSet).sort();
  var selMes = $('#fUsuarioMes');
  var curMes = selMes.value;
  selMes.innerHTML = '<option value="">Todos los meses</option>' +
    meses.map(function(k){ return '<option value="'+k+'">'+mesLabel(k)+'</option>'; }).join('');
  if(meses.indexOf(curMes)>=0) selMes.value=curMes; else selMes.value='';
  var mesFiltro = selMes.value;

  // Dispensas segun el mes seleccionado (para KPIs y tabla)
  var disp = mesFiltro ? dispAll.filter(function(d){ return mesKey(d.fechaDate)===mesFiltro; }) : dispAll;

  // Agrupacion por usuario de creacion
  var byUsu = {};
  disp.forEach(function(d){
    var u = d.usuarioCrea || 'SIN USUARIO';
    if(!byUsu[u]) byUsu[u] = { usuario:u, total:0, entregadas:0, pendientes:0, lineas:0 };
    byUsu[u].total++;
    if(d.entregada) byUsu[u].entregadas++; else byUsu[u].pendientes++;
    byUsu[u].lineas += d.lineas.length;
  });
  var rows = Object.keys(byUsu).map(function(k){
    var r = byUsu[k];
    r.eficiencia = pct1(r.entregadas, r.total);
    return r;
  });

  // Totales globales
  var totUsuarios = rows.length;
  var totDisp = disp.length;
  var promedio = totUsuarios>0 ? Math.round((totDisp/totUsuarios)*10)/10 : 0;

  $('#statsUsuarios').innerHTML =
    statCard('Total usuarios activos', fmt(totUsuarios), 'con ≥1 dispensa') +
    statCard('Total dispensas generadas', fmt(totDisp), 'únicas (doc+bodega)') +
    statCard('Promedio de dispensas por usuario', fmt(promedio), 'documentos / usuario');

  USU_CACHE = rows.slice();
  USU_TOT = { totUsuarios:totUsuarios, totDisp:totDisp, promedio:promedio };

  pintarTablaUsuarios();
  pintarUsuariosPorMes(dispAll);

  // Select de usuario para la curva diaria
  var selCurva = $('#fUsuarioCurva');
  var curU = selCurva.value;
  var usuariosOrden = rows.slice().sort(function(a,b){return b.total-a.total;}).map(function(r){return r.usuario;});
  selCurva.innerHTML = '<option value="">Selecciona un usuario…</option>' +
    usuariosOrden.map(function(u){ return '<option value="'+u+'">'+u+'</option>'; }).join('');
  if(usuariosOrden.indexOf(curU)>=0) selCurva.value=curU;
  else if(usuariosOrden.length) selCurva.value=usuariosOrden[0];
  pintarCurvaUsuario();
}

/* Tabla "Dispensas por mes" (todos los meses, ignora el filtro de mes). */
function pintarUsuariosPorMes(dispAll){
  var byMes={};
  dispAll.forEach(function(d){
    var k=mesKey(d.fechaDate) || 'SIN FECHA';
    if(!byMes[k]) byMes[k]={mes:k,total:0,entregadas:0,pendientes:0};
    byMes[k].total++;
    if(d.entregada) byMes[k].entregadas++; else byMes[k].pendientes++;
  });
  var rows=Object.keys(byMes).sort().map(function(k){return byMes[k];});
  var tb=$('#tblUsuariosMes tbody'); tb.innerHTML='';
  if(!rows.length){ tb.innerHTML='<tr><td colspan="5" class="txt" style="text-align:center;color:#94a3b8">Sin datos.</td></tr>'; return; }
  var tot=dispAll.length;
  rows.forEach(function(r){
    var lbl = r.mes==='SIN FECHA' ? 'Sin fecha' : mesLabel(r.mes);
    tb.insertAdjacentHTML('beforeend','<tr><td class="txt">'+lbl+'</td><td>'+fmt(r.total)+'</td><td>'+
      fmt(r.entregadas)+'</td><td>'+fmt(r.pendientes)+'</td><td>'+pct1(r.total,tot)+'%</td></tr>');
  });
}

/* Curva diaria (line chart SVG) del usuario seleccionado.
   Respeta filtros globales + filtro de mes; una serie de entregadas y otra de total. */
function pintarCurvaUsuario(){
  var svg=$('#chartUsuarioDia');
  var leg=$('#legendUsuarioDia');
  var tbDia=$('#tblUsuarioDia tbody');
  var u=$('#fUsuarioCurva').value;
  var mesFiltro=$('#fUsuarioMes').value;
  if(tbDia) tbDia.innerHTML='';
  if(!u){ svg.innerHTML='<text x="480" y="210" text-anchor="middle" fill="#94a3b8" font-size="15">Selecciona un usuario para ver su curva diaria.</text>'; leg.innerHTML=''; return; }

  var disp=USU_DISP.filter(function(d){ return (d.usuarioCrea||'SIN USUARIO')===u && d.fechaDate && (!mesFiltro || mesKey(d.fechaDate)===mesFiltro); });
  if(!disp.length){ svg.innerHTML='<text x="480" y="210" text-anchor="middle" fill="#94a3b8" font-size="15">Sin dispensas con fecha para este usuario.</text>'; leg.innerHTML=''; return; }

  var byDia={};
  disp.forEach(function(d){
    var k=fmtFecha(d.fechaDate);
    if(!byDia[k]) byDia[k]={dia:k,total:0,entregadas:0};
    byDia[k].total++;
    if(d.entregada) byDia[k].entregadas++;
  });
  var dias=Object.keys(byDia).sort();
  var serieTotal=dias.map(function(k){return byDia[k].total;});
  var serieEnt=dias.map(function(k){return byDia[k].entregadas;});
  lineChart(svg, dias, [
    { vals:serieTotal, color:'#2563eb', label:'Total', dyLabel:-12 },
    { vals:serieEnt,   color:'#16a34a', label:'Entregadas', dyLabel:20 }
  ]);
  legend(leg, [
    { c:'#2563eb', l:'Total dispensas', v:fmt(serieTotal.reduce(function(a,b){return a+b;},0)) },
    { c:'#16a34a', l:'Entregadas', v:fmt(serieEnt.reduce(function(a,b){return a+b;},0)) }
  ]);
  // Tabla de detalle diario (respaldo numerico exacto)
  if(tbDia){
    dias.forEach(function(k){
      var r=byDia[k]; var pen=r.total-r.entregadas;
      tbDia.insertAdjacentHTML('beforeend','<tr><td class="txt">'+k+'</td><td>'+fmt(r.total)+'</td><td>'+
        fmt(r.entregadas)+'</td><td>'+fmt(pen)+'</td></tr>');
    });
  }
}

/* Grafico de lineas SVG reutilizable, con etiquetas de valor en cada punto.
   labels: array de etiquetas eje X. series: [{vals:[], color, label, dyLabel}]. */
function lineChart(svg, labels, series){
  svg.innerHTML='';
  var ns='http://www.w3.org/2000/svg';
  var W=960,H=420, mL=52,mR=22,mT=34,mB=70;
  svg.setAttribute('viewBox','0 0 '+W+' '+H);
  var pw=W-mL-mR, ph=H-mT-mB;
  var maxV=0; series.forEach(function(s){ s.vals.forEach(function(v){ if(v>maxV) maxV=v; }); });
  if(maxV<=0) maxV=1;
  maxV = Math.ceil(maxV*1.12);
  var n=labels.length;
  function x(i){ return n<=1 ? mL+pw/2 : mL + (pw*i/(n-1)); }
  function y(v){ return mT + ph - (ph*v/maxV); }
  function mk(tag,attrs){ var e=document.createElementNS(ns,tag); for(var k in attrs) e.setAttribute(k,attrs[k]); return e; }

  // Grid + eje Y (5 lineas)
  var steps=4;
  for(var g=0; g<=steps; g++){
    var vv=maxV*g/steps, yy=y(vv);
    svg.appendChild(mk('line',{x1:mL,y1:yy,x2:W-mR,y2:yy,stroke:'#e2e8f0','stroke-width':1}));
    var t=mk('text',{x:mL-8,y:yy+4,'text-anchor':'end','font-size':13,fill:'#64748b'});
    t.textContent=String(Math.round(vv)); svg.appendChild(t);
  }
  // Guias verticales + etiquetas eje X
  var stepX = n>16 ? Math.ceil(n/12) : 1;
  for(var i=0;i<n;i++){
    var showX = (i%stepX===0 || i===n-1);
    if(showX){
      svg.appendChild(mk('line',{x1:x(i),y1:mT,x2:x(i),y2:mT+ph,stroke:'#f1f5f9','stroke-width':1}));
      var lbl=labels[i].length>5 ? labels[i].slice(5) : labels[i];
      var xx=x(i);
      var tx=mk('text',{x:xx,y:H-mB+20,'text-anchor':'end','font-size':12,fill:'#475569'});
      tx.textContent=lbl;
      tx.setAttribute('transform','rotate(-42 '+xx+' '+(H-mB+20)+')');
      svg.appendChild(tx);
    }
  }
  // Series (linea + puntos + etiqueta de valor en cada punto)
  series.forEach(function(s){
    var pts=s.vals.map(function(v,i){ return x(i)+','+y(v); }).join(' ');
    svg.appendChild(mk('polyline',{points:pts,fill:'none',stroke:s.color,'stroke-width':2.6,'stroke-linejoin':'round','stroke-linecap':'round'}));
    var dy = s.dyLabel!=null ? s.dyLabel : -12;
    s.vals.forEach(function(v,i){
      var cx=x(i), cy=y(v);
      var c=mk('circle',{cx:cx,cy:cy,r:4,fill:'#fff',stroke:s.color,'stroke-width':2});
      var tt=mk('title',{}); tt.textContent=labels[i]+' · '+s.label+': '+v; c.appendChild(tt);
      svg.appendChild(c);
      if(v>0){
        var vt=mk('text',{x:cx,y:cy+dy,'text-anchor':'middle','font-size':12,'font-weight':'700',fill:s.color,'paint-order':'stroke','stroke':'#ffffff','stroke-width':3.5,'stroke-linejoin':'round'});
        vt.textContent=String(v); svg.appendChild(vt);
      }
    });
  });
}

/* Aplica busqueda de texto + orden y repinta la tabla (sin recalcular). */
function pintarTablaUsuarios(){
  var q = sinAcentos($('#fUsuarioBuscar').value || '');
  var orden = $('#fUsuarioOrden').value || 'total';
  var rows = USU_CACHE.filter(function(r){
    return !q || sinAcentos(r.usuario).indexOf(q) >= 0;
  });
  rows.sort(function(a,b){
    if(orden==='eficiencia'){ return (b.eficiencia-a.eficiencia) || (b.total-a.total); }
    return (b.total-a.total) || (b.eficiencia-a.eficiencia);
  });

  var tb = $('#tblUsuarios tbody'); tb.innerHTML='';
  if(!rows.length){
    tb.innerHTML='<tr><td colspan="6" class="txt" style="text-align:center;color:#94a3b8">Sin usuarios que coincidan.</td></tr>';
    return;
  }
  rows.forEach(function(r){
    tb.insertAdjacentHTML('beforeend',
      '<tr><td class="txt wrapcell">'+r.usuario+'</td>'+
      '<td>'+fmt(r.total)+'</td>'+
      '<td>'+fmt(r.entregadas)+'</td>'+
      '<td>'+fmt(r.pendientes)+'</td>'+
      '<td>'+fmt(r.lineas)+'</td>'+
      '<td class="'+pctCls(r.eficiencia)+'">'+r.eficiencia+'%</td></tr>');
  });
}

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

function expCapita(){
  var head=['Bodega','Dispensas cápita','Con soporte','Sin soporte','% Con soporte'];
  var rows=CAP_CACHE.map(function(r){ var s=r.total-r.con; return [r.bodega,r.total,r.con,s,pct1(r.con,r.total)]; });
  rows.push(['TOTAL',CAP_TOT.totT,CAP_TOT.totC,CAP_TOT.totS,pct1(CAP_TOT.totC,CAP_TOT.totT)]);
  var resumen=[
    ['Métrica','Cantidad','%'],
    ['Total dispensas Cápita', CAP_TOT.totT, '100%'],
    ['Con soporte', CAP_TOT.totC, pct1(CAP_TOT.totC,CAP_TOT.totT)+'%'],
    ['Sin soporte', CAP_TOT.totS, pct1(CAP_TOT.totS,CAP_TOT.totT)+'%']
  ];
  exportar('capita_con_sin_soporte.xlsx',[{name:'Totalizador',data:resumen},{name:'Por bodega',data:[head].concat(rows)}]);
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

function expUsuarios(){
  var head=['Usuario de creación','Dispensas totales','Entregadas','Pendientes','Total líneas','% Eficiencia'];
  var rows=USU_CACHE.slice().sort(function(a,b){return b.total-a.total;}).map(function(r){
    return [r.usuario,r.total,r.entregadas,r.pendientes,r.lineas,pct1(r.entregadas,r.total)]; });
  var tE=0,tP=0,tL=0,tT=0;
  USU_CACHE.forEach(function(r){ tE+=r.entregadas; tP+=r.pendientes; tL+=r.lineas; tT+=r.total; });
  rows.push(['TOTAL',tT,tE,tP,tL,pct1(tE,tT)]);
  var resumen=[
    ['Métrica','Valor'],
    ['Total usuarios activos', USU_TOT.totUsuarios||0],
    ['Total dispensas generadas', USU_TOT.totDisp||0],
    ['Promedio de dispensas por usuario', USU_TOT.promedio||0]
  ];
  exportar('dispensaciones_por_usuario.xlsx',[{name:'Resumen',data:resumen},{name:'Por usuario',data:[head].concat(rows)},{name:'Por mes',data:expUsuariosPorMes()}]);
}

function expUsuariosPorMes(){
  var byMes={};
  (USU_DISP||[]).forEach(function(d){
    var k=mesKey(d.fechaDate) || 'SIN FECHA';
    if(!byMes[k]) byMes[k]={mes:k,total:0,entregadas:0,pendientes:0};
    byMes[k].total++;
    if(d.entregada) byMes[k].entregadas++; else byMes[k].pendientes++;
  });
  var out=[['Mes','Dispensas','Entregadas','Pendientes','% Eficiencia']];
  Object.keys(byMes).sort().forEach(function(k){
    var r=byMes[k];
    out.push([k==='SIN FECHA'?'Sin fecha':mesLabel(k), r.total, r.entregadas, r.pendientes, pct1(r.entregadas,r.total)]);
  });
  return out;
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
  renderUsuarios();
  renderInactivas();
}

function poblarFiltrosGlobales(){
  var contratos={}, epss={}, deptos={}, zonas={}, bodegas={};
  RAW.forEach(function(o){
    if(o.contrato) contratos[o.contrato]=1;
    if(o.eps) epss[o.eps]=1;
    if(o.depto) deptos[o.depto]=1;
    if(o.zona) zonas[o.zona]=1;
    if(o.bodega) bodegas[o.bodega]=1;
  });
  function fill(sel, arr){ sel.innerHTML='<option value="">'+sel.options[0].text+'</option>'+
    arr.sort().map(function(v){return '<option value="'+v+'">'+v+'</option>';}).join(''); }
  fill($('#fEps'), Object.keys(epss));
  fill($('#fDepto'), Object.keys(deptos));
  fill($('#fZona'), Object.keys(zonas));
  fill($('#fContrato'), Object.keys(contratos));
  // Datalist de bodegas: lista + buscador a la vez
  var dl=$('#bodegaList');
  if(dl){ dl.innerHTML=Object.keys(bodegas).sort().map(function(v){return '<option value="'+v.replace(/"/g,'&quot;')+'"></option>';}).join(''); }
}

function mostrarReporteLimpieza(rep, mapa){
  var faltan=[];
  ['documento','bodega','estado','contrato','diferencia','cohorte'].forEach(function(k){
    if(mapa[k]<0) faltan.push(k);
  });
  var box=$('#cleanReport');
  box.style.display='block';
  // Distribucion por ESTADO (fuente para activas/inactivas)
  var estCount={}, sinEstado=0;
  RAW.forEach(function(o){ var e=o.estado||''; if(!e){sinEstado++;} else { estCount[e]=(estCount[e]||0)+1; } });
  var estStr=Object.keys(estCount).sort(function(a,b){return estCount[b]-estCount[a];})
    .map(function(e){ return e+': '+fmt(estCount[e]); }).join(' · ') || 'no disponible';
  var estWarn = mapa.estado<0;
  box.innerHTML='<div class="cr-head">✓ Archivo procesado correctamente</div><ul>'+
    '<li>Filas leídas del archivo: <b>'+fmt(rep.totalCrudas)+'</b></li>'+
    '<li>Encabezados detectados en la fila <b>'+rep.headerRow+'</b> (la fila 1 de metadata se omitió)</li>'+
    '<li>Filas divididas reconstruidas: <b>'+fmt(rep.fragmentos)+'</b></li>'+
    '<li>Filas incompletas eliminadas: <b>'+fmt(rep.incompletas)+'</b></li>'+
    '<li>Filas de datos válidas usadas: <b>'+fmt(rep.usadas)+'</b></li>'+
    '<li>Columna <b>Estado</b> '+(estWarn?'<span style="color:#b23a34">NO detectada</span>':'detectada')+' — valores: <b>'+estStr+'</b> <span class="muted">(define activas vs. inactivas)</span></li>'+
    (faltan.length?'<li style="color:#b23a34">⚠ No se reconocieron columnas: <b>'+faltan.join(', ')+'</b> (revisa los encabezados del archivo)</li>':'')+
    '</ul>';
}

/* ---------- Carga ACUMULATIVA de varios archivos (varias bodegas) ----------
   Cada archivo se procesa y sus lineas se AGREGAN a RAW. Las dispensas son
   unicas por Documento + Bodega, de modo que varias bodegas conviven sin
   pisarse. Se evita recargar dos veces el mismo archivo (por nombre). */
function claveLineaUnica(o){
  return [o.documento,o.bodega,o.codigo,o.paciente,o.fechaRaw,o.diferenciaRaw,o.entregado,o.formulado].join('\u241f');
}

function cargarArchivos(fileList){
  var files = Array.prototype.slice.call(fileList||[]);
  if(!files.length) return;
  var validos = files.filter(function(f){ return ['xlsx','xls','csv'].indexOf((f.name.split('.').pop()||'').toLowerCase())>=0; });
  if(!validos.length){ toast('Formato no admitido. Usa .xlsx, .xls o .csv',true); return; }
  // procesar en secuencia para no saturar memoria
  var i=0, agregadasTot=0, ultimoRes=null, ultimoMapa=null;
  function siguiente(){
    if(i>=validos.length){ finalizar(); return; }
    var file=validos[i++];
    if(ARCHIVOS.some(function(a){ return a.nombre===file.name; })){
      toast('El archivo "'+file.name+'" ya estaba cargado; se omite.',true);
      siguiente(); return;
    }
    $('#dbStatusText').textContent='Procesando '+file.name+'\u2026';
    leerArchivo(file, function(err, aoa){
      if(err){ toast('Error al leer '+file.name+': '+err.message,true); siguiente(); return; }
      try{
        var res = procesarAoA(aoa);
        if(!res.objs.length){ toast('Sin filas validas en '+file.name,true); siguiente(); return; }
        var nuevas = enriquecer(res.objs);
        // dedupe frente a lo ya cargado
        var vistos={}; RAW.forEach(function(o){ vistos[claveLineaUnica(o)]=1; });
        var agregadas=0, bodSet={}, dispSet={};
        nuevas.forEach(function(o){
          var k=claveLineaUnica(o);
          if(vistos[k]) return; vistos[k]=1;
          o.__src=file.name;
          RAW.push(o); agregadas++;
          bodSet[o.bodega]=1; dispSet[o.clave]=1;
        });
        agregadasTot+=agregadas;
        ARCHIVOS.push({ nombre:file.name, lineas:agregadas, bodegas:Object.keys(bodSet).length, disp:Object.keys(dispSet).length });
        ultimoRes=res.rep; ultimoMapa=res.mapa;
      }catch(e){ toast('Error al procesar '+file.name+': '+e.message,true); }
      siguiente();
    });
  }
  function finalizar(){
    if(!RAW.length){ $('#dbStatusText').textContent='Sin datos cargados'; return; }
    META.fecha=new Date().toLocaleString('es-CO');
    if(ultimoRes) mostrarReporteLimpieza(ultimoRes, ultimoMapa);
    renderLoadedFiles();
    poblarFiltrosGlobales();
    renderTodo();
    $('#filtersCard').style.display='block';
    $('#viewerCard').style.display='block';
    $('#dbDot').className='dot on';
    var totBod=contarBodegas(), totDisp=contarDispensas();
    $('#dbStatusText').textContent=fmt(RAW.length)+' lineas \u00b7 '+fmt(totDisp)+' dispensas \u00b7 '+fmt(totBod)+' bodegas \u00b7 '+ARCHIVOS.length+' archivo(s)';
    $('#fechaDatos').textContent='Actualizado: '+META.fecha;
    if(agregadasTot>0) toast('Se agregaron '+fmt(agregadasTot)+' lineas ('+ARCHIVOS.length+' archivo(s), '+fmt(totBod)+' bodegas)');
  }
  siguiente();
}

/* Total de dispensas UNICAS (Documento + Bodega) en todo lo cargado. */
function contarDispensas(){ var s={}; RAW.forEach(function(o){ s[o.clave]=1; }); return Object.keys(s).length; }
function contarBodegas(){ var s={}; RAW.forEach(function(o){ s[o.bodega]=1; }); return Object.keys(s).length; }

function renderLoadedFiles(){
  var box=$('#loadedFiles'); if(!box) return;
  if(!ARCHIVOS.length){ box.style.display='none'; box.innerHTML=''; return; }
  box.style.display='block';
  var totBod=contarBodegas(), totDisp=contarDispensas();
  var lista=ARCHIVOS.map(function(a,idx){
    return '<li><span class="lf-name">'+a.nombre+'</span>'+
      '<span class="lf-meta">'+fmt(a.lineas)+' lineas \u00b7 '+fmt(a.disp)+' dispensas \u00b7 '+a.bodegas+' bodega(s)</span>'+
      '<button type="button" class="lf-x" data-idx="'+idx+'" title="Quitar">\u2715</button></li>';
  }).join('');
  box.innerHTML='<div class="lf-head">Archivos cargados ('+ARCHIVOS.length+') \u00b7 total '+fmt(RAW.length)+' lineas \u00b7 '+fmt(totDisp)+' dispensas \u00b7 '+fmt(totBod)+' bodegas '+
    '<button type="button" class="btn btn-ghost lf-clear" id="btnResetDatos">Vaciar todo</button></div>'+
    '<ul class="lf-list">'+lista+'</ul>';
  $('#btnResetDatos').addEventListener('click', vaciarDatos);
  $$('.lf-x', box).forEach(function(b){
    b.addEventListener('click', function(){ quitarArchivo(+b.getAttribute('data-idx')); });
  });
}

function quitarArchivo(idx){
  var a=ARCHIVOS[idx]; if(!a) return;
  ARCHIVOS.splice(idx,1);
  // reconstruir RAW re-leyendo? no guardamos objetos por archivo; marcamos por nombre.
  RAW = RAW.filter(function(o){ return o.__src!==a.nombre; });
  if(!RAW.length){ vaciarDatos(); return; }
  renderLoadedFiles(); poblarFiltrosGlobales(); renderTodo();
  var totBod=contarBodegas(), totDisp=contarDispensas();
  $('#dbStatusText').textContent=fmt(RAW.length)+' lineas \u00b7 '+fmt(totDisp)+' dispensas \u00b7 '+fmt(totBod)+' bodegas \u00b7 '+ARCHIVOS.length+' archivo(s)';
  toast('Archivo "'+a.nombre+'" retirado');
}

function vaciarDatos(){
  RAW=[]; ARCHIVOS=[]; META={ archivo:'', fecha:'', filasCrudas:0, filasUsadas:0 };
  FILTROS={bodega:'',eps:'',depto:'',zona:'',contrato:''};
  $('#loadedFiles').style.display='none'; $('#loadedFiles').innerHTML='';
  $('#cleanReport').style.display='none';
  $('#filtersCard').style.display='none'; $('#viewerCard').style.display='none';
  $('#dbDot').className='dot'; $('#dbStatusText').textContent='Sin datos cargados';
  $('#fechaDatos').textContent='';
  toast('Datos vaciados');
}

function initEventos(){
  var dz=$('#dropzone'), fi=$('#fileInput');
  $('#btnBrowse').addEventListener('click', function(e){ e.stopPropagation(); fi.click(); });
  dz.addEventListener('click', function(){ fi.click(); });
  fi.addEventListener('change', function(){ if(fi.files && fi.files.length) cargarArchivos(fi.files); fi.value=''; });
  ['dragenter','dragover'].forEach(function(ev){ dz.addEventListener(ev,function(e){ e.preventDefault(); dz.classList.add('drag'); }); });
  ['dragleave','drop'].forEach(function(ev){ dz.addEventListener(ev,function(e){ e.preventDefault(); dz.classList.remove('drag'); }); });
  dz.addEventListener('drop', function(e){ if(e.dataTransfer.files && e.dataTransfer.files.length) cargarArchivos(e.dataTransfer.files); });

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

  // Filtros dispensaciones por usuario
  $('#fUsuarioBuscar').addEventListener('input', pintarTablaUsuarios);
  $('#fUsuarioOrden').addEventListener('change', pintarTablaUsuarios);
  $('#fUsuarioMes').addEventListener('change', renderUsuarios);
  $('#fUsuarioCurva').addEventListener('change', pintarCurvaUsuario);

  // Exportaciones
  $('#btnExpDispensa').addEventListener('click', function(){ if(RAW.length) expDispensa(); });
  $('#btnExpCapita').addEventListener('click', function(){ if(RAW.length) expCapita(); });
  $('#btnExpSoporte').addEventListener('click', function(){ if(RAW.length) expSoporte(); });
  $('#btnExpCohortes').addEventListener('click', function(){ if(RAW.length) expCohortes(); });
  $('#btnExpInactivas').addEventListener('click', function(){ if(RAW.length) expInactivas(); });
  $('#btnExpUsuarios').addEventListener('click', function(){ if(RAW.length) expUsuarios(); });
}

document.addEventListener('DOMContentLoaded', initEventos);

