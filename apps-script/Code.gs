/** OLTE SPORT — base de datos y API privada de inventario. */
function setupOlteSportDatabase() {
  const ss = SpreadsheetApp.openById("1AG3PLwHM5UCNGc6yhVzvfsXwyh4WnzboRe6uFsBOy7M");
  if (!ss) throw new Error('Abre Apps Script desde la hoja de cálculo de OLTE SPORT.');
  const defs = [
    {name:'Inventario',headers:['Producto','Talla','Cantidad','Estado','Última actualización'],rows:[
      ['Chaleco de hombre','XS',0],['Chaleco de hombre','S',0],['Chaleco de hombre','M',0],['Chaleco de hombre','L',0],['Chaleco de hombre','XL',0],['Chaleco de hombre','XXL',0],
      ['Chaleco de mujer','XS',0],['Chaleco de mujer','S',0],['Chaleco de mujer','M',0],['Chaleco de mujer','L',0],['Chaleco de mujer','XL',0],['Chaleco de mujer','XXL',0]]},
    {name:'Ventas',headers:['ID','Fecha','Producto','Talla','Cantidad','Costo unitario (S/)','Precio unitario (S/)','Canal','Nota','Anulada']},
    {name:'Movimientos',headers:['ID','Fecha','Tipo de movimiento','Producto','Talla','Cantidad','Detalle','ID relacionado']},
    {name:'Configuracion',headers:['Producto','Costo unitario (S/)','Precio de venta (S/)','Umbral de bajo stock'],rows:[['Chaleco de hombre',32,65,3],['Chaleco de mujer',28,60,3]]}
  ];
  defs.forEach(d=>{
    let sh=ss.getSheetByName(d.name); if(!sh) sh=ss.insertSheet(d.name);
    if(sh.getLastRow()===0){
      sh.getRange(1,1,1,d.headers.length).setValues([d.headers]);
      if(d.rows) sh.getRange(2,1,d.rows.length,d.headers.length).setValues(d.rows.map(r=>r.concat(Array(d.headers.length-r.length).fill(''))));
      sh.setFrozenRows(1); sh.getRange(1,1,1,d.headers.length).setBackground('#20251e').setFontColor('#c4f500').setFontWeight('bold'); sh.autoResizeColumns(1,d.headers.length);
    }
  });
  SpreadsheetApp.flush();
  return 'Base de OLTE SPORT preparada.';
}

/** API JSONP. Mantén la implementación de la aplicación web limitada a "Solo yo". */
function doGet(e) {
  const p=(e&&e.parameter)||{};
  const callback=String(p.callback||'').replace(/[^a-zA-Z0-9_.$]/g,'');
  if(!callback) return ContentService.createTextOutput('Falta callback').setMimeType(ContentService.MimeType.TEXT);
  let result;
  try {
    if(p.action==='read') result={ok:true,state:readState_()};
    else if(p.action==='save') { saveState_(p.data||''); result={ok:true,savedAt:new Date().toISOString()}; }
    else result={ok:false,error:'Acción no reconocida'};
  } catch(err) { result={ok:false,error:String(err&&err.message||err)}; }
  return ContentService.createTextOutput(callback+'('+JSON.stringify(result)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function readState_(){
  const ss=SpreadsheetApp.openById("1AG3PLwHM5UCNGc6yhVzvfsXwyh4WnzboRe6uFsBOy7M");
  const sizes=['XS','S','M','L','XL','XXL'];
  const stock={Hombre:{},Mujer:{}}; ['Hombre','Mujer'].forEach(t=>sizes.forEach(z=>stock[t][z]=0));
  const inv=ss.getSheetByName('Inventario');
  if(inv&&inv.getLastRow()>1) inv.getRange(2,1,inv.getLastRow()-1,3).getValues().forEach(r=>{const t=productType_(r[0]),z=String(r[1]).toUpperCase();if(stock[t]&&sizes.includes(z))stock[t][z]=Number(r[2])||0;});
  const sales=[]; const shSales=ss.getSheetByName('Ventas');
  if(shSales&&shSales.getLastRow()>1) shSales.getRange(2,1,shSales.getLastRow()-1,10).getValues().forEach(r=>sales.push({id:String(r[0]),date:dateString_(r[1]),type:productType_(r[2]),size:String(r[3]),qty:Number(r[4])||0,cost:Number(r[5])||0,price:Number(r[6])||0,channel:String(r[7]||''),note:String(r[8]||''),voided:r[9]===true||String(r[9]).toLowerCase()==='true'}));
  const moves=[]; const shMoves=ss.getSheetByName('Movimientos');
  if(shMoves&&shMoves.getLastRow()>1) shMoves.getRange(2,1,shMoves.getLastRow()-1,8).getValues().forEach(r=>moves.push({id:String(r[0]),date:dateString_(r[1]),kind:String(r[2]),type:productType_(r[3]),size:String(r[4]),qty:Number(r[5])||0,detail:String(r[6]||''),ref:String(r[7]||'')}));
  const prices={Hombre:{cost:32,price:65},Mujer:{cost:28,price:60}}; let lowStock=3;
  const conf=ss.getSheetByName('Configuracion');
  if(conf&&conf.getLastRow()>1) conf.getRange(2,1,conf.getLastRow()-1,4).getValues().forEach(r=>{let t=productType_(r[0]);if(prices[t]){prices[t]={cost:Number(r[1])||0,price:Number(r[2])||0};lowStock=Number(r[3])||0;}});
  return {stock,sales,moves,settings:{lowStock,prices}};
}
function saveState_(encoded){
  const bytes=Utilities.base64Decode(encoded); const json=Utilities.ungzip(Utilities.newBlob(bytes,'application/gzip','state.gz')).getDataAsString('UTF-8');
  const data=JSON.parse(json); if(!data||!data.stock||!Array.isArray(data.sales)||!Array.isArray(data.moves)||!data.settings) throw new Error('Los datos recibidos no tienen el formato esperado.');
  const lock=LockService.getDocumentLock(); lock.waitLock(20000);
  try {
    const ss=SpreadsheetApp.openById("1AG3PLwHM5UCNGc6yhVzvfsXwyh4WnzboRe6uFsBOy7M"); const now=new Date(); const sizes=['XS','S','M','L','XL','XXL'];
    const inv=ss.getSheetByName('Inventario'); const invRows=[];
    ['Hombre','Mujer'].forEach(t=>sizes.forEach(z=>{const n=Math.max(0,Math.floor(Number(data.stock[t]&&data.stock[t][z])||0));invRows.push([t==='Hombre'?'Chaleco de hombre':'Chaleco de mujer',z,n,n===0?'Agotado':'Disponible',now]);}));
    replaceRows_(inv,invRows);
    const salesRows=data.sales.map(s=>[String(s.id||Utilities.getUuid()),s.date||'',s.type==='Mujer'?'Chaleco de mujer':'Chaleco de hombre',s.size||'',Number(s.qty)||0,Number(s.cost)||0,Number(s.price)||0,s.channel||'',s.note||'',!!s.voided]);
    replaceRows_(ss.getSheetByName('Ventas'),salesRows);
    const moveRows=data.moves.map(m=>[String(m.id||Utilities.getUuid()),m.date||'',m.kind||'',m.type==='Mujer'?'Chaleco de mujer':'Chaleco de hombre',m.size||'',Number(m.qty)||0,m.detail||'',m.ref||'']);
    replaceRows_(ss.getSheetByName('Movimientos'),moveRows);
    const prices=data.settings.prices||{}; const low=Math.max(0,Number(data.settings.lowStock)||0);
    replaceRows_(ss.getSheetByName('Configuracion'),[['Chaleco de hombre',Number(prices.Hombre&&prices.Hombre.cost)||0,Number(prices.Hombre&&prices.Hombre.price)||0,low],['Chaleco de mujer',Number(prices.Mujer&&prices.Mujer.cost)||0,Number(prices.Mujer&&prices.Mujer.price)||0,low]]);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
}
function replaceRows_(sheet,rows){
  if(!sheet) throw new Error('Falta una pestaña. Ejecuta setupOlteSportDatabase primero.');
  const cols=sheet.getLastColumn(); const count=Math.max(0,sheet.getLastRow()-1); if(count) sheet.getRange(2,1,count,cols).clearContent();
  if(rows.length) sheet.getRange(2,1,rows.length,rows[0].length).setValues(rows);
}
function productType_(value){return String(value).toLowerCase().includes('mujer')?'Mujer':'Hombre';}
function dateString_(value){if(value instanceof Date)return Utilities.formatDate(value,Session.getScriptTimeZone(),'yyyy-MM-dd');return String(value||'').slice(0,10);}
