/** A-RY Coffee POS. Set the Apps Script project timezone to Asia/Bangkok. */
const POS = Object.freeze({
  spreadsheetId: '1nlblA50Jl8WV6y3eApE8zURov-kuCDsbtV8WNvhp4DM',
  timezone: 'Asia/Bangkok',
  orders: 'ARY_POS_Orders', menu: 'ARY_POS_Menu',
  headers: ['orderId','createdAt','businessDate','payment','total','itemsJson','note','status','voidReason','voidedAt'],
  menuHeaders: ['id','name','price','category','active','image','art'],
  maxImage: 48000, // a sheet cell holds 50,000 characters
  artKeys: ['americano','espresso','cappuccino','mocha','latte','cocoa','thai','green','honey','coconut','coconutlatte','orange','caramel','esthai','esgreen','coconutwater','juice','soda','milk'],
  seedArt: {M01:'americano',M02:'espresso',M03:'cappuccino',M04:'mocha',M05:'latte',M06:'cocoa',M07:'thai',M08:'green',M09:'honey',M10:'coconut',M11:'orange',M12:'coconutlatte',M13:'esthai',M14:'esgreen',M15:'caramel'},
  seed: [
    ['M01','อเมริกาโน่',45,'กาแฟ',true],['M02','เอสเพรสโซ่',45,'กาแฟ',true],
    ['M03','คาปูชิโน่',45,'กาแฟ',true],['M04','มอคค่า',50,'กาแฟ',true],
    ['M05','ลาเต้',50,'กาแฟ',true],['M06','โกโก้',45,'ชาและโกโก้',true],
    ['M07','ชาไทย',45,'ชาและโกโก้',true],['M08','ชาเขียว',45,'ชาและโกโก้',true],
    ['M09','อเมริกาโน่น้ำผึ้ง',50,'กาแฟพิเศษ',true],['M10','อเมริกาโน่มะพร้าว',60,'กาแฟพิเศษ',true],
    ['M11','อเมริกาโน่ส้ม',50,'กาแฟพิเศษ',true],['M12','ลาเต้มะพร้าว',65,'กาแฟพิเศษ',true],
    ['M13','เอส-ชาไทย',60,'กาแฟพิเศษ',true],['M14','เอส-ชาเขียว',60,'กาแฟพิเศษ',true],
    ['M15','คาราเมลมัคคิอาโต้',60,'กาแฟพิเศษ',true]
  ]
});

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('A-RY Coffee POS')
    .addMetaTag('viewport','width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
/** Used by Index.html to pull in Style and Script. */
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }

/** Run once from the editor after authorizing. Existing unrelated sheets are never touched. */
function setupPOS_() {
  const ss = SpreadsheetApp.openById(POS.spreadsheetId);
  let orders = ss.getSheetByName(POS.orders);
  if (!orders) {
    orders = ss.insertSheet(POS.orders);
    orders.getRange(1,1,1,POS.headers.length).setValues([POS.headers]).setFontWeight('bold');
    orders.setFrozenRows(1); orders.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  } else if (orders.getRange(1,1,1,POS.headers.length).getValues()[0].join('|') !== POS.headers.join('|')) {
    throw new Error('ชีต ARY_POS_Orders มีอยู่แล้วแต่หัวตารางไม่ตรง กรุณาตรวจสอบก่อน');
  }
  let menu = ss.getSheetByName(POS.menu);
  if (!menu) {
    menu = ss.insertSheet(POS.menu);
    if (menu.getMaxColumns() < 7) menu.insertColumnsAfter(menu.getMaxColumns(), 7 - menu.getMaxColumns());
    menu.getRange(1,1,1,7).setValues([POS.menuHeaders]).setFontWeight('bold');
    menu.getRange(2,1,POS.seed.length,7).setValues(POS.seed.map(r => r.concat(['', POS.seedArt[r[0]] || 'americano'])));
    menu.setFrozenRows(1);
  } else {
    ensureMenuCols_(menu);
  }
}

/** Edit the values below and run once in the editor; leave email empty to disable email. */
function configurePOS_() {
  const email = Session.getEffectiveUser().getEmail();
  const pin = '';   // choose at least 4 digits; protects the cancel action in the app
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('อีเมลไม่ถูกต้อง');
  if (pin && !/^\d{4,12}$/.test(pin)) throw new Error('PIN ต้องเป็นตัวเลข 4–12 หลัก');
  const p = PropertiesService.getScriptProperties();
  if (email) p.setProperty('SUMMARY_EMAIL',email); else p.deleteProperty('SUMMARY_EMAIL');
  if (pin) p.setProperty('VOID_PIN_HASH',hash_(pin)); else p.deleteProperty('VOID_PIN_HASH');
}

/** Run once to install a daily trigger, approximately 13:00 Bangkok time. */
function installDailySummary_() {
  if (!PropertiesService.getScriptProperties().getProperty('SUMMARY_EMAIL')) throw new Error('ตั้งอีเมลใน configurePOS ก่อน');
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'sendDailySummary_').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailySummary_').timeBased().everyDays(1).atHour(13).inTimezone(POS.timezone).create();
  PropertiesService.getScriptProperties().setProperty('TRIGGER_OK','1');
}

function hash_(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s)); }
function ss_() { return SpreadsheetApp.openById(POS.spreadsheetId); }
function sheet_(name) { const s = ss_().getSheetByName(name); if (!s) throw new Error('ยังไม่ได้ติดตั้งระบบ: กรุณารัน setupPOS_ ใน Apps Script'); return s; }
function today_() { return Utilities.formatDate(new Date(),POS.timezone,'yyyy-MM-dd'); }
function money_(n) { return Math.round(Number(n) * 100) / 100; }
function safeText_(s,max) { return String(s == null ? '' : s).trim().slice(0,max); }
function sheetText_(s) { s=String(s); return /^[\s]*[=+\-@]/.test(s) ? "'"+s : s; }

/** Adds the image and art columns to an older 5-column menu sheet. Never changes existing menu data. */
function ensureMenuCols_(s) {
  if (s.getMaxColumns() < 7) s.insertColumnsAfter(s.getMaxColumns(), 7 - s.getMaxColumns());
  const h = s.getRange(1,1,1,7).getValues()[0].map(String);
  if (h.slice(0,5).join('|') !== 'id|name|price|category|active') throw new Error('ชีต ARY_POS_Menu หัวตารางไม่ตรง กรุณาตรวจสอบก่อน');
  if (h[5] === 'image' && h[6] === 'art') return;
  s.getRange(1,6,1,2).setValues([['image','art']]).setFontWeight('bold');
  const n = s.getLastRow();
  if (n > 1) {
    const ids = s.getRange(2,1,n-1,1).getValues(), art = s.getRange(2,7,n-1,1).getValues();
    s.getRange(2,7,n-1,1).setValues(ids.map((r,i) => [art[i][0] || POS.seedArt[String(r[0])] || 'americano']));
  }
}
function isActive_(v) { return v === true || String(v).toLowerCase() === 'true'; }
function menuAll_() {
  const s = sheet_(POS.menu); ensureMenuCols_(s);
  const n = s.getLastRow(); if (n < 2) return [];
  const seen = {};
  return s.getRange(2,1,n-1,7).getValues().filter(r => r[0]).map(r => {
    const id = safeText_(r[0],40), name = safeText_(r[1],100), price = Number(r[2]);
    if (!id || !name || !isFinite(price) || price < 0 || seen[id]) throw new Error('ข้อมูลเมนูไม่ถูกต้องหรือรหัสซ้ำ กรุณาตรวจสอบชีตเมนู');
    seen[id] = true;
    const img = String(r[5] || ''), art = String(r[6] || '');
    return {id:id,name:name,price:money_(price),category:safeText_(r[3],60) || 'อื่น ๆ',active:isActive_(r[4]),
      image:/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(img) ? img : '', art:POS.artKeys.indexOf(art) >= 0 ? art : 'americano'};
  });
}
function menu_() { return menuAll_().filter(m => m.active); }

function bootstrap() {
  const props = PropertiesService.getScriptProperties();
  const result = {menu:menuAll_(),today:today_(),warning:'',pinConfigured:!!props.getProperty('VOID_PIN_HASH'),summaryEmail:props.getProperty('SUMMARY_EMAIL') || ''};
  if (result.summaryEmail && props.getProperty('TRIGGER_OK')) return result; // fast path: nothing to set up
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(5000)) { result.warning='ยังตั้งอีเมลสรุปอัตโนมัติไม่ได้ กรุณาโหลดหน้าใหม่'; return result; }
    let email = props.getProperty('SUMMARY_EMAIL');
    if (!email) {
      email = Session.getEffectiveUser().getEmail();
      if (!email) throw new Error('ไม่พบอีเมลบัญชีเจ้าของ');
      props.setProperty('SUMMARY_EMAIL',email);
    }
    if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'sendDailySummary_')) {
      ScriptApp.newTrigger('sendDailySummary_').timeBased().everyDays(1).atHour(13).inTimezone(POS.timezone).create();
    }
    props.setProperty('TRIGGER_OK','1');
    result.summaryEmail = email;
  } catch(e) { result.warning='ยังตั้งสรุปทางอีเมลอัตโนมัติไม่ได้: '+e.message+' กรุณาตรวจสอบสิทธิ์และโหลดหน้าใหม่'; }
  finally { if (lock.hasLock()) lock.releaseLock(); }
  return result;
}

/** Only deploy to the owner account. PIN guards accidental voids, not account access. */
function setPOSPin(currentPin,newPin) {
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const props=PropertiesService.getScriptProperties(), existing=props.getProperty('VOID_PIN_HASH');
    if (existing && hash_(String(currentPin||''))!==existing) throw new Error('PIN เดิมไม่ถูกต้อง');
    if (!/^\d{4,12}$/.test(String(newPin||''))) throw new Error('PIN ใหม่ต้องเป็นตัวเลข 4–12 หลัก');
    props.setProperty('VOID_PIN_HASH',hash_(String(newPin))); return true;
  } finally { lock.releaseLock(); }
}

/** Add a new menu item (no id) or edit one (id given). Returns the saved item. */
function saveMenuItem(input) {
  if (!input) throw new Error('ข้อมูลเมนูไม่ถูกต้อง');
  const name = safeText_(input.name,60), price = Number(input.price), category = safeText_(input.category,40) || 'อื่น ๆ';
  if (!name) throw new Error('กรุณาใส่ชื่อเมนู');
  if (!isFinite(price) || price < 0 || price > 99999) throw new Error('ราคาไม่ถูกต้อง');
  const art = POS.artKeys.indexOf(input.art) >= 0 ? input.art : 'americano';
  const image = String(input.image || '');
  if (image && (image.length > POS.maxImage || !/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(image))) throw new Error('รูปไม่ถูกต้องหรือใหญ่เกินไป ลองถ่ายหรือเลือกรูปใหม่');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = sheet_(POS.menu); ensureMenuCols_(s);
    const n = s.getLastRow(), ids = n > 1 ? s.getRange(2,1,n-1,1).getValues().map(r => String(r[0])) : [];
    let id = safeText_(input.id,40);
    if (id) {
      const i = ids.indexOf(id); if (i < 0) throw new Error('ไม่พบเมนูนี้');
      const active = s.getRange(i+2,5).getValue();
      s.getRange(i+2,2,1,6).setValues([[sheetText_(name),money_(price),sheetText_(category),active,image,art]]);
    } else {
      let max = 0; ids.forEach(x => { const m = /^M(\d+)$/.exec(x); if (m) max = Math.max(max, Number(m[1])); });
      id = 'M' + (max + 1 < 10 ? '0' : '') + (max + 1);
      s.appendRow([id,sheetText_(name),money_(price),sheetText_(category),true,image,art]);
    }
    SpreadsheetApp.flush();
    return {id:id,name:name,price:money_(price),category:category,active:true,image:image,art:art};
  } finally { lock.releaseLock(); }
}

/** Hide (active=false) or show a menu item. Past sales are kept. */
function setMenuActive(id,active) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = sheet_(POS.menu), n = s.getLastRow(); if (n < 2) throw new Error('ไม่พบเมนูนี้');
    const i = s.getRange(2,1,n-1,1).getValues().findIndex(r => String(r[0]) === String(id));
    if (i < 0) throw new Error('ไม่พบเมนูนี้');
    s.getRange(i+2,5).setValue(!!active); SpreadsheetApp.flush(); return true;
  } finally { lock.releaseLock(); }
}

/** One order is one sheet row, so a retry cannot produce a partial order. Payment is always recorded as transfer. */
function saveOrder(input) {
  if (!input || !/^[A-Za-z0-9][A-Za-z0-9-]{15,79}$/.test(String(input.id || ''))) throw new Error('รหัสคำสั่งซื้อไม่ถูกต้อง');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s = sheet_(POS.orders), last = s.getLastRow();
    if (last > 1) {
      const ids = s.getRange(2,1,last-1,1).getValues();
      const i = ids.findIndex(r => String(r[0]) === input.id);
      if (i >= 0) { const r=s.getRange(i+2,1,1,10).getValues()[0]; return {id:String(r[0]),total:Number(r[4]),duplicate:true,status:String(r[7])}; }
    }
    if (!Array.isArray(input.items) || input.items.length < 1 || input.items.length > 100) return {rejected:true,message:'รายการสินค้าไม่ถูกต้อง'};
    const catalog = {}; menu_().forEach(m => catalog[m.id]=m);
    let total=0, validationError='';
    const items=input.items.map(x => {
      const m=catalog[x.id], qty=Number(x.qty), sweetness=safeText_(x.sweetness,20), note=safeText_(x.note,100);
      if (!m || money_(x.price) !== m.price || !Number.isInteger(qty) || qty < 1 || qty > 100) { validationError='เมนูหรือราคาเปลี่ยน หรือจำนวนไม่ถูกต้อง กรุณาโหลดหน้าใหม่'; return null; }
      if (!['ปกติ','ไม่หวาน','หวานน้อย','หวานมาก'].includes(sweetness)) { validationError='ระดับความหวานไม่ถูกต้อง'; return null; }
      total=money_(total + m.price*qty); return {id:m.id,name:m.name,price:m.price,qty:qty,sweetness:sweetness,note:note};
    });
    if (validationError) return {rejected:true,message:validationError};
    const now=new Date(), note=safeText_(input.note,200);
    s.appendRow([input.id,now,Utilities.formatDate(now,POS.timezone,'yyyy-MM-dd'),'โอน',total,JSON.stringify(items),sheetText_(note),'PAID','','']);
    SpreadsheetApp.flush();
    return {id:input.id,total:total,duplicate:false,status:'PAID'};
  } finally { lock.releaseLock(); }
}

function voidOrder(id,pin,reason) {
  const p=PropertiesService.getScriptProperties().getProperty('VOID_PIN_HASH');
  if (!p) throw new Error('ยังไม่ได้ตั้ง PIN ยกเลิกบิล ไปที่แท็บ ตั้งค่า');
  if (hash_(String(pin)) !== p) throw new Error('PIN ไม่ถูกต้อง');
  reason=safeText_(reason,200); if (!reason) throw new Error('กรุณาระบุเหตุผล');
  const lock=LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const s=sheet_(POS.orders), n=s.getLastRow(); if(n<2) throw new Error('ไม่พบบิล');
    const rows=s.getRange(2,1,n-1,10).getValues(), i=rows.findIndex(r => String(r[0]) === String(id));
    if(i<0) throw new Error('ไม่พบบิล'); if(rows[i][7] !== 'PAID') throw new Error('บิลนี้ถูกยกเลิกแล้ว');
    s.getRange(i+2,8,1,3).setValues([['VOID',sheetText_(reason),new Date()]]); SpreadsheetApp.flush(); return true;
  } finally { lock.releaseLock(); }
}

function dayList_(from,to) {
  const a = from.split('-').map(Number), b = to.split('-').map(Number);
  let t = Date.UTC(a[0],a[1]-1,a[2]); const end = Date.UTC(b[0],b[1]-1,b[2]), out = [];
  if ((end - t) / 86400000 > 92) throw new Error('เลือกช่วงวันที่ได้ไม่เกิน 93 วัน');
  for (; t <= end; t += 86400000) out.push(new Date(t).toISOString().slice(0,10));
  return out;
}

function report(from,to) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) throw new Error('ช่วงวันที่ไม่ถูกต้อง');
  const days = dayList_(from,to), byDayMap = {};
  const byDay = days.map(d => (byDayMap[d] = {date:d,revenue:0,cups:0,orders:0}));
  const s=sheet_(POS.orders), n=s.getLastRow(), rows=n>1?s.getRange(2,1,n-1,10).getValues():[];
  const result={from:from,to:to,revenue:0,cups:0,orders:0,voids:0,byDay:byDay,products:[],recent:[]};
  const products={};
  rows.forEach(r=>{
    const day = r[2] instanceof Date ? Utilities.formatDate(r[2],POS.timezone,'yyyy-MM-dd') : String(r[2]);
    if(day<from || day>to) return;
    let items; try { items=JSON.parse(r[5]); if(!Array.isArray(items)) throw Error(); } catch(e) { throw new Error('ข้อมูลรายการขายเสียหาย: '+r[0]); }
    const status=String(r[7]); if(status !== 'PAID' && status !== 'VOID') throw new Error('สถานะบิลไม่ถูกต้อง: '+r[0]);
    const entry={id:String(r[0]),time:r[1] instanceof Date?Utilities.formatDate(r[1],POS.timezone,'dd/MM HH:mm'):String(r[1]),total:Number(r[4]),status:status,cups:items.reduce((a,x)=>a+Number(x.qty),0),items:items,note:String(r[6] || ''),reason:String(r[8] || '')};
    result.recent.push(entry);
    if(status==='VOID') { result.voids++; return; }
    result.orders++; result.cups+=entry.cups; result.revenue=money_(result.revenue+entry.total);
    const d = byDayMap[day]; d.orders++; d.cups+=entry.cups; d.revenue=money_(d.revenue+entry.total);
    items.forEach(x=>{ const key=x.id+'|'+x.name; if(!products[key]) products[key]={id:String(x.id),name:x.name,qty:0,revenue:0}; products[key].qty+=Number(x.qty); products[key].revenue=money_(products[key].revenue+Number(x.qty)*Number(x.price)); });
  });
  result.products=Object.values(products).sort((a,b)=>b.qty-a.qty || b.revenue-a.revenue);
  result.recent.reverse(); if (result.recent.length > 200) result.recent.length = 200;
  return result;
}
function sendDailySummary_() {
  const email=PropertiesService.getScriptProperties().getProperty('SUMMARY_EMAIL'); if(!email) throw new Error('ยังไม่ได้ตั้งอีเมล');
  const r=report(today_(),today_());
  const lines=['สรุปยอด A-RY Coffee วันที่ '+r.from,'ยอดขาย '+r.revenue.toFixed(2)+' บาท | '+r.cups+' แก้ว | '+r.orders+' บิล','ยกเลิก '+r.voids+' บิล','','เมนูขายดี'];
  r.products.forEach((p,i)=>lines.push((i+1)+'. '+p.name+' — '+p.qty+' แก้ว / '+p.revenue.toFixed(2)+' บาท'));
  MailApp.sendEmail(email,'สรุปยอด A-RY Coffee '+r.from,lines.join('\n'));
}
