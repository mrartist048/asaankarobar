/* SHOP MANAGER - ADMIN SCRIPT (owner logins + packages live in this Google Sheet)
   Setup: Extensions -> Apps Script -> paste this -> Deploy -> Web app (Execute as: Me, Who has access: Anyone).
   Updating an existing deployment: Deploy -> Manage deployments -> Edit -> New version (the URL stays the same).
   Sheet "Shopkeepers" is created automatically. Columns you can edit by hand:
     Status : active / suspended / removed  (suspended/removed = owner AND staff cannot log in)
     Expiry : package end date like 31/12/2026 (empty = never expires). After that day the owner AND all staff are blocked.
   Staff accounts are NOT stored here: each owner keeps staff in his own Google Sheet. */
const AUTO_APPROVE = true;  // true = new shops can log in immediately
const TRIAL_DAYS = 0;       // >0 = new shops get this many days; 0 = no expiry until you set one
const H = ["Username","Salt","Hash","Name","Shop","Role","Owner","Token","Status","ShopURL","ShopKey","Created","Expiry"];

function doGet(){ return out({ok:1}); }
function doPost(e){
  const L = LockService.getScriptLock(); L.waitLock(25000);
  try { return out(run(JSON.parse(e.postData.contents))); }
  catch (x) { return out({err:"server", msg:String(x.message||x)}); }
  finally { L.releaseLock(); }
}
function out(j){ return ContentService.createTextOutput(JSON.stringify(j)).setMimeType(ContentService.MimeType.JSON); }

function S(){
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName("Shopkeepers");
  if (!s) {
    s = ss.insertSheet("Shopkeepers"); s.appendRow(H); s.setFrozenRows(1);
    s.getRange(1,1,1,H.length).setFontWeight("bold").setBackground("#0d5c4b").setFontColor("#ffffff");
    s.getRange(2,1,999,H.length).setNumberFormat("@");
    tidySheet();
  } else if (s.getLastColumn() < H.length) {
    s.getRange(1,1,1,H.length).setValues([H]); s.getRange(2,H.length,999,1).setNumberFormat("@");
    tidySheet();
  }
  return s;
}
// Makes the sheet easy to read: Username column stays visible while scrolling, technical columns are hidden,
// and rows turn RED when the package is expired and ORANGE when it ends within 7 days.
// Run it once by hand: Apps Script editor -> select "tidySheet" -> Run.
function tidySheet(){
  const s = S(), n = H.length;
  s.getRange(1,1,1,n).setFontWeight("bold").setBackground("#0d5c4b").setFontColor("#ffffff");
  s.setFrozenRows(1); s.setFrozenColumns(1);
  [2,3,8,11].forEach(c => s.hideColumns(c)); // Salt, Hash, Token, ShopKey (never needed by hand)
  s.getRange(2,12,998,2).setNumberFormat("@");
  const lr = s.getLastRow();
  if (lr>1) {
    const rg = s.getRange(2,12,lr-1,2), v = rg.getValues();
    rg.setValues(v.map(r => r.map((c,i) => { const d = c instanceof Date ? c : parseD(c); return d ? fmtD(d) : String(c); })));
  }
  const r = s.getRange(2,1,998,n);
  s.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(LEN($M2)=10,DATE(VALUE(RIGHT($M2,4)),VALUE(MID($M2,4,2)),VALUE(LEFT($M2,2)))<TODAY())').setBackground("#f4c7c3").setRanges([r]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(LEN($M2)=10,DATE(VALUE(RIGHT($M2,4)),VALUE(MID($M2,4,2)),VALUE(LEFT($M2,2)))>=TODAY(),DATE(VALUE(RIGHT($M2,4)),VALUE(MID($M2,4,2)),VALUE(LEFT($M2,2)))-TODAY()<=7)').setBackground("#fce8b2").setRanges([r]).build()
  ]);
}
function rows(){
  return S().getDataRange().getValues().slice(1).map((r,i) => {
    const x = {_r:i+2};
    H.forEach((h,k) => { const c = r[k]; x[h] = c instanceof Date ? fmtD(c) : String(c==null?"":c); });
    return x;
  }).filter(x => x.Username);
}
const find = u => rows().find(x => x.Username.toLowerCase() == String(u||"").trim().toLowerCase());
const rnd = n => (Utilities.getUuid()+Utilities.getUuid()).replace(/-/g,"").slice(0,n);
const hash = (salt,pw) => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt+pw));

const fmtD = d => Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
function parseD(t){ // accepts 05/10/2026 or 2026-10-05, returns the END of that day (or null)
  t = String(t==null?"":t).trim(); let d,m,y,k;
  if ((k = t.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})/))) { d=+k[1]; m=+k[2]; y=+k[3]; }
  else if ((k = t.match(/^(\d{4})-(\d{2})-(\d{2})/))) { y=+k[1]; m=+k[2]; d=+k[3]; }
  else return null;
  const x = new Date(y,m-1,d,23,59,59); return x.getMonth()==m-1 ? x : null;
}
function expMs(x){ const d = parseD(x.Expiry); return d ? d.getTime() : 0; }
// Typing an expiry by hand (e.g. 5/10/2026) is auto-formatted to 05/10/2026.
function onEdit(e){
  try {
    const r = e.range; if (r.getSheet().getName()!="Shopkeepers" || r.getColumn()!=13 || r.getRow()<2) return;
    const v = r.getValue(); if (v==="") return;
    const d = v instanceof Date ? v : parseD(v);
    if (d) { r.setNumberFormat("@"); r.setValue(fmtD(d)); }
  } catch (x) {}
}
function block(x){
  const st = String(x.Status).toLowerCase().trim();
  if (st=="suspended" || st=="removed") return st;
  const e = expMs(x); if (e && Date.now()>e) return "expired";
  return "";
}
function prof(x){
  if (x.Role!="owner") return {err:"bad"};
  const b = block(x); if (b) return {err:"inactive", status:b};
  return {t:x.Token, name:x.Name, role:"owner", shop:x.Shop, url:x.ShopURL, key:x.ShopKey, exp:expMs(x)};
}
function me(b){
  const x = find(b.u);
  if (!x || x.Token != b.t) throw new Error("Session expired. Please log in again.");
  return x;
}
function run(b){
  if (b.a=="register"){
    const u = String(b.u||"").trim().toLowerCase();
    if (!/^[a-z0-9_-]{3,20}$/.test(u) || String(b.pw||"").length<6) throw new Error("Invalid username or password.");
    if (find(u)) return {err:"username"};
    const salt = rnd(16);
    const exp = TRIAL_DAYS>0 ? fmtD(new Date(Date.now()+TRIAL_DAYS*864e5)) : "";
    S().appendRow([u,salt,hash(salt,b.pw),b.name||"",b.shop||"","owner","",rnd(32),AUTO_APPROVE?"active":"pending","","",fmtD(new Date()),exp]);
    return {ok:1};
  }
  if (b.a=="login"){
    const x = find(b.u);
    if (!x || x.Hash != hash(x.Salt,String(b.pw||""))) return {err:"bad"};
    return prof(x);
  }
  if (b.a=="shop"){ // used by staff before they log in: tells the app which sheet belongs to this shop
    const o = find(b.o);
    if (!o || o.Role!="owner") return {err:"nofound", msg:"Shop ID not found. Ask your owner for the correct Shop ID."};
    const bl = block(o); if (bl) return {err:"inactive", status:bl};
    if (!o.ShopURL) return {err:"nofound", msg:"This shop has not connected its Google Sheet yet."};
    return {url:o.ShopURL, shop:o.Shop, exp:expMs(o)};
  }
  const x = me(b);
  if (b.a=="session") return prof(x);
  if (b.a=="setShop"){ S().getRange(x._r,10,1,2).setValues([[String(b.url),String(b.key)]]); return {ok:1}; }
  throw new Error("Unknown action");
}
