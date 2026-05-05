// ============================================================
// GREEN SALON — BILLING SYSTEM v9.0
// Owner: Harsha | Developer: Shebin K Babu
// ============================================================
// SETUP:
// 1. Paste this into Apps Script (Extensions → Apps Script → clear all → paste)
// 2. Replace the 4 SHEET ID placeholders below
// 3. Run: firstTimeSetup  → approve all permissions
// 4. Run: setupTriggers
// 5. Deploy → New Deployment → Web App → Execute as: Me → Anyone
// 6. Copy /exec URL → paste into owner-panel HTML as API_URL
// ============================================================

const MASTER_SHEET_ID = "1YUfe5XL6yFirq3CNfjHFHcqORMMcYAW9gI7pPnJyXuY";
const OWNER_PASSWORD  = "harsha@greensalon2026";
const BRANCH_SHEETS   = {
  "branch1": "1fRTEOMjhjqZ0P3pVfagC3hNdFeE3rL1LD_T98YraWGQ",
  "branch2": "14sPzYtF13bYldvYIC0xtGy0CbjKyWdpEVzCLQrKV5tA",
  "branch3": "1nQ2svtVxaKhCGGDltKS0u2iGooGlSkVboNj1OnZk9rc",
};

// Sheet colours — match client's green Excel theme
const C_DARK  = "#1a5c38";
const C_MED   = "#2d8653";
const C_WHITE = "#ffffff";
const C_ALT   = "#e8f5ee";
const C_RAW   = "#0f172a";

// ── ROUTER ───────────────────────────────────────────────────
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      // Auth
      case "ownerLogin":        return R(ownerLogin(d));
      // Branches
      case "getBranches":       return R(getBranches());
      case "addBranch":         return R(addBranch(d));
      case "removeBranch":      return R(removeBranch(d));
      case "recoverBranch":     return R(recoverBranch(d));
      case "renameBranch":      return R(renameBranch(d));
      // Staff
      case "getStaffAdmin":     return R(getStaffAdmin(d));
      case "addStaff":          return R(addStaff(d));
      case "removeStaff":       return R(removeStaff(d));
      case "renameStaff":       return R(renameStaff(d));
      case "updateStaffComm":   return R(updateStaffComm(d));
      // Services / Products lists
      case "updateServices":    return R(updateServices(d));
      case "updateProducts":    return R(updateProducts(d));
      // Entries
      case "submitEntry":       return R(submitEntry(d));
      case "submitProduct":     return R(submitProduct(d));   // now GLOBAL (no staffId needed)
      case "submitExpense":     return R(submitExpense(d));   // now GLOBAL (no staffId needed)
      case "getMyEntries":      return R(getMyEntries(d));
      case "getTodayAll":       return R(getTodayAll(d));
      case "getBranchSummary":  return R(getBranchSummary(d)); // NEW: landing page card
      case "deleteEntry":       return R(deleteEntry(d));
      case "deleteProduct":     return R(deleteProduct(d));
      case "deleteExpense":     return R(deleteExpense(d));
      // Reporting
      case "getMonthSummary":   return R(getMonthSummary(d));
      case "setReportEmails":   return R(setReportEmails(d));   // NEW: multi-email
      case "getReportEmails":   return R(getReportEmails(d));   // NEW: multi-email
      case "sendManualReport":  return R(sendManualReport(d));  // NEW: manual trigger
      default: return E("Unknown action: " + d.action);
    }
  } catch (ex) { return E(ex.message); }
}

function doGet(e) {
  try {
    const a = e.parameter.action, bid = e.parameter.branchId;
    switch (a) {
      case "getStaff":    return R(getStaff(bid));
      case "getServices": return R(getServices(bid));
      case "getProducts": return R(getProducts(bid));
      case "getBranches": return R(getBranches());
      default: return E("Unknown action: " + a);
    }
  } catch (ex) { return E(ex.message); }
}

function R(d) { return ContentService.createTextOutput(JSON.stringify({success:true,...d})).setMimeType(ContentService.MimeType.JSON); }
function E(m) { return ContentService.createTextOutput(JSON.stringify({success:false,error:m})).setMimeType(ContentService.MimeType.JSON); }

// ── SHEET HELPERS ─────────────────────────────────────────────
function masterTab(name) {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function branchSS(branchId) {
  const sid = BRANCH_SHEETS[branchId];
  if (sid && !sid.includes("_SHEET_ID_HERE")) return SpreadsheetApp.openById(sid);
  const rows = masterTab("Branches").getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === branchId && rows[i][4] !== false && rows[i][4] !== "FALSE")
      return SpreadsheetApp.openById(rows[i][3]);
  }
  throw new Error("Branch not found: " + branchId);
}

function branchTab(branchId, tabName) {
  const ss = branchSS(branchId);
  return ss.getSheetByName(tabName) || ss.insertSheet(tabName);
}

function getExistingTab(branchId, tabName) {
  return branchSS(branchId).getSheetByName(tabName);
}

function todayStr() {
  return Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MM-yyyy");
}
function nowIST() {
  return Utilities.formatDate(new Date(), "Asia/Kolkata", "dd-MMM-yyyy HH:mm:ss");
}
function monthName() {
  const ist = new Date(new Date().toLocaleString("en-US", {timeZone:"Asia/Kolkata"}));
  return ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][ist.getMonth()]+" "+ist.getFullYear();
}
function hdrStyle(rng, bg, fg) {
  rng.setBackground(bg).setFontColor(fg).setFontWeight("bold").setHorizontalAlignment("center");
}
function branchDisplayName(branchId) {
  try {
    const rows = masterTab("Branches").getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) { if (rows[i][0]===branchId) return rows[i][1]; }
  } catch(ex) {}
  return "Green Salon";
}

// ── FIRST TIME SETUP ──────────────────────────────────────────
function firstTimeSetup() {
  // Master → Branches tab
  const bsh = masterTab("Branches");
  if (bsh.getLastRow()===0) {
    bsh.appendRow(["BranchID","Name","Location","SheetID","Active","CreatedAt","DeletedAt"]);
    hdrStyle(bsh.getRange(1,1,1,7), C_RAW, C_WHITE);
  }
  // Master → Settings tab (now stores 3 email fields)
  const set = masterTab("Settings");
  if (set.getLastRow()===0) {
    set.appendRow(["BranchID","Email1","Email2","Email3","UpdatedAt"]);
    hdrStyle(set.getRange(1,1,1,5), C_RAW, C_WHITE);
  }
  const defs = [
    {id:"branch1", name:"Branch 1", loc:"JC Nagar"},
    {id:"branch2", name:"Branch 2", loc:"Koramangala"},
    {id:"branch3", name:"Branch 3", loc:"Indiranagar"},
  ];
  const existing = bsh.getDataRange().getValues().map(r=>r[0]);
  defs.forEach(b=>{
    const sid=BRANCH_SHEETS[b.id];
    if(sid.includes("_SHEET_ID_HERE")){Logger.log("⚠️ "+b.id+" not set");return;}
    if(!existing.includes(b.id)) bsh.appendRow([b.id,b.name,b.loc,sid,true,nowIST(),""]);
    try{initBranch(sid,b.name);}catch(ex){Logger.log("❌ "+b.name+": "+ex.message);}
  });
  Logger.log("✅ Setup done. Next: run setupTriggers(), then Deploy as Web App.");
}

function initBranch(sheetId, branchName) {
  const ss = SpreadsheetApp.openById(sheetId);
  // Staff
  let st = ss.getSheetByName("Staff") || ss.insertSheet("Staff");
  if(st.getLastRow()===0){
    st.appendRow(["ID","Name","PIN","Commission%","HasCommission","PhotoURL","Active"]);
    hdrStyle(st.getRange(1,1,1,7),C_RAW,C_WHITE);
    st.appendRow(["S001","Staff 1","1111",40,true,"",true]);
    st.appendRow(["S002","Staff 2","2222",40,true,"",true]);
    st.appendRow(["S003","Staff 3","3333",35,true,"",true]);
  }
  // Services
  let sv = ss.getSheetByName("Services") || ss.insertSheet("Services");
  if(sv.getLastRow()===0){
    sv.appendRow(["ServiceName","Price","Active"]); hdrStyle(sv.getRange(1,1,1,3),C_RAW,C_WHITE);
    [["Haircut",150],["Shave",80],["Facial",300],["Hair Colour",500],["Head Massage",100],["Beard Trim",60],["Threading",40],["Waxing",200]].forEach(r=>sv.appendRow([r[0],r[1],true]));
  }
  // Products
  let pd = ss.getSheetByName("Products") || ss.insertSheet("Products");
  if(pd.getLastRow()===0){
    pd.appendRow(["ProductName","Price","Active"]); hdrStyle(pd.getRange(1,1,1,3),C_RAW,C_WHITE);
    [["Shampoo",200],["Hair Oil",150],["Conditioner",180],["Hair Serum",250]].forEach(r=>pd.appendRow([r[0],r[1],true]));
  }
  // Entries (service log — staff-level)
  let en = ss.getSheetByName("Entries") || ss.insertSheet("Entries");
  if(en.getLastRow()===0){en.appendRow(["RowID","Timestamp","Date","StaffID","StaffName","Service","Amount","Tip","Payment","CommApplies","Flagged"]);hdrStyle(en.getRange(1,1,1,11),C_RAW,C_WHITE);}
  // ProductSales (GLOBAL — branch level, staffId may be blank)
  let ps = ss.getSheetByName("ProductSales") || ss.insertSheet("ProductSales");
  if(ps.getLastRow()===0){ps.appendRow(["RowID","Timestamp","Date","StaffID","StaffName","Product","Amount","Payment","Flagged"]);hdrStyle(ps.getRange(1,1,1,9),C_RAW,C_WHITE);}
  // Expenses (GLOBAL — branch level, staffId may be blank)
  let ex = ss.getSheetByName("Expenses") || ss.insertSheet("Expenses");
  if(ex.getLastRow()===0){ex.appendRow(["RowID","Timestamp","Date","StaffID","StaffName","Description","Amount","Payment","Flagged"]);hdrStyle(ex.getRange(1,1,1,9),C_RAW,C_WHITE);}
  // Daily + Monthly tabs
  buildDailyTab(ss, branchName);
  buildMonthlyTab(ss, branchName, monthName());
  Logger.log("✅ "+branchName+" initialized");
}

// ── DAILY TAB ─────────────────────────────────────────────────
function buildDailyTab(ss, branchName) {
  let sh = ss.getSheetByName("Daily") || ss.insertSheet("Daily");
  if(sh.getLastRow()>0) return sh;
  const names=activeStaffNames(ss);
  const headers=[];
  names.forEach(n=>{headers.push(n);headers.push(n+" Time");});
  headers.push("Product"); headers.push("Product Time");
  sh.appendRow(headers);
  hdrStyle(sh.getRange(1,1,1,headers.length),C_DARK,C_WHITE);
  sh.setFrozenRows(1);
  for(let c=1;c<=headers.length;c++) sh.setColumnWidth(c,c%2===0?185:78);
  return sh;
}

function dailyColMap(ss) {
  const sh=ss.getSheetByName("Daily"); if(!sh||sh.getLastRow()===0) return {};
  const h=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const m={};
  h.forEach((v,i)=>{if(v&&!String(v).endsWith(" Time")) m[String(v)]={amt:i+1,time:i+2};});
  return m;
}

function ensureStaffInDaily(ss, name) {
  const sh=ss.getSheetByName("Daily"); if(!sh) return;
  if(dailyColMap(ss)[name]) return;
  const last=sh.getLastColumn();
  sh.insertColumnsAfter(last-2,2);
  sh.getRange(1,last-1).setValue(name); sh.getRange(1,last).setValue(name+" Time");
  hdrStyle(sh.getRange(1,last-1,1,2),C_DARK,C_WHITE);
  sh.setColumnWidth(last-1,78); sh.setColumnWidth(last,185);
}

function writeDailyEntry(ss, staffName, amount, payment, isProduct) {
  const sh=ss.getSheetByName("Daily"); if(!sh) return;
  if(!isProduct) ensureStaffInDaily(ss,staffName);
  const map=dailyColMap(ss);
  const info=map[isProduct?"Product":staffName]; if(!info) return;
  const val=amount+(payment==="Cash"?"C":"P");
  const ts=Utilities.formatDate(new Date(),"Asia/Kolkata","dd-MMM-yyyy hh:mm:ss a");
  const data=sh.getRange(2,info.amt,Math.max(sh.getLastRow(),2),1).getValues();
  let row=2;
  for(let r=0;r<data.length;r++){if(!data[r][0]){row=r+2;break;}if(r===data.length-1)row=data.length+2;}
  sh.getRange(row,info.amt).setValue(val).setHorizontalAlignment("center");
  sh.getRange(row,info.time).setValue(ts);
}

// ── MONTHLY TAB ───────────────────────────────────────────────
function buildMonthlyTab(ss, branchName, tabName) {
  if(ss.getSheetByName(tabName)) return ss.getSheetByName(tabName);
  const sh=ss.insertSheet(tabName);
  drawMonthlyFrame(sh,branchName,activeStaffNames(ss));
  return sh;
}

function drawMonthlyFrame(sh, branchName, staffNames) {
  const cols=["Date",...staffNames,"Extra","Total","Product","Expenses","Commission","Online","Cash","Difference"];
  const nc=cols.length;
  sh.getRange(1,1,1,nc).merge().setValue(branchName).setBackground(C_DARK).setFontColor(C_WHITE).setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1,36);
  sh.getRange(2,1,1,nc).setValues([cols]);
  hdrStyle(sh.getRange(2,1,1,nc),C_MED,C_WHITE);
  sh.setRowHeight(2,26); sh.setFrozenRows(2);
  sh.setColumnWidth(1,110);
  for(let c=2;c<=nc;c++) sh.setColumnWidth(c,90);
}

function monthColMap(sh) {
  if(!sh||sh.getLastRow()<2) return {};
  const h=sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0];
  const m={}; h.forEach((v,i)=>{if(v)m[String(v)]=i+1;}); return m;
}

function ensureStaffInMonthly(sh, staffName, branchName) {
  const h=sh.getRange(2,1,1,sh.getLastColumn()).getValues()[0];
  if(h.includes(staffName)) return;
  const ei=h.indexOf("Extra"); if(ei<0) return;
  sh.insertColumnBefore(ei+1);
  sh.getRange(2,ei+1).setValue(staffName);
  hdrStyle(sh.getRange(2,ei+1,1,1),C_MED,C_WHITE);
  sh.setColumnWidth(ei+1,90);
  const nc=sh.getLastColumn();
  sh.getRange(1,1,1,nc).merge().setValue(branchName).setBackground(C_DARK).setFontColor(C_WHITE).setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");
  for(let r=3;r<=sh.getLastRow();r++){if(String(sh.getRange(r,1).getValue())!=="TOTAL")sh.getRange(r,ei+1).setValue(0);}
}

// KEY: updateMonthly — fixed appendRow bug + correct per-type routing
function updateMonthly(branchId, entryType, staffName, svcAmt, tipAmt, payment, prodAmt, expAmt) {
  const ss=branchSS(branchId); const bn=branchDisplayName(branchId); const tab=monthName();
  let sh=ss.getSheetByName(tab); if(!sh) sh=buildMonthlyTab(ss,bn,tab);
  if(entryType==="service"&&staffName&&svcAmt>0) ensureStaffInMonthly(sh,staffName,bn);
  const cm=monthColMap(sh); const dt=todayStr(); const nc=sh.getLastColumn();
  const allVals=sh.getLastRow()>=3?sh.getRange(3,1,sh.getLastRow()-2,1).getValues():[];
  let dr=-1,tr=-1;
  allVals.forEach((row,idx)=>{const v=String(row[0]).trim();const r=idx+3;if(v===dt)dr=r;if(v==="TOTAL")tr=r;});
  if(dr<0){
    const zeroRow=new Array(nc).fill(0); zeroRow[0]=dt;
    if(tr>0){sh.insertRowBefore(tr);sh.getRange(tr,1,1,nc).setValues([zeroRow]);dr=tr;}
    else{sh.appendRow(zeroRow);dr=sh.getLastRow();}
    const bg=dr%2===0?C_WHITE:C_ALT;
    sh.getRange(dr,1,1,nc).setBackground(bg).setHorizontalAlignment("center");
    sh.getRange(dr,1).setHorizontalAlignment("left");
  }
  function addVal(colKey,val){if(!colKey||!val||Number(val)<=0)return;const col=cm[colKey];if(!col)return;sh.getRange(dr,col).setValue((Number(sh.getRange(dr,col).getValue())||0)+Number(val));}
  if(entryType==="service"){
    if(svcAmt>0)addVal(staffName,svcAmt);
    if(tipAmt>0)addVal("Extra",tipAmt);
    const money=(svcAmt||0)+(tipAmt||0);
    if(money>0)addVal(payment==="Cash"?"Cash":"Online",money);
  }else if(entryType==="product"){
    if(prodAmt>0){addVal("Product",prodAmt);addVal(payment==="Cash"?"Cash":"Online",prodAmt);}
  }else if(entryType==="expense"){
    if(expAmt>0)addVal("Expenses",expAmt);
  }
  recalcRow(sh,dr,cm,nc); rebuildTotal(sh,nc);
}

function recalcRow(sh,row,cm,nc){
  const FIXED=new Set(["Date","Extra","Total","Product","Expenses","Commission","Online","Cash","Difference"]);
  const h=sh.getRange(2,1,1,nc).getValues()[0]; let sum=0;
  h.forEach((v,i)=>{if(v&&!FIXED.has(String(v)))sum+=Number(sh.getRange(row,i+1).getValue())||0;});
  if(cm["Total"])      sh.getRange(row,cm["Total"]).setValue(sum);
  if(cm["Commission"]) sh.getRange(row,cm["Commission"]).setValue(Math.round(sum*0.40));
  const onl=cm["Online"]?(Number(sh.getRange(row,cm["Online"]).getValue())||0):0;
  const csh=cm["Cash"]  ?(Number(sh.getRange(row,cm["Cash"]).getValue())||0):0;
  if(cm["Difference"])  sh.getRange(row,cm["Difference"]).setValue(onl+csh-sum);
}

function rebuildTotal(sh,nc){
  const last=sh.getLastRow(); let tr=-1;
  for(let r=3;r<=last;r++){if(String(sh.getRange(r,1).getValue())==="TOTAL"){tr=r;break;}}
  const sums=new Array(nc).fill(0);
  const endR=tr>0?tr:last+1;
  for(let r=3;r<endR;r++){const v=sh.getRange(r,1,1,nc).getValues()[0];for(let c=1;c<nc;c++)sums[c]+=Number(v[c])||0;}
  const totalRow=["TOTAL",...sums.slice(1)];
  if(tr<0){sh.appendRow(totalRow);tr=sh.getLastRow();}
  else sh.getRange(tr,1,1,nc).setValues([totalRow]);
  hdrStyle(sh.getRange(tr,1,1,nc),C_DARK,C_WHITE);
  sh.getRange(tr,1).setHorizontalAlignment("left");
}

function activeStaffNames(ss){
  const st=ss.getSheetByName("Staff"); if(!st||st.getLastRow()<2) return [];
  return st.getDataRange().getValues().slice(1).filter(r=>r[6]!==false&&r[6]!=="FALSE").map(r=>String(r[1]));
}

// ── TRIGGERS ──────────────────────────────────────────────────
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t=>{
    if(["midnightReset","checkMonthEnd","sendDailyReport","sendMonthlyReport"].includes(t.getHandlerFunction()))
      ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("midnightReset").timeBased().everyDays(1).atHour(23).nearMinute(55).inTimezone("Asia/Kolkata").create();
  ScriptApp.newTrigger("checkMonthEnd").timeBased().everyDays(1).atHour(0).nearMinute(5).inTimezone("Asia/Kolkata").create();
  // Daily report at 11:59 PM IST
  ScriptApp.newTrigger("sendDailyReport").timeBased().everyDays(1).atHour(23).nearMinute(59).inTimezone("Asia/Kolkata").create();
  // Monthly report check runs daily — sends on last day of month
  ScriptApp.newTrigger("sendMonthlyReport").timeBased().everyDays(1).atHour(23).nearMinute(45).inTimezone("Asia/Kolkata").create();
  Logger.log("✅ All triggers set.");
}

function midnightReset(){
  masterTab("Branches").getDataRange().getValues().slice(1).forEach(row=>{
    if(row[4]===false||row[4]==="FALSE")return;
    try{const sh=SpreadsheetApp.openById(row[3]).getSheetByName("Daily");if(sh&&sh.getLastRow()>1)sh.deleteRows(2,sh.getLastRow()-1);}
    catch(ex){Logger.log("reset err "+row[1]+": "+ex.message);}
  });
}

function checkMonthEnd(){
  const tab=monthName();
  masterTab("Branches").getDataRange().getValues().slice(1).forEach(row=>{
    if(row[4]===false||row[4]==="FALSE")return;
    try{const ss=SpreadsheetApp.openById(row[3]);if(!ss.getSheetByName(tab))buildMonthlyTab(ss,row[1],tab);}
    catch(ex){Logger.log("monthEnd err "+row[1]+": "+ex.message);}
  });
}

// ── BRANCHES ──────────────────────────────────────────────────
function getBranches(){
  const rows=masterTab("Branches").getDataRange().getValues();
  const all=rows.slice(1).map(r=>({id:r[0],name:r[1],location:r[2],sheetId:r[3],active:r[4]!==false&&r[4]!=="FALSE",deletedAt:r[6]||""}));
  return{branches:all.filter(b=>b.active),deleted:all.filter(b=>!b.active)};
}

function addBranch(d){
  if(!d.name||!d.sheetId) throw new Error("Name and SheetID required");
  try{SpreadsheetApp.openById(d.sheetId);}catch(ex){throw new Error("Cannot access Sheet — must be same Gmail");}
  const id="branch"+Date.now();
  masterTab("Branches").appendRow([id,d.name,d.location||"",d.sheetId,true,nowIST(),""]);
  initBranch(d.sheetId,d.name);
  return{branchId:id};
}
function removeBranch(d){
  const sh=masterTab("Branches"),rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){sh.getRange(i+1,5).setValue(false);sh.getRange(i+1,7).setValue(nowIST());return{};}}
  throw new Error("Branch not found");
}
function recoverBranch(d){
  const sh=masterTab("Branches"),rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){sh.getRange(i+1,5).setValue(true);sh.getRange(i+1,7).setValue("");return{name:rows[i][1]};}}
  throw new Error("Branch not found");
}
function renameBranch(d){
  if(!d.newName) throw new Error("New name required");
  const sh=masterTab("Branches"),rows=sh.getDataRange().getValues(); let oldName="";
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){oldName=rows[i][1];sh.getRange(i+1,2).setValue(d.newName);break;}}
  try{const ss=branchSS(d.branchId),msh=ss.getSheetByName(monthName());if(msh){const nc=msh.getLastColumn();msh.getRange(1,1,1,nc).merge().setValue(d.newName).setBackground(C_DARK).setFontColor(C_WHITE).setFontWeight("bold").setFontSize(13).setHorizontalAlignment("center").setVerticalAlignment("middle");}}catch(ex){Logger.log("rename hdr: "+ex.message);}
  return{oldName,newName:d.newName};
}

// ── STAFF ─────────────────────────────────────────────────────
function getStaff(bid){return{staff:_staffRows(bid)};}
function getStaffAdmin(d){return{staff:_staffRows(d.branchId)};}
function _staffRows(bid){
  return branchTab(bid,"Staff").getDataRange().getValues().slice(1)
    .filter(r=>r[6]!==false&&r[6]!=="FALSE")
    .map(r=>({id:r[0],name:r[1],pin:r[2],photoUrl:r[5]||"",hasCommission:r[4],commissionPct:Number(r[3])||0}));
}
function addStaff(d){
  const sh=branchTab(d.branchId,"Staff"); const id="S"+Date.now();
  sh.appendRow([id,d.name,d.pin||"0000",d.commissionPct||0,d.hasCommission!==false,d.photoUrl||"",true]);
  const ss=branchSS(d.branchId); const bn=branchDisplayName(d.branchId);
  ensureStaffInDaily(ss,d.name);
  const msh=ss.getSheetByName(monthName()); if(msh) ensureStaffInMonthly(msh,d.name,bn);
  return{staffId:id};
}
function removeStaff(d){
  const sh=branchTab(d.branchId,"Staff"),rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.staffId){sh.getRange(i+1,7).setValue(false);return{};}}
  throw new Error("Staff not found");
}
function renameStaff(d){
  if(!d.newName) throw new Error("New name required");
  const sh=branchTab(d.branchId,"Staff"),rows=sh.getDataRange().getValues(); let oldName="";
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.staffId){oldName=rows[i][1];sh.getRange(i+1,2).setValue(d.newName);return{oldName,newName:d.newName};}}
  throw new Error("Staff not found");
}
function updateStaffComm(d){
  const sh=branchTab(d.branchId,"Staff"),rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.staffId){sh.getRange(i+1,4).setValue(Number(d.commissionPct)||0);sh.getRange(i+1,5).setValue(d.hasCommission===true||d.hasCommission==="true");return{};}}
  throw new Error("Staff not found");
}

// ── SERVICES & PRODUCTS ───────────────────────────────────────
function getServices(bid){return{services:branchTab(bid,"Services").getDataRange().getValues().slice(1).filter(r=>r[2]!==false&&r[2]!=="FALSE").map(r=>({name:r[0],price:r[1]}))}}
function updateServices(d){const sh=branchTab(d.branchId,"Services");if(sh.getLastRow()>1)sh.deleteRows(2,sh.getLastRow()-1);if(d.services&&d.services.length)d.services.forEach(s=>sh.appendRow([s.name,s.price,true]));return{};}
function getProducts(bid){return{products:branchTab(bid,"Products").getDataRange().getValues().slice(1).filter(r=>r[2]!==false&&r[2]!=="FALSE").map(r=>({name:r[0],price:r[1]}))}}
function updateProducts(d){const sh=branchTab(d.branchId,"Products");if(sh.getLastRow()>1)sh.deleteRows(2,sh.getLastRow()-1);if(d.products&&d.products.length)d.products.forEach(p=>sh.appendRow([p.name,p.price,true]));return{};}

// ── SUBMIT ENTRIES ────────────────────────────────────────────
// Service — staff-level
function submitEntry(d){
  const{branchId,staffId,staffName,service,amount,tip,paymentMethod}=d;
  if(!branchId||!staffName||!service||!paymentMethod) throw new Error("Missing required fields");
  const amt=Number(amount)||0,tip2=Number(tip)||0;
  if(amt<=0) throw new Error("Amount must be > 0");
  const sRows=branchTab(branchId,"Staff").getDataRange().getValues(); let comm=true;
  for(let i=1;i<sRows.length;i++){if(sRows[i][0]===staffId){comm=sRows[i][4]===true||sRows[i][4]==="TRUE";break;}}
  const rid="E"+Date.now(),ts=nowIST(),dt=todayStr();
  branchTab(branchId,"Entries").appendRow([rid,ts,dt,String(staffId||""),String(staffName),String(service),amt,tip2,String(paymentMethod),comm,false]);
  writeDailyEntry(branchSS(branchId),staffName,amt,paymentMethod,false);
  updateMonthly(branchId,"service",staffName,amt,tip2,paymentMethod,0,0);
  return{rowId:rid,timestamp:ts};
}

// Product — GLOBAL / branch-level (no staff required)
function submitProduct(d){
  const{branchId,staffId,staffName,product,amount,paymentMethod}=d;
  if(!branchId||!product||!paymentMethod) throw new Error("Missing required fields");
  const amt=Number(amount)||0; if(amt<=0) throw new Error("Amount must be > 0");
  const rid="P"+Date.now(),ts=nowIST(),dt=todayStr();
  branchTab(branchId,"ProductSales").appendRow([rid,ts,dt,String(staffId||"GLOBAL"),String(staffName||"Branch"),String(product),amt,String(paymentMethod),false]);
  writeDailyEntry(branchSS(branchId),"Product",amt,paymentMethod,true);
  updateMonthly(branchId,"product","",0,0,paymentMethod,amt,0);
  return{rowId:rid,timestamp:ts};
}

// Expense — GLOBAL / branch-level (no staff required)
function submitExpense(d){
  const{branchId,staffId,staffName,description,amount,paymentMethod}=d;
  if(!branchId||!description||!paymentMethod) throw new Error("Missing required fields");
  const amt=Number(amount)||0; if(amt<=0) throw new Error("Amount must be > 0");
  const rid="X"+Date.now(),ts=nowIST(),dt=todayStr();
  branchTab(branchId,"Expenses").appendRow([rid,ts,dt,String(staffId||"GLOBAL"),String(staffName||"Branch"),String(description),amt,String(paymentMethod),false]);
  updateMonthly(branchId,"expense","",0,0,paymentMethod,0,amt);
  return{rowId:rid,timestamp:ts};
}

// ── GET ENTRIES ───────────────────────────────────────────────
// getMyEntries — staff-level (only services)
function getMyEntries(d){
  const{branchId,staffId}=d; const dt=todayStr();
  const eR=branchTab(branchId,"Entries").getDataRange().getValues();
  const out=[]; let ta=0,tt=0;
  eR.slice(1).forEach(r=>{
    if(String(r[3])===String(staffId)&&String(r[2])===dt&&r[10]!==true&&r[10]!=="TRUE"){
      out.push({rowId:r[0],timestamp:r[1],service:r[5],amount:r[6],tip:r[7],paymentMethod:r[8]});
      ta+=Number(r[6])||0; tt+=Number(r[7])||0;
    }
  });
  return{entries:out,totalAmount:ta,totalTip:tt};
}

// NEW: getBranchSummary — for landing page card
function getBranchSummary(d){
  const bid=d.branchId; const dt=todayStr();
  const eR=branchTab(bid,"Entries").getDataRange().getValues();
  let totalEntries=0,totalRevenue=0,totalTips=0;
  eR.slice(1).forEach(r=>{
    if(String(r[2])===dt&&r[10]!==true&&r[10]!=="TRUE"){
      totalEntries++; totalRevenue+=Number(r[6])||0; totalTips+=Number(r[7])||0;
    }
  });
  return{totalEntries,totalRevenue,totalTips};
}

function getTodayAll(d){
  const bid=d.branchId; const dt=todayStr();
  const eR=branchTab(bid,"Entries").getDataRange().getValues();
  const entries=[]; const sm={};
  eR.slice(1).forEach(r=>{
    if(String(r[2])!==dt) return;
    const fl=r[10]===true||r[10]==="TRUE";
    entries.push({rowId:r[0],timestamp:r[1],staffId:r[3],staffName:r[4],service:r[5],amount:r[6],tip:r[7],paymentMethod:r[8],commissionApplies:r[9],flagged:fl});
    if(!fl){const sn=String(r[4]);if(!sm[sn])sm[sn]={name:sn,totalAmount:0,totalTip:0,entries:0,products:0};sm[sn].totalAmount+=Number(r[6])||0;sm[sn].totalTip+=Number(r[7])||0;sm[sn].entries++;}
  });
  const pR=branchTab(bid,"ProductSales").getDataRange().getValues(); const ps=[];
  pR.slice(1).forEach(r=>{
    if(String(r[2])!==dt) return;
    const fl=r[8]===true||r[8]==="TRUE";
    ps.push({rowId:r[0],timestamp:r[1],staffName:r[4],product:r[5],amount:r[6],paymentMethod:r[7],flagged:fl});
    if(!fl){const sn=String(r[4]);if(!sm[sn])sm[sn]={name:sn,totalAmount:0,totalTip:0,entries:0,products:0};sm[sn].products+=Number(r[6])||0;}
  });
  const xR=branchTab(bid,"Expenses").getDataRange().getValues(); const xs=[]; let xe=0;
  xR.slice(1).forEach(r=>{
    if(String(r[2])!==dt) return;
    const fl=r[8]===true||r[8]==="TRUE";
    xs.push({rowId:r[0],timestamp:r[1],staffName:r[4],description:r[5],amount:r[6],paymentMethod:r[7],flagged:fl});
    if(!fl) xe+=Number(r[6])||0;
  });
  return{entries,staffTotals:Object.values(sm),productSales:ps,expenses:xs,totalExp:xe};
}

function getMonthSummary(d){
  const bid=d.branchId; const tab=monthName();
  const sh=getExistingTab(bid,tab);
  if(!sh) return{summary:[],month:tab};
  const all=sh.getDataRange().getValues();
  if(all.length<3) return{summary:[],month:tab};
  const headers=all[1];
  return{summary:all.slice(2).filter(r=>r[0]).map(r=>{const o={};headers.forEach((k,i)=>{o[String(k)]=r[i];});return o;}),month:tab};
}

function deleteEntry(d){const sh=branchTab(d.branchId,"Entries"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.rowId){sh.getRange(i+1,11).setValue(true);return{};}}throw new Error("Not found");}
function deleteProduct(d){const sh=branchTab(d.branchId,"ProductSales"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.rowId){sh.getRange(i+1,9).setValue(true);return{};}}throw new Error("Not found");}
function deleteExpense(d){const sh=branchTab(d.branchId,"Expenses"),rows=sh.getDataRange().getValues();for(let i=1;i<rows.length;i++){if(rows[i][0]===d.rowId){sh.getRange(i+1,9).setValue(true);return{};}}throw new Error("Not found");}

// ── EMAIL — MULTI-RECIPIENT WITH CSV ATTACHMENTS ──────────────
// Settings tab now has: BranchID | Email1 | Email2 | Email3 | UpdatedAt
function setReportEmails(d){
  const sh=masterTab("Settings"),rows=sh.getDataRange().getValues();
  const emails=d.emails||[]; // array of up to 3 emails
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]===d.branchId){
      sh.getRange(i+1,2).setValue(emails[0]||"");
      sh.getRange(i+1,3).setValue(emails[1]||"");
      sh.getRange(i+1,4).setValue(emails[2]||"");
      sh.getRange(i+1,5).setValue(nowIST());
      return{};
    }
  }
  sh.appendRow([d.branchId,emails[0]||"",emails[1]||"",emails[2]||"",nowIST()]);
  return{};
}

function getReportEmails(d){
  const rows=masterTab("Settings").getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]===d.branchId) return{emails:[rows[i][1]||"",rows[i][2]||"",rows[i][3]||""].filter(Boolean)};
  }
  return{emails:[]};
}

// NEW: getSheetAsCSV — converts any sheet to CSV string
function getSheetAsCSV(sheetId, sheetName){
  const ss=SpreadsheetApp.openById(sheetId);
  const sh=ss.getSheetByName(sheetName); if(!sh) return "";
  const data=sh.getDataRange().getValues();
  return data.map(row=>row.map(cell=>{
    const s=String(cell).replace(/"/g,'""');
    return s.includes(",")||s.includes('"')||s.includes('\n')?`"${s}"`:s;
  }).join(",")).join("\n");
}

// Helper: get all recipient emails for a branch
function getBranchEmails(branchId){
  const rows=masterTab("Settings").getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(rows[i][0]===branchId) return [rows[i][1],rows[i][2],rows[i][3]].filter(Boolean);
  }
  return [];
}

// NEW: sendDailyReport — sends CSV-attached daily report to all recipients
function sendDailyReport(){
  masterTab("Branches").getDataRange().getValues().slice(1).forEach(row=>{
    if(row[4]===false||row[4]==="FALSE") return;
    const emails=getBranchEmails(row[0]);
    if(!emails.length) return;
    try{_sendDailyCSV(row[0],row[1],row[3],emails);}
    catch(ex){Logger.log("daily email err "+row[1]+": "+ex.message);}
  });
}

function _sendDailyCSV(bid,bname,sheetId,emails){
  const dt=todayStr();
  const csvEntries=getSheetAsCSV(sheetId,"Entries");
  const csvProducts=getSheetAsCSV(sheetId,"ProductSales");
  const csvExpenses=getSheetAsCSV(sheetId,"Expenses");
  const body=`Hello,\n\nPlease find attached the daily report for ${bname}.\nDate: ${dt}\n\nRegards,\nSalon Management System`;
  const attachments=[
    {fileName:`Daily_Entries_${dt.replace(/-/g,"")}.csv`, content:csvEntries, mimeType:"text/csv"},
    {fileName:`Daily_Products_${dt.replace(/-/g,"")}.csv`, content:csvProducts, mimeType:"text/csv"},
    {fileName:`Daily_Expenses_${dt.replace(/-/g,"")}.csv`, content:csvExpenses, mimeType:"text/csv"},
  ].filter(a=>a.content).map(a=>Utilities.newBlob(a.content,a.mimeType,a.fileName));
  emails.forEach(email=>{
    try{MailApp.sendEmail({to:email,subject:`Green Salon — ${bname} — Daily Report — ${dt}`,body,attachments});}
    catch(ex){Logger.log("send err "+email+": "+ex.message);}
  });
}

// NEW: sendMonthlyReport — sends monthly CSV on last day of month
function sendMonthlyReport(){
  const now=new Date();
  const ist=new Date(now.toLocaleString("en-US",{timeZone:"Asia/Kolkata"}));
  const lastDay=new Date(ist.getFullYear(),ist.getMonth()+1,0).getDate();
  if(ist.getDate()!==lastDay){Logger.log("Not last day — skip monthly");return;}
  masterTab("Branches").getDataRange().getValues().slice(1).forEach(row=>{
    if(row[4]===false||row[4]==="FALSE") return;
    const emails=getBranchEmails(row[0]);
    if(!emails.length) return;
    try{_sendMonthlyCSV(row[0],row[1],row[3],emails);}
    catch(ex){Logger.log("monthly email err "+row[1]+": "+ex.message);}
  });
}

function _sendMonthlyCSV(bid,bname,sheetId,emails){
  const tab=monthName();
  const csvMonthly=getSheetAsCSV(sheetId,tab);
  if(!csvMonthly){Logger.log("No monthly data for "+bname);return;}
  const ist=new Date(new Date().toLocaleString("en-US",{timeZone:"Asia/Kolkata"}));
  const ym=ist.getFullYear()+"-"+String(ist.getMonth()+1).padStart(2,"0");
  const body=`Hello,\n\nPlease find attached the monthly report for ${bname}.\nMonth: ${tab}\n\nRegards,\nSalon Management System`;
  const blob=Utilities.newBlob(csvMonthly,"text/csv",`Monthly_Report_${ym}.csv`);
  emails.forEach(email=>{
    try{MailApp.sendEmail({to:email,subject:`Green Salon — ${bname} — Monthly Report — ${tab}`,body,attachments:[blob]});}
    catch(ex){Logger.log("monthly send err "+email+": "+ex.message);}
  });
}

// NEW: sendManualReport — triggered from owner panel button
function sendManualReport(d){
  const emails=getBranchEmails(d.branchId);
  if(!emails.length) throw new Error("No email addresses set for this branch");
  const bname=branchDisplayName(d.branchId);
  const rows=masterTab("Branches").getDataRange().getValues();
  let sheetId="";
  for(let i=1;i<rows.length;i++){if(rows[i][0]===d.branchId){sheetId=rows[i][3];break;}}
  if(d.type==="monthly"){_sendMonthlyCSV(d.branchId,bname,sheetId,emails);}
  else{_sendDailyCSV(d.branchId,bname,sheetId,emails);}
  return{sent:emails.length,recipients:emails};
}

function ownerLogin(d){if(d.password!==OWNER_PASSWORD)throw new Error("Wrong password");return{ownerName:"Harsha"};}
