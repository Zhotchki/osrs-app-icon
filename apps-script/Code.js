/**
 * OSRS Herblore + Character + Quest/Diary + High Alchemy Tracker
 */

const BACKEND_SHEETS_ = ['Price_Data','Levels','Potions','Skill_Methods','Apps_Script','WikiSync_Cache','Quest_Requirements_Cache'];
const WIKISYNC_CACHE_SHEET_ = 'WikiSync_Cache';

function applyAdvancedTabVisibility_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName('Settings');
  const mode = settings ? String(settings.getRange('G19').getValue()).trim() : 'Hidden';
  if (mode === 'Shown') showAdvancedTabs();
  else hideAdvancedTabs();
}

function hideAdvancedTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  BACKEND_SHEETS_.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && !sh.isSheetHidden()) sh.hideSheet();
  });
  const settings = ss.getSheetByName('Settings');
  if (settings) settings.getRange('G19').setValue('Hidden');
}

function showAdvancedTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  BACKEND_SHEETS_.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && sh.isSheetHidden()) sh.showSheet();
  });
  const settings = ss.getSheetByName('Settings');
  if (settings) settings.getRange('G19').setValue('Shown');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('RuneScape GE')
    .addItem('Refresh all now', 'refreshAll')
    .addSeparator()
    .addItem('Enable automatic refresh', 'setupAutoRefresh')
    .addItem('Disable automatic refresh', 'disableAutoRefresh')
    .addSeparator()
    .addItem('Refresh character stats', 'refreshCharacterStats')
    .addItem('Refresh quest status', 'refreshQuestStatus')
    .addItem('Refresh quest requirements', 'refreshQuestRequirements')
    .addItem('Sync WikiSync in Browser', 'showWikiSyncBrowserSync')
        .addItem('Refresh progress tracker', 'refreshProgressTracker')
    .addItem('Refresh High Alchemy', 'refreshHighAlchemy')
    .addItem('Refresh Fastest Leveling', 'updateFastestLeveling')
    .addItem('Refresh GE prices', 'refreshGEPrices')
    .addSeparator()
    .addItem('Show Advanced Tabs', 'showAdvancedTabs')
    .addItem('Hide Advanced Tabs', 'hideAdvancedTabs')
    .addToUi();

  applyAdvancedTabVisibility_();
}



function onEdit(e) {
  if (!e || !e.range) return;
  const sh=e.range.getSheet(),a1=e.range.getA1Notation();
  if(sh.getName()==='Skill_Planner'&&a1==='B3')updateSkillPlanner();
  if(sh.getName()==='All_Skill_Optimizer'&&['B3','B4','B5','B6'].includes(a1))updateOptimizer();
  if(sh.getName()==='Settings'&&a1==='G19')applyAdvancedTabVisibility_();
}


function clearColumnsContent_(sheet, startRow, numRows, startColumn, numColumns) {
  // Google Sheets Tables can reject multi-column "column level" mutations.
  // Clearing one column at a time avoids the "selection within a single column" exception.
  for (let c = 0; c < numColumns; c++) {
    sheet.getRange(startRow, startColumn + c, numRows, 1).clearContent();
  }
}



function setColumnsNumberFormat_(sheet, startRow, numRows, startColumn, numColumns, numberFormat) {
  // Google Sheets Tables can reject formatting across multiple table columns at once.
  // Format each column independently to avoid:
  // "Please make a selection within a single column to perform column level actions."
  for (let c = 0; c < numColumns; c++) {
    sheet.getRange(startRow, startColumn + c, numRows, 1).setNumberFormat(numberFormat);
  }
}


function setColumnsValues_(sheet, startRow, startColumn, rows) {
  // Google Sheets Tables can reject multi-column writes. Write each data
  // column independently so High Alchemy remains compatible with Sheets Tables.
  if (!rows || !rows.length) return;
  const width = rows[0].length;
  for (let c = 0; c < width; c++) {
    const colValues = rows.map(r => [r[c]]);
    sheet.getRange(startRow, startColumn + c, rows.length, 1).setValues(colValues);
  }
}

function ensureOptimizerFilter_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All_Skill_Optimizer');
  if (!sh) return;
  const existing = sh.getFilter();
  if (existing) existing.remove();
  sh.getRange('A13:T263').createFilter();
}

function updateOptimizer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('All_Skill_Optimizer');
  const dataSh = ss.getSheetByName('Skill_Methods');
  if (!sh || !dataSh) return;

  const selectedSkill=String(sh.getRange('B3').getValue()).trim()||'All Skills';
  const objective=String(sh.getRange('B4').getValue()).trim()||'Fastest XP/hr';
  const showLocked=String(sh.getRange('B5').getValue()).trim()!=='No';
  const all=dataSh.getDataRange().getValues();

  const stats=ss.getSheetByName('Character_Stats').getRange(11,1,25,4).getValues();
  const levels={},xps={};
  stats.forEach(r=>{if(r[0]){levels[String(r[0])]=Number(r[2])||1;xps[String(r[0])]=Number(r[3])||0;}});

  const targetLevel=Number(sh.getRange('B6').getValue())||99;
  const levelRows=ss.getSheetByName('Levels').getRange(1,1,100,2).getValues();
  const xpByLevel={};
  levelRows.forEach(r=>{if(r[0])xpByLevel[Number(r[0])]=Number(r[1])||0;});
  const targetXp=xpByLevel[targetLevel]||0;

  const priceRows=ss.getSheetByName('Price_Data').getDataRange().getValues(),priceMap={};
  for(let i=1;i<priceRows.length;i++){
    const n=String(priceRows[i][0]||'').toLowerCase();
    if(n)priceMap[n]={buy:Number(priceRows[i][2])||0,sell:Number(priceRows[i][3])||0};
  }
  const buy=n=>n?((priceMap[String(n).toLowerCase()]||{}).buy||0):0;
  const sell=n=>n?((priceMap[String(n).toLowerCase()]||{}).sell||0):0;

  const potionRows=ss.getSheetByName('Potions').getDataRange().getValues(),potionBase={};
  for(let i=1;i<potionRows.length;i++){
    if(potionRows[i][1])potionBase[String(potionRows[i][1])]={name:potionRows[i][6]||'',qty:Number(potionRows[i][7])||0};
  }

  const oldFilter=sh.getFilter();
  if(oldFilter)oldFilter.remove();
  clearColumnsContent_(sh,14,250,1,20);

  const methods=[];
  for(let i=1;i<all.length;i++){
    const r=all[i],skill=String(r[0]||'').trim();
    if(!skill||(selectedSkill!=='All Skills'&&skill!==selectedSkill))continue;
    const req=Number(r[1])||1,method=r[2]||'',xpAction=Number(r[3])||0,actionsHr=Number(r[4])||0;
    const input1=r[5]||'',qty1=Number(r[6])||0,input2=r[7]||'',qty2=Number(r[8])||0;
    const output=r[9]||'',qtyOut=Number(r[10])||0,notes=r[11]||'',manualGp=Number(r[12])||0,source=r[13]||'';
    const curLevel=levels[skill]||1,curXp=xps[skill]||0,unlocked=req<=curLevel;
    if(!showLocked&&!unlocked)continue;

    let inputCost=buy(input1)*qty1+buy(input2)*qty2;
    if(skill==='Herblore'&&potionBase[method])inputCost+=buy(potionBase[method].name)*potionBase[method].qty;
    const outputValue=sell(output)*qtyOut,netAction=outputValue-inputCost+manualGp;
    const xpHr=xpAction*actionsHr,gpXp=xpAction?netAction/xpAction:0,gpHr=netAction*actionsHr;

    let actionsTarget='',totalInput='',totalOutput='',totalNet='',hoursTarget='';
    if(selectedSkill!=='All Skills'){
      const remain=Math.max(0,targetXp-curXp);
      actionsTarget=xpAction?Math.ceil(remain/xpAction):0;
      totalInput=inputCost*actionsTarget;totalOutput=outputValue*actionsTarget;
      totalNet=netAction*actionsTarget;hoursTarget=xpHr?remain/xpHr:'';
    }

    const status=!unlocked?'Locked':(inputCost===0&&outputValue===0&&manualGp===0)?'No direct GE value':gpHr>=0?'Profit':'Cost';
    methods.push({row:[skill,method,req,unlocked?'Yes':'No',xpAction,actionsHr,inputCost,outputValue,xpHr,gpXp,
      actionsTarget,totalInput,totalOutput,totalNet,gpHr,netAction,hoursTarget,status,notes,source],
      xpHr,gpHr,gpXp,totalInput:typeof totalInput==='number'?totalInput:Number.POSITIVE_INFINITY,
      hours:typeof hoursTarget==='number'?hoursTarget:Number.POSITIVE_INFINITY});
  }

  methods.sort((a,b)=>{
    if(objective==='Best GP/hr')return b.gpHr-a.gpHr;
    if(objective==='Lowest GP/XP')return a.gpXp-b.gpXp;
    if(objective==='Lowest Total Cost')return a.totalInput-b.totalInput;
    if(objective==='Shortest Time')return a.hours-b.hours;
    return b.xpHr-a.xpHr;
  });

  const out=methods.slice(0,250).map(x=>x.row);
  if(out.length)sh.getRange(14,1,out.length,20).setValues(out);
  else sh.getRange('A14').setValue('No methods match the current filters.');
  ensureOptimizerFilter_();
  SpreadsheetApp.flush();
}

function ensureSkillPlannerFilter_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Skill_Planner');
  if (!sh) return;
  const existing = sh.getFilter();
  if (existing) existing.remove();
  sh.getRange('A12:R100').createFilter();
}

function updateSkillPlanner() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Skill_Planner');
  const dataSh = ss.getSheetByName('Skill_Methods');
  const skill = String(sh.getRange('B3').getValue()).trim();
  if (!skill) return;

  const all = dataSh.getDataRange().getValues();
  const oldFilter = sh.getFilter();
  if (oldFilter) oldFilter.remove();

  const methods = [];
  for (let i = 1; i < all.length; i++) {
    const r = all[i];
    if (String(r[0]).trim() !== skill) continue;
    methods.push(r);
  }

  // Always show every method, including methods above the player's current level.
  methods.sort((a,b) => (Number(a[1]) || 1) - (Number(b[1]) || 1));

  clearColumnsContent_(sh, 13, 88, 1, 18);

  if (!methods.length) {
    sh.getRange('A13').setValue('No methods loaded for ' + skill);
    ensureSkillPlannerFilter_();
    return;
  }

  const rows = methods.map(r => {
    const level = Number(r[1]) || 1;
    const method = r[2] || '';
    const xp = Number(r[3]) || 0;
    const actionsHr = Number(r[4]) || 0;
    const input1 = r[5] || '';
    const qty1 = Number(r[6]) || 0;
    const input2 = r[7] || '';
    const qty2 = Number(r[8]) || 0;
    const output = r[9] || '';
    const qtyOut = Number(r[10]) || 0;
    const notes = r[11] || '';
    const manualGp = Number(r[12]) || 0;

    const detail = [];
    if (input1) detail.push(qty1 + '× ' + input1);
    if (input2) detail.push(qty2 + '× ' + input2);
    if (output) detail.push('→ ' + qtyOut + '× ' + output);

    return [
      method, level, '', xp, actionsHr, '', '', '', '', '', '', '', '', '',
      manualGp, notes, detail.join('; '), ''
    ];
  });

  sh.getRange(13,1,rows.length,18).setValues(rows);

  methods.forEach((r0, idx) => {
    const r = 13 + idx;
    const input1 = r0[5] || '';
    const qty1 = Number(r0[6]) || 0;
    const input2 = r0[7] || '';
    const qty2 = Number(r0[8]) || 0;
    const output = r0[9] || '';
    const qtyOut = Number(r0[10]) || 0;

    const esc1 = String(input1).replace(/"/g,'""');
    const esc2 = String(input2).replace(/"/g,'""');
    const escOut = String(output).replace(/"/g,'""');

    sh.getRange('C' + r).setFormula('=IF(B' + r + '<=$B$4,"Yes","No")');

    const inputParts = [];
    if (input1) inputParts.push('IFERROR(VLOOKUP("' + esc1 + '",Price_Data!A:G,3,FALSE),0)*' + qty1);
    if (input2) inputParts.push('IFERROR(VLOOKUP("' + esc2 + '",Price_Data!A:G,3,FALSE),0)*' + qty2);
    sh.getRange('F' + r).setFormula(inputParts.length ? '=' + inputParts.join('+') : '=0');

    sh.getRange('G' + r).setFormula(
      output ? '=IFERROR(VLOOKUP("' + escOut + '",Price_Data!A:G,4,FALSE),0)*' + qtyOut : '=0'
    );

    sh.getRange('H' + r).setFormula('=D' + r + '*E' + r);
    sh.getRange('I' + r).setFormula('=IF(D' + r + '=0,0,(G' + r + '-F' + r + '+O' + r + ')/D' + r + ')');
    sh.getRange('J' + r).setFormula('=IF(D' + r + '=0,0,ROUNDUP($B$8/D' + r + ',0))');
    sh.getRange('K' + r).setFormula('=F' + r + '*J' + r);
    sh.getRange('L' + r).setFormula('=G' + r + '*J' + r);
    sh.getRange('M' + r).setFormula('=(G' + r + '-F' + r + '+O' + r + ')*J' + r);
    sh.getRange('N' + r).setFormula('=(G' + r + '-F' + r + '+O' + r + ')*E' + r);
    sh.getRange('R' + r).setFormula(
      '=IF(C' + r + '="No","Locked",IF(AND(F' + r + '=0,G' + r + '=0,O' + r + '=0),"No direct GE value",IF(N' + r + '>=0,"Profit","Cost")))'
    );
  });

  sh.getRange(13,4,rows.length,12).setNumberFormat('#,##0.00');
  sh.getRange(13,10,rows.length,5).setNumberFormat('#,##0');

  // Summary recommendations intentionally use unlocked rows only.
  sh.getRange('U2').setFormula('=IFERROR(INDEX(A13:A100,MATCH(MAXIFS(H13:H100,C13:C100,"Yes"),H13:H100,0)),"")');
  sh.getRange('U3').setFormula('=IFERROR(MAXIFS(H13:H100,C13:C100,"Yes"),0)');
  sh.getRange('U4').setFormula('=IFERROR(INDEX(A13:A100,MATCH(MAXIFS(N13:N100,C13:C100,"Yes"),N13:N100,0)),"")');
  sh.getRange('U5').setFormula('=IFERROR(MAXIFS(N13:N100,C13:C100,"Yes"),0)');
  sh.getRange('U6').setFormula('=IFERROR(INDEX(A13:A100,MATCH(MINIFS(I13:I100,C13:C100,"Yes",A13:A100,"<>"),I13:I100,0)),"")');

  ensureSkillPlannerFilter_();
  SpreadsheetApp.flush();
}


function ensureHourlyTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  let found = false;

  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'scheduledAutoRefresh') {
      if (!found) found = true;
      else ScriptApp.deleteTrigger(t);
    }
  });

  if (!found) {
    ScriptApp.newTrigger('scheduledAutoRefresh')
      .timeBased()
      .everyHours(1)
      .create();
  }

  const settings = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (settings) {
    settings.getRange('G12').setValue('Enabled');
    settings.getRange('G13').setValue('Hourly');
  }
}

function updateFastestLeveling() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Fastest_Leveling');
  const dataSh = ss.getSheetByName('Skill_Methods');
  if (!sh || !dataSh) return;

  const stats = ss.getSheetByName('Character_Stats').getRange(11,1,25,4).getValues();
  const levels = {};
  stats.forEach(r => { if (r[0]) levels[String(r[0])] = Number(r[2]) || 1; });

  const priceRows = ss.getSheetByName('Price_Data').getDataRange().getValues();
  const priceMap = {};
  for (let i=1;i<priceRows.length;i++) {
    const n = String(priceRows[i][0] || '').toLowerCase();
    if (n) priceMap[n] = {buy:Number(priceRows[i][2])||0, sell:Number(priceRows[i][3])||0};
  }
  const buy = n => n ? ((priceMap[String(n).toLowerCase()] || {}).buy || 0) : 0;
  const sell = n => n ? ((priceMap[String(n).toLowerCase()] || {}).sell || 0) : 0;

  const potionRows = ss.getSheetByName('Potions').getDataRange().getValues();
  const potionBase = {};
  for (let i=1;i<potionRows.length;i++) {
    if (potionRows[i][1]) potionBase[String(potionRows[i][1])] = {
      name:potionRows[i][6] || '', qty:Number(potionRows[i][7]) || 0
    };
  }

  const methods = dataSh.getDataRange().getValues();
  const skills = [
    'Attack','Defence','Strength','Hitpoints','Ranged','Prayer','Magic','Cooking',
    'Woodcutting','Fletching','Fishing','Firemaking','Crafting','Smithing','Mining',
    'Herblore','Agility','Thieving','Slayer','Farming','Runecraft','Hunter','Construction','Sailing'
  ];

  const result = [];

  skills.forEach(skill => {
    const lvl = levels[skill] || 1;
    let best = null;
    const locked = [];

    for (let i=1;i<methods.length;i++) {
      const r = methods[i];
      if (String(r[0] || '') !== skill) continue;

      const req=Number(r[1])||1, method=r[2]||'', xpAction=Number(r[3])||0, actionsHr=Number(r[4])||0;
      const input1=r[5]||'', qty1=Number(r[6])||0, input2=r[7]||'', qty2=Number(r[8])||0;
      const output=r[9]||'', qtyOut=Number(r[10])||0, manualGp=Number(r[12])||0, source=r[13]||'';
      const xpHr=xpAction*actionsHr;

      let inputCost=buy(input1)*qty1+buy(input2)*qty2;
      if (skill==='Herblore' && potionBase[method]) inputCost += buy(potionBase[method].name)*potionBase[method].qty;
      const outputValue=sell(output)*qtyOut;
      const netAction=outputValue-inputCost+manualGp;
      const gpHr=netAction*actionsHr;
      const gpXp=xpAction ? netAction/xpAction : 0;
      const status=(inputCost===0&&outputValue===0&&manualGp===0)?'No direct GE value':gpHr>=0?'Profit':'Cost';

      const entry={method,req,xpHr,gpHr,gpXp,status,source};

      if (req <= lvl) {
        if (!best || xpHr > best.xpHr) best=entry;
      } else {
        locked.push(entry);
      }
    }

    let next = null;
    if (best) {
      locked
        .filter(x => x.xpHr > best.xpHr)
        .sort((a,b) => a.req-b.req || b.xpHr-a.xpHr);
      next = locked.filter(x => x.xpHr > best.xpHr).sort((a,b)=>a.req-b.req || b.xpHr-a.xpHr)[0] || null;
    } else {
      next = locked.sort((a,b)=>a.req-b.req || b.xpHr-a.xpHr)[0] || null;
    }

    result.push([
      skill,lvl,
      best ? best.method : 'No unlocked method',
      best ? best.req : '',
      best ? best.xpHr : 0,
      best ? best.gpHr : 0,
      best ? best.gpXp : 0,
      best ? best.status : 'Locked',
      next ? next.method : '',
      next ? next.req : '',
      next ? next.xpHr : '',
      best ? best.source : (next ? next.source : '')
    ]);
  });

  clearColumnsContent_(sh,10,31,1,12);
  sh.getRange(10,1,result.length,12).setValues(result);
  sh.getRange(10,5,result.length,3).setNumberFormat('#,##0.00');
  sh.getRange(10,11,result.length,1).setNumberFormat('#,##0');
  sh.getRange('O6').setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  SpreadsheetApp.flush();
}

function doGet(e) {
  const action = e && e.parameter ? String(e.parameter.action || '').trim() : '';

  // Public, read-only API used by the GitHub Pages PWA.
  if (action === 'mobileData') {
    return jsonOutput_(getMobileAppData());
  }

  if (action === 'optimizer') {
    const p = e.parameter || {};
    return jsonOutput_(getMobileOptimizer(
      p.skill || 'All Skills',
      p.objective || 'Fastest XP/hr',
      String(p.showLocked || '').toLowerCase() === 'true',
      Number(p.targetLevel) || 99
    ));
  }

  // Preserve the existing Apps Script mobile page for compatibility.
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('OSRS Companion')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      generatedAt: new Date().toISOString(),
      data: payload
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getMobileAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName('Settings');
  const player = settings ? String(settings.getRange('E4').getValue()).trim() : '';

  const statsRaw = ss.getSheetByName('Character_Stats').getRange(11,1,25,4).getValues();
  const stats = statsRaw.filter(r=>r[0] && r[0] !== 'Overall').map(r=>({
    skill:r[0], rank:r[1], level:r[2], xp:r[3]
  }));

  const fastestRaw = ss.getSheetByName('Fastest_Leveling').getRange(10,1,24,12).getValues();
  const fastest = fastestRaw.filter(r=>r[0]).map(r=>({
    skill:r[0], level:r[1], method:r[2], req:r[3], xpHr:r[4], gpHr:r[5], gpXp:r[6],
    status:r[7], nextMethod:r[8], nextLevel:r[9], nextXpHr:r[10], source:r[11]
  }));

  const questSheet = ss.getSheetByName('Quest_Tracker');
  const quests = questSheet.getRange(11,1,220,8).getValues().filter(r=>r[0]).map(r=>({
    name:r[0], status:r[1], readiness:r[2], missing:r[3], notes:r[4], prereqs:r[5], wiki:r[6]
  }));
  const diarySheet = ss.getSheetByName('Diary_Tracker');
  const diaries = diarySheet.getRange(10,1,48,8).getValues().filter(r=>r[0]).map(r=>({
    region:r[0], tier:r[1], status:r[2], readiness:r[3], missing:r[4], notes:r[5], wiki:r[6]
  }));

  return {
    player,
    lastRefresh: settings ? settings.getRange('G14').getDisplayValue() : '',
    refreshStatus: sanitizeMobileRefreshStatus_(settings ? settings.getRange('G15').getDisplayValue() : ''),
    stats, fastest, quests, diaries
  };
}


function sanitizeMobileRefreshStatus_(status) {
  const text = String(status || '').trim();
  // G15 can retain a failure message from a previous deployed version even
  // after that underlying code error has been fixed. Do not surface that
  // obsolete WikiSync constant error in the mobile app.
  if (/WIKISYNC_CACHE_SHEET_\s+is\s+not\s+defined/i.test(text)) {
    return 'Live data loaded';
  }
  return text;
}

function getMobileOptimizer(skill, objective, showLocked, targetLevel) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('All_Skill_Optimizer');
  sh.getRange('B3').setValue(skill || 'All Skills');
  sh.getRange('B4').setValue(objective || 'Fastest XP/hr');
  sh.getRange('B5').setValue(showLocked ? 'Yes' : 'No');
  sh.getRange('B6').setValue(Number(targetLevel) || 99);
  updateOptimizer();

  const rows=sh.getRange(14,1,100,20).getValues().filter(r=>r[0]);
  return rows.map(r=>({
    skill:r[0],method:r[1],req:r[2],unlocked:r[3],xpAction:r[4],actionsHr:r[5],
    inputCost:r[6],outputValue:r[7],xpHr:r[8],gpXp:r[9],actionsTarget:r[10],
    totalInput:r[11],totalOutput:r[12],netGp:r[13],gpHr:r[14],netAction:r[15],
    hours:r[16],status:r[17],notes:r[18],source:r[19]
  }));
}

function refreshMobileApp() {
  refreshAll();
  return getMobileAppData();
}

function setupAutoRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureHourlyTrigger_();
  scheduledAutoRefresh();
  ss.toast('Automatic hourly refresh enabled.', 'OSRS Auto Refresh', 6);
}

function disableAutoRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName('Settings');

  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'scheduledAutoRefresh') {
      ScriptApp.deleteTrigger(t);
    }
  });

  settings.getRange('G12').setValue('Disabled');
  settings.getRange('G13').setValue('');
  ss.toast('Automatic refresh disabled.', 'OSRS Auto Refresh', 5);
}

function scheduledAutoRefresh() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName('Settings');
  const lock = LockService.getScriptLock();

  // Prevent overlapping trigger runs.
  if (!lock.tryLock(10000)) return;

  try {
    const errors = [];

    try { refreshCharacterStats(); }
    catch (e) { errors.push('Character: ' + e.message); }

    try { refreshProgressTracker(); }
    catch (e) { errors.push('Progress: ' + e.message); }

    try { refreshGEPrices(); }
    catch (e) { errors.push('GE: ' + e.message); }

    try { refreshHighAlchemy(); }
    catch (e) { errors.push('High Alch: ' + e.message); }

    try { updateSkillPlanner(); }
    catch (e) { errors.push('Skill Planner: ' + e.message); }
    try { updateOptimizer(); }
    catch (e) { errors.push('Optimizer: ' + e.message); }
    try { updateFastestLeveling(); }
    catch (e) { errors.push('Fastest Leveling: ' + e.message); }

    settings.getRange('G14').setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    settings.getRange('G15').setValue(
      errors.length ? 'Completed with errors: ' + errors.join(' | ') : 'Success'
    );
  } finally {
    lock.releaseLock();
  }
}

function refreshAll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName('Settings');
  const errors = [];

  // Running Refresh All manually also installs/repairs the hourly trigger.
  try { ensureHourlyTrigger_(); }
  catch (e) { errors.push('Auto Refresh: ' + e.message); }

  try { refreshCharacterStats(); }
  catch (e) { errors.push('Character Stats: ' + e.message); }

  try { refreshProgressTracker(); }
  catch (e) { errors.push('Quest/Diary: ' + e.message); }

  try { refreshGEPrices(); }
  catch (e) { errors.push('GE Prices: ' + e.message); }

  try { refreshHighAlchemy(); }
  catch (e) { errors.push('High Alchemy: ' + e.message); }

  try { updateSkillPlanner(); }
  catch (e) { errors.push('Skill Planner: ' + e.message); }

  try { updateOptimizer(); }
  catch (e) { errors.push('Optimizer: ' + e.message); }

  try { updateFastestLeveling(); }
  catch (e) { errors.push('Fastest Leveling: ' + e.message); }

  settings.getRange('G14').setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  settings.getRange('G15').setValue(errors.length ? 'Refresh issues: ' + errors.join(' | ') : 'Success');

  SpreadsheetApp.flush();

  if (errors.length) {
    ss.toast('Refresh completed with ' + errors.length + ' issue(s). Check Settings → Last result.', 'OSRS Refresh', 8);
  } else {
    ss.toast('All RuneScape data refreshed successfully.', 'OSRS Refresh', 5);
  }
}

function refreshCharacterStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName('Settings');
  const statSheet = ss.getSheetByName('Character_Stats');
  const player = String(settings.getRange('E4').getValue()).trim();
  if (!player) throw new Error('Enter an OSRS character name in Settings!E4.');

  const url = 'https://secure.runescape.com/m=hiscore_oldschool/index_lite.ws?player=' + encodeURIComponent(player);
  const resp = UrlFetchApp.fetch(url, {
    muteHttpExceptions:true,
    headers:{'User-Agent':'OSRS-Training-GoogleSheet/10.14.9 - personal Google Sheets tool'}
  });
  if (resp.getResponseCode() !== 200) {
    const msg = resp.getResponseCode() === 404 ? 'Character not found / not ranked' : 'Hiscores request failed (' + resp.getResponseCode() + ')';
    settings.getRange('E8').setValue(msg);
    throw new Error(msg);
  }

  const skillNames = [
    'Overall','Attack','Defence','Strength','Hitpoints','Ranged','Prayer','Magic',
    'Cooking','Woodcutting','Fletching','Fishing','Firemaking','Crafting','Smithing',
    'Mining','Herblore','Agility','Thieving','Slayer','Farming','Runecraft','Hunter',
    'Construction','Sailing'
  ];
  const lines = resp.getContentText().trim().split(/\r?\n/);
  const rows = [];
  let herbLevel = '', herbXp = '';
  for (let i = 0; i < skillNames.length; i++) {
    const p = String(lines[i] || '-1,1,0').split(',');
    const row = [skillNames[i], Number(p[0]), Number(p[1]), Number(p[2])];
    rows.push(row);
    if (skillNames[i] === 'Herblore') { herbLevel = row[2]; herbXp = row[3]; }
  }
  statSheet.getRange(11,1,25,4).clearContent();
  statSheet.getRange(11,1,skillNames.length,4).setValues(rows);
  settings.getRange('E6').setValue(herbLevel);
  settings.getRange('E7').setValue(herbXp);
  settings.getRange('E8').setValue('Loaded: ' + player);
  settings.getRange('E9').setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}


function ensureHighAlchemyLayout_(sh) {
  const headers = [
    'Item','Members?','GE Guide Price','Live Instant Buy','Cost Used','Cost Source',
    '5m Avg Buy','5m Buy Volume','1h Avg Buy','1h Buy Volume',
    'Latest Buy','Latest Buy Age (min)','High Alch Value','Nature Rune',
    'Net Profit / Cast','GP / XP','GE Buy Limit (4h)','Max Qty / 4h',
    'Cash Needed for Limit','Profit per GE Limit','Magic XP / Cast','Casts to Target',
    'Qty Useful to Target','Projected Profit to Target','XP / Hour','Hours to Target','Status'
  ];

  setColumnsValues_(sh, 12, 1, [headers]);

  sh.getRange('A9').setValue('Nature rune cost used');
  sh.getRange('C9').setValue('GE guide');
  sh.getRange('E9').setValue('Live instant buy');
  sh.getRange('G9').setValue('Cost source');
  sh.getRange('I9').setValue('Latest age (min)');
  sh.getRange('K9').setValue('Hide >30m trades?');

  // Clickable checkbox toggle in Google Sheets.
  const ageToggle = sh.getRange('L9');
  const priorToggle = ageToggle.getValue() === true;
  ageToggle.setDataValidation(
    SpreadsheetApp.newDataValidation().requireCheckbox().build()
  );
  ageToggle.setValue(priorToggle);

  sh.getRange('A10').setValue(
    'v10.14.9 separates two different GE prices: GE Guide Price = the in-game/Jagex guide price; ' +
    'Live Instant Buy = recent real trades. High Alch profit uses Cost Used, which prefers a liquid 5m average, ' +
    'then 1h average, then latest instant-buy. This avoids calling a live-trade estimate the GE guide price. Use the checkbox in L9 to hide items whose latest instant-buy trade is older than 30 minutes.'
  );
  sh.getRange('A11').setValue(
    'Guide price source: https://api.weirdgloop.org/exchange/history/osrs/latest | ' +
    'Live trade sources: https://prices.runescape.wiki/api/v1/osrs/5m | ' +
    'https://prices.runescape.wiki/api/v1/osrs/1h | https://prices.runescape.wiki/api/v1/osrs/latest'
  );
}

function fetchGuidePricesById_(ids) {
  // Weird Gloop's latest endpoint returns the official/Jagex GE daily price.
  // It supports up to 100 item IDs per request, so batch efficiently.
  const out = {};
  const unique = [...new Set(ids.map(String).filter(Boolean))];
  const batches = [];
  for (let i = 0; i < unique.length; i += 100) {
    batches.push(unique.slice(i, i + 100));
  }
  if (!batches.length) return out;

  const requests = batches.map(batch => ({
    url:'https://api.weirdgloop.org/exchange/history/osrs/latest?id=' +
        encodeURIComponent(batch.join('|')),
    muteHttpExceptions:true,
    headers:{'User-Agent':'OSRS-Training-GoogleSheet/10.14.9 - personal Google Sheets tool'}
  }));

  const responses = UrlFetchApp.fetchAll(requests);
  responses.forEach(resp => {
    if (resp.getResponseCode() !== 200) return;
    try {
      const json = JSON.parse(resp.getContentText());
      Object.keys(json || {}).forEach(id => {
        const entry = json[id];
        if (entry && Number(entry.price) > 0) {
          out[String(id)] = Number(entry.price);
        }
      });
    } catch (e) {}
  });
  return out;
}

function refreshHighAlchemy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('High_Alchemy');
  const settings = ss.getSheetByName('Settings');
  if (!sh) throw new Error('High_Alchemy sheet not found.');
  if (!settings) throw new Error('Settings sheet not found.');

  ensureHighAlchemyLayout_(sh);

  const base = 'https://prices.runescape.wiki/api/v1/osrs';
  const options = {
    muteHttpExceptions:true,
    headers:{'User-Agent':'OSRS-Training-GoogleSheet/10.14.9 - personal Google Sheets tool'}
  };

  const responses = UrlFetchApp.fetchAll([
    {url:base + '/mapping', ...options},
    {url:base + '/latest', ...options},
    {url:base + '/5m', ...options},
    {url:base + '/1h', ...options}
  ]);

  if (responses.some(r => r.getResponseCode() !== 200)) {
    throw new Error(
      'High Alchemy live-price request failed. HTTP codes: ' +
      responses.map(r => r.getResponseCode()).join(', ')
    );
  }

  const mapping = JSON.parse(responses[0].getContentText());
  const latest = JSON.parse(responses[1].getContentText()).data || {};
  const fiveMin = (JSON.parse(responses[2].getContentText()).data || {});
  const oneHour = (JSON.parse(responses[3].getContentText()).data || {});
  const nowSec = Math.floor(Date.now() / 1000);

  // Fetch official GE guide prices only for alchable, tradeable candidates.
  const candidateIds = mapping
    .filter(item => Number(item.highalch) > 0 && Number(item.limit) > 0)
    .map(item => item.id);
  const guidePrices = fetchGuidePricesById_(candidateIds);

  function marketPrice_(id) {
    const key = String(id);
    const p5 = fiveMin[key] || {};
    const p1 = oneHour[key] || {};
    const pl = latest[key] || {};

    const fivePrice = Number(p5.avgHighPrice) || 0;
    const fiveVol = Number(p5.highPriceVolume) || 0;
    const hourPrice = Number(p1.avgHighPrice) || 0;
    const hourVol = Number(p1.highPriceVolume) || 0;
    const latestPrice = Number(pl.high) || 0;
    const latestTime = Number(pl.highTime) || 0;
    const ageMinutes = latestTime > 0 ? Math.max(0, (nowSec - latestTime) / 60) : '';

    let used = 0;
    let source = '';
    if (fivePrice > 0 && fiveVol >= 5) {
      used = fivePrice;
      source = '5m avg';
    } else if (hourPrice > 0 && hourVol >= 1) {
      used = hourPrice;
      source = '1h avg';
    } else if (latestPrice > 0) {
      used = latestPrice;
      source = 'Latest';
    }

    return {used,source,fivePrice,fiveVol,hourPrice,hourVol,latestPrice,ageMinutes};
  }

  const nature = mapping.find(x => String(x.name).toLowerCase() === 'nature rune');
  if (!nature) throw new Error('Could not find Nature rune in price mapping.');
  const natureMarket = marketPrice_(nature.id);
  const natureGuide = Number(guidePrices[String(nature.id)]) || 0;
  if (!natureMarket.used) throw new Error('Could not find a usable live Nature rune price.');
  const naturePrice = natureMarket.used;
  settings.getRange('E17').setValue(naturePrice);

  const magicLevel = Number(settings.getRange('E12').getValue()) || 1;
  const minProfit = Number(settings.getRange('E16').getValue()) || 0;
  const castsPerHour = Number(settings.getRange('E15').getValue()) || 1200;
  const castsToTarget = Number(sh.getRange('B8').getValue()) || 0;
  const hideStaleTrades = sh.getRange('L9').getValue() === true;
  const maxTradeAgeMinutes = 30;

  const rows = [];
  mapping.forEach(item => {
    const highAlch = Number(item.highalch);
    const limit = Number(item.limit) || 0;
    if (!Number.isFinite(highAlch) || highAlch <= 0 || limit <= 0) return;

    const market = marketPrice_(item.id);
    const guide = Number(guidePrices[String(item.id)]) || 0;
    const buy = market.used;
    if (buy <= 0) return;

    // Optional freshness filter. Trade age is based on the timestamp of the
    // latest instant-buy (high) trade from the OSRS Wiki real-time feed.
    if (
      hideStaleTrades &&
      market.ageMinutes !== '' &&
      Number(market.ageMinutes) > maxTradeAgeMinutes
    ) return;

    const profit = highAlch - buy - naturePrice;
    if (profit < minProfit) return;

    const gpPerXp = profit / 65;
    const qtyUseful = castsToTarget > 0 ? Math.min(limit, castsToTarget) : 0;
    const cashNeeded = buy * limit;
    const profitLimit = profit * limit;
    const projectedProfit = profit * qtyUseful;
    const xpHour = 65 * castsPerHour;
    const hours = castsPerHour > 0 ? castsToTarget / castsPerHour : 0;
    const status = magicLevel < 55 ? 'Magic < 55' : (profit >= 0 ? 'Profitable' : 'Loss');

    rows.push([
      item.name, item.members ? 'Yes' : 'No',
      guide, market.latestPrice, buy, market.source,
      market.fivePrice, market.fiveVol, market.hourPrice, market.hourVol,
      market.latestPrice, market.ageMinutes,
      highAlch, naturePrice, profit, gpPerXp,
      limit, limit, cashNeeded, profitLimit,
      65, castsToTarget, qtyUseful, projectedProfit,
      xpHour, hours, status
    ]);
  });

  rows.sort((a,b) => b[14] - a[14]);

  clearColumnsContent_(sh, 13, 2000, 1, 27);
  if (rows.length) {
    const maxRows = Math.min(rows.length, 2000);
    const outputRows = rows.slice(0,maxRows);
    setColumnsValues_(sh, 13, 1, outputRows);

    // Numeric formatting, table-safe one column at a time.
    [3,4,5,7,8,9,10,11,13,14,15,17,18,19,20,21,22,23,24,25].forEach(c =>
      setColumnsNumberFormat_(sh, 13, maxRows, c, 1, '#,##0')
    );
    setColumnsNumberFormat_(sh, 13, maxRows, 12, 1, '0.0');
    setColumnsNumberFormat_(sh, 13, maxRows, 16, 1, '#,##0.00');
    setColumnsNumberFormat_(sh, 13, maxRows, 26, 1, '0.00');
  }

  sh.getRange('B9').setValue(naturePrice);
  sh.getRange('D9').setValue(natureGuide);
  sh.getRange('F9').setValue(natureMarket.latestPrice);
  sh.getRange('H9').setValue(natureMarket.source);
  sh.getRange('J9').setValue(natureMarket.ageMinutes);

  SpreadsheetApp.flush();
  ss.toast(
    'High Alchemy refreshed' +
      (hideStaleTrades ? ' — trades older than 30 minutes hidden.' : '.') ,
    'High Alchemy',
    6
  );
}

function getWikiSyncPlayerName() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  return sh ? String(sh.getRange('E4').getValue()).trim() : '';
}


function ensureWikiSyncCacheSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('WikiSync_Cache');

  if (!sh) {
    sh = ss.insertSheet('WikiSync_Cache');
  }

  // Keep a simple two-column cache layout.
  if (sh.getMaxColumns() < 2) {
    sh.insertColumnsAfter(sh.getMaxColumns(), 2 - sh.getMaxColumns());
  }

  if (!sh.getRange('A1').getValue()) {
    sh.getRange('A1:B1').setValues([['WikiSync Cache','Value']]);
  }

  try { sh.hideSheet(); } catch(e) {}
  return sh;
}

function saveWikiSyncBrowserPayload(player, profile, rawPayload, questRequirements, browserDiaryData, browserDiaryPath) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const settings=ss.getSheetByName('Settings');
  const normalized=unwrapWikiSyncPayload_(rawPayload);
  const quests=normalized.quests || findKeyRecursive_(normalized,'quests');

  // Prefer the browser's explicit extraction because it sees the untouched
  // WikiSync response. Fall back to searching both raw and normalized payloads.
  const rawDiaryInfo=extractWikiSyncDiaryData_(rawPayload);
  const normalizedDiaryInfo=extractWikiSyncDiaryData_(normalized);
  const diaries=(browserDiaryData && typeof browserDiaryData === 'object')
    ? browserDiaryData
    : (rawDiaryInfo.data || normalizedDiaryInfo.data || null);
  const diaryPath=String(browserDiaryPath || rawDiaryInfo.path || normalizedDiaryInfo.path || '');

  const cacheSheet=ensureWikiSyncCacheSheet_();
  const json=JSON.stringify({
    player:String(player||''),
    profile:String(profile||'STANDARD'),
    timestamp:Date.now(),
    data:normalized,
    quests:quests || {},
    achievement_diaries:diaries || null,
    achievement_diaries_source:diaryPath || ''
  });

  clearColumnsContent_(cacheSheet,1,Math.max(cacheSheet.getLastRow(),2),1,2);
  cacheSheet.getRange('A1:B1').setValues([['WikiSync Cache','Value']]);
  const chunks=[];
  for(let i=0;i<json.length;i+=40000) chunks.push([json.slice(i,i+40000)]);
  if(chunks.length) cacheSheet.getRange(2,1,chunks.length,1).setValues(chunks);
  cacheSheet.getRange('B1').setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  try { cacheSheet.hideSheet(); } catch(e) {}

  saveQuestRequirementsFromBridge_(questRequirements || {});

  try {
    settings.getRange('J12').setValue('Connected');
    settings.getRange('J13').setValue(profile);
    settings.getRange('J14').setValue('OSRS Wiki browser bridge');
    settings.getRange('J15').setValue(quests ? Object.keys(quests).length + ' quests' : 'Missing');
    settings.getRange('J16').setValue(diaries ? 'Present' : 'Missing');
    settings.getRange('J17').setValue(diaries
      ? ('Diary source: ' + (diaryPath || 'detected') + ' | keys: ' + Object.keys(diaries).slice(0,12).join(', '))
      : ('No diary section found. Raw keys: ' + Object.keys(rawPayload || {}).slice(0,20).join(', ')));
    settings.getRange('J18').setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    settings.getRange('J19').setValue('');
    settings.getRange('J20').setValue('Cached');
  } catch(e) {}

  refreshQuestsExact_({
    profile:String(profile||'STANDARD'),
    timestamp:Date.now(),
    data:normalized,
    quests:quests || {},
    achievement_diaries:diaries || null,
    achievement_diaries_source:diaryPath || ''
  });

  try { refreshDiaries_({
    profile:String(profile||'STANDARD'),
    timestamp:Date.now(),
    data:normalized,
    quests:quests || {},
    achievement_diaries:diaries || null,
    achievement_diaries_source:diaryPath || ''
  }); } catch(e) {}

  return {
    profile:String(profile||'STANDARD'),
    questCount:quests ? Object.keys(quests).length : 0,
    diaries:!!diaries,
    diaryCount:diaries && typeof diaries === 'object' ? Object.keys(diaries).length : 0,
    diarySource:diaryPath || '',
    requirementCount:questRequirements ? Object.keys(questRequirements).length : 0
  };
}

function loadCachedWikiSync_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cache = ss.getSheetByName(WIKISYNC_CACHE_SHEET_);
  if (!cache) return null;

  const player = String(cache.getRange('B2').getValue()).trim();
  const profile = String(cache.getRange('B3').getValue()).trim();
  const timestamp = cache.getRange('B4').getValue();
  const chunkCount = Number(cache.getRange('B8').getValue()) || 0;
  if (!player || !profile || !chunkCount) return null;

  const expected = String(ss.getSheetByName('Settings').getRange('E4').getValue()).trim();
  if (expected && player.toLowerCase() !== expected.toLowerCase()) return null;

  const vals = cache.getRange(11,2,chunkCount,1).getValues();
  const text = vals.map(r=>String(r[0]||'')).join('');
  if (!text) return null;

  try {
    const data = JSON.parse(text);
    const quests = data.quests || findKeyRecursive_(data,'quests');
    const diaryInfo=extractWikiSyncDiaryData_(data);
    const diaries=data.achievement_diaries || data.achievementDiaries || diaryInfo.data || null;
    return {
      player,profile,timestamp,data,quests,diaries,source:'cache',
      diarySource:data.achievement_diaries_source || diaryInfo.path || ''
    };
  } catch (e) {
    wikiSyncDiagnostics_({
      connection:'Cache error',
      profile:profile,
      source:'Cached browser data',
      quests:'Unknown',
      diaries:'Unknown',
      keys:'',
      cached:'Corrupt',
      error:e.message
    });
    return null;
  }
}

function showWikiSyncBrowserSync() {
  const html = HtmlService.createHtmlOutputFromFile('WikiSync')
    .setWidth(460)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, 'Sync WikiSync in Browser');
}


function refreshQuestRequirements() {
  SpreadsheetApp.getUi().alert(
    'Quest requirements now come from the OSRS Wiki Questreq engine through the browser bridge.\n\n' +
    'Use RuneScape GE → Sync WikiSync in Browser. That single sync now refreshes both WikiSync status and quest requirements.'
  );
}

function refreshQuestStatus() {
  const sync = loadCachedWikiSync_();
  if (!sync) {
    throw new Error('No cached browser WikiSync data. Use RuneScape GE → Sync WikiSync in Browser first.');
  }
  refreshQuestsExact_(sync);
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Quest states refreshed from cached browser WikiSync (' + sync.profile + ')',
    'OSRS WikiSync',
    5
  );
}

function refreshProgressTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName('Settings');
  const player = String(settings.getRange('E4').getValue()).trim();
  if (!player) throw new Error('Enter an OSRS character name in Settings!E4.');

  const levels = getSkillLevels_();
  const sync = loadCachedWikiSync_();

  if (sync) {
    refreshQuestsExact_(sync);
    refreshDiaries_(player, levels, sync.data, 'Cached browser WikiSync: ' + sync.profile);
    wikiSyncDiagnostics_({
      connection:'Cached',
      profile:sync.profile,
      source:'Cached browser data',
      quests:sync.quests && typeof sync.quests === 'object' ? 'Present (' + Object.keys(sync.quests).length + ')' : 'Not present',
      diaries:sync.diaries && typeof sync.diaries === 'object' ? 'Present' : 'Not present',
      keys:Object.keys(sync.data || {}).slice(0,20).join(', '),
      cached:'Yes',
      error:''
    });
  } else {
    const qsh = ss.getSheetByName('Quest_Tracker');
    qsh.getRange('B4').setValue('Browser sync required');
    qsh.getRange('B5:B6').clearContent();
    clearColumnsContent_(qsh, 11, 220, 1, 8);
    qsh.getRange('A11:H11').setValues([[
      'WikiSync browser sync required','Unknown','Unknown','','',
      '', 'https://oldschool.runescape.wiki/w/RuneScape:WikiSync', new Date()
    ]]);
    refreshDiaries_(player, levels, null, 'Run Wiki bridge WikiSync first');
    wikiSyncDiagnostics_({
      connection:'Browser sync required',
      profile:'',
      source:'Wiki bridge',
      quests:'No cache',
      diaries:'No cache',
      keys:'',
      cached:'No',
      error:'Use RuneScape GE → Sync WikiSync in Browser or open the mobile app.'
    });
  }

  SpreadsheetApp.flush();
  ss.toast(
    sync ? 'Quest and diary tracker refreshed from cached WikiSync' : 'WikiSync browser sync required',
    'OSRS Progress',
    5
  );
}

function wikiSyncDiagnostics_(values) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  if (!sh) return;
  sh.getRange('J12').setValue(values.connection || '');
  sh.getRange('J13').setValue(values.profile || '');
  sh.getRange('J14').setValue(values.source || 'Wiki bridge');
  sh.getRange('J15').setValue(values.quests || '');
  sh.getRange('J16').setValue(values.diaries || '');
  sh.getRange('J17').setValue(values.keys || '');
  sh.getRange('J18').setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.getRange('J19').setValue(values.error || '');
  sh.getRange('J20').setValue(values.cached || '');
}

function unwrapWikiSyncPayload_(raw) {
  if (!raw || typeof raw !== 'object') return raw;
  // Some clients/proxies expose the response under a data property.
  if (raw.data && typeof raw.data === 'object') {
    const d = raw.data;
    if (d.quests || d.achievement_diaries || d.achievementDiaries || d.levels ||
        d.combat_achievements || d.collection_log || d.music_tracks) {
      return d;
    }
  }
  return raw;
}

function fetchWikiSyncProfile_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = ss.getSheetByName('Settings');
  const player = String(settings.getRange('E4').getValue()).trim();
  if (!player) {
    wikiSyncDiagnostics_({connection:'Failed',error:'Character name is blank in Settings!E4'});
    throw new Error('Enter an OSRS character name in Settings!E4.');
  }

  const profiles = ['STANDARD','IRONMAN','HARDCORE_IRONMAN','ULTIMATE_IRONMAN'];
  const options = {
    muteHttpExceptions:true,
    followRedirects:true,
    headers:{
      'User-Agent':'OSRS-Training-GoogleSheet/10.12 (personal WikiSync client)',
      'Accept':'application/json',
      'Cache-Control':'no-cache'
    }
  };

  const errors = [];

  for (let i=0;i<profiles.length;i++) {
    const profile = profiles[i];
    const url = 'https://sync.runescape.wiki/runelite/player/' +
      encodeURIComponent(player) + '/' + profile + '?t=' + Date.now();

    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      const text = resp.getContentText() || '';

      if (code !== 200) {
        errors.push(profile + ': HTTP ' + code);
        continue;
      }

      let raw;
      try {
        raw = JSON.parse(text);
      } catch (parseErr) {
        errors.push(profile + ': invalid JSON');
        continue;
      }

      const data = unwrapWikiSyncPayload_(raw);
      if (!data || typeof data !== 'object' || Array.isArray(data) || !Object.keys(data).length) {
        errors.push(profile + ': empty payload');
        continue;
      }

      const questData = data.quests || findKeyRecursive_(data, 'quests');
      const diaryData = data.achievement_diaries || data.achievementDiaries ||
        findKeyRecursive_(data, 'achievement_diaries') || findKeyRecursive_(data, 'achievementDiaries');

      const keys = Object.keys(data).slice(0,20).join(', ');
      wikiSyncDiagnostics_({
        connection:'Connected',
        profile:profile,
        http:'HTTP 200',
        quests:questData && typeof questData === 'object' ? 'Present' : 'Not present',
        diaries:diaryData && typeof diaryData === 'object' ? 'Present' : 'Not present',
        keys:keys,
        error:''
      });

      return {
        player:player,
        profile:profile,
        timestamp:data.timestamp || raw.timestamp || '',
        data:data,
        quests:questData,
        diaries:diaryData
      };
    } catch (e) {
      errors.push(profile + ': ' + e.message);
    }
  }

  const msg = errors.join(' | ');
  wikiSyncDiagnostics_({
    connection:'Failed',
    profile:'',
    http:'No valid HTTP 200 payload',
    quests:'Unknown',
    diaries:'Unknown',
    keys:'',
    error:msg
  });

  throw new Error(
    'No usable WikiSync profile found for "' + player + '". ' +
    'Enable WikiSync in RuneLite, log into that character, then run Test WikiSync Connection. ' + msg
  );
}

function testWikiSyncConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    const sync = fetchWikiSyncProfile_();
    const q = sync.quests && typeof sync.quests === 'object' ? Object.keys(sync.quests).length : 0;
    const d = sync.diaries && typeof sync.diaries === 'object' ? Object.keys(sync.diaries).length : 0;
    ss.toast(
      'Connected: ' + sync.profile + ' | quests: ' + q + ' | diary keys: ' + d,
      'OSRS WikiSync',
      8
    );
  } catch (e) {
    ss.toast('WikiSync test failed. See Settings → WikiSync Diagnostics.', 'OSRS WikiSync', 8);
    throw e;
  }
}


function extractQuestRequirementsBlock_(wt) {
  const text = String(wt || '');
  if (!text) return '';

  const sec = extractSection_(text, 'Requirements');
  if (sec && sec.trim().length > 20) return sec.slice(0, 16000);

  const m = /\|\s*requirements\s*=\s*/i.exec(text);
  if (!m) return text.slice(0,10000);

  let i = m.index + m[0].length;
  let braceDepth = 0;
  let linkDepth = 0;
  let out = '';

  while (i < text.length) {
    const two = text.slice(i,i+2);

    if (two === '{{') { braceDepth++; out += two; i += 2; continue; }
    if (two === '}}') { if (braceDepth > 0) braceDepth--; out += two; i += 2; continue; }
    if (two === '[[') { linkDepth++; out += two; i += 2; continue; }
    if (two === ']]') { if (linkDepth > 0) linkDepth--; out += two; i += 2; continue; }

    if (braceDepth === 0 && linkDepth === 0 && text[i] === '\n') {
      const rest = text.slice(i);
      if (/^\n\s*\|\s*[A-Za-z_][A-Za-z0-9 _-]*\s*=/.test(rest)) break;
      if (/^\n\s*\}\}/.test(rest)) break;
    }

    out += text[i];
    i++;
  }

  out = out.trim();
  return out ? out.slice(0,16000) : text.slice(0,10000);
}








function saveQuestRequirementsFromBridge_(reqMap) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName('Quest_Requirements_Cache');
  if(!sh) sh=ss.insertSheet('Quest_Requirements_Cache');

  clearColumnsContent_(sh,1,Math.max(sh.getLastRow(),2),1,7);
  sh.getRange('A1:G1').setValues([[
    'Quest','Skill Requirements','Quest Prereqs','Source','Fetched At','Parse Status','Raw Skill Text'
  ]]);

  const rows=[];
  Object.keys(reqMap || {}).sort((a,b)=>a.localeCompare(b)).forEach(q=>{
    const r=reqMap[q] || {};
    const skills=r.skills || {};
    const skillText=Object.keys(skills).sort().map(k=>k+' '+skills[k]).join(', ');
    const prereqs=Array.isArray(r.prereqs) ? r.prereqs.join(', ') : String(r.prereqs||'');
    rows.push([
      q,
      skillText,
      prereqs,
      r.source || 'https://oldschool.runescape.wiki/w/Template:Questreq',
      r.fetchedAt ? new Date(r.fetchedAt) : new Date(),
      'Loaded via Questreq template',
      r.skillText || ''
    ]);
  });

  if(rows.length) sh.getRange(2,1,rows.length,7).setValues(rows);
  sh.getRange('E2:E').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  try { sh.hideSheet(); } catch(e) {}
}

function questRequirementCacheMap_() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName('Quest_Requirements_Cache');
  if(!sh) return {sheet:null,map:{}};
  const last=sh.getLastRow();
  const map={};
  if(last < 2) return {sheet:sh,map:map};

  sh.getRange(2,1,last-1,7).getValues().forEach((r,idx)=>{
    if(!r[0]) return;
    const skills={};
    String(r[1]||'').split(/\s*,\s*/).forEach(x=>{
      const m=/^(.+?)\s+(\d{1,3})$/.exec(x.trim());
      if(m) skills[m[1]]=Number(m[2]);
    });
    map[String(r[0]).toLowerCase()]={
      row:idx+2,
      quest:String(r[0]),
      skills:skills,
      prereqs:String(r[2]||''),
      source:String(r[3]||''),
      fetched:r[4],
      status:String(r[5]||''),
      rawSkillText:String(r[6]||'')
    };
  });
  return {sheet:sh,map:map};
}

function missingSkillsFromMap_(skills, levels) {
  const misses=[];
  Object.keys(skills||{}).sort().forEach(skill=>{
    const have=Number(levels[String(skill).toLowerCase()])||0;
    const need=Number(skills[skill])||0;
    if(need>have) misses.push(skill+' '+have+'/'+need);
  });
  return misses.join(', ');
}

function refreshQuestsExact_(sync) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sh=ss.getSheetByName('Quest_Tracker');
  const quests=sync.quests || (sync.data ? (sync.data.quests || findKeyRecursive_(sync.data,'quests')) : null);
  const levels=getSkillLevels_();
  const reqCache=questRequirementCacheMap_().map;

  clearColumnsContent_(sh,11,220,1,8);
  sh.getRange('B4').setValue('Connected');
  sh.getRange('B5').setValue(sync.profile || '');

  let ts=new Date();
  if(sync.timestamp){
    const raw=Number(sync.timestamp);
    if(!isNaN(raw)&&raw>0) ts=new Date(raw<1000000000000?raw*1000:raw);
  }
  sh.getRange('B6').setValue(ts).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  if(!quests || typeof quests!=='object' || Array.isArray(quests)){
    sh.getRange('A11:H11').setValues([[
      'Quest section not present in WikiSync payload','Unknown','','',
      'WikiSync connected, but no quest map was returned.','',
      'https://oldschool.runescape.wiki/w/RuneScape:WikiSync',new Date()
    ]]);
    return;
  }

  const names=Object.keys(quests).sort((a,b)=>a.localeCompare(b));
  const rows=names.slice(0,220).map(name=>{
    const rawValue=quests[name];
    const raw=Number(rawValue && typeof rawValue==='object' && rawValue.state!==undefined ? rawValue.state : rawValue);
    const status=raw===2?'Complete':raw===1?'In Progress':raw===0?'Incomplete':normalizeQuestState_(rawValue);

    const cached=reqCache[String(name).toLowerCase()];
    const skills=cached ? cached.skills : {};
    const prereqs=cached ? cached.prereqs : '';
    const missing=missingSkillsFromMap_(skills,levels);
    const skillSummary=Object.keys(skills||{}).sort().map(k=>k+' '+skills[k]).join(', ');

    let readiness='Unknown';
    if(status==='Complete') readiness='Complete';
    else if(cached) readiness=missing?'Missing':'Ready';

    let notes='';
    if(cached){
      notes=skillSummary ? ('Skill requirements: '+skillSummary) : 'No skill level requirements';
    } else {
      notes='Run Sync WikiSync in Browser to load Questreq data';
    }

    const wiki='https://oldschool.runescape.wiki/w/'+encodeURIComponent(name.replace(/ /g,'_'));
    return [name,status,readiness,missing,notes,prereqs,wiki,new Date()];
  });

  if(rows.length){
    sh.getRange(11,1,rows.length,8).setValues(rows);
    sh.getRange(11,8,rows.length,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
}

function getSkillLevels_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Character_Stats');
  const vals = sh.getRange(11,1,25,4).getValues();
  const out = {};
  vals.forEach(r => { if (r[0]) out[String(r[0]).toLowerCase()] = Number(r[2]) || 0; });
  return out;
}

function findKeyRecursive_(obj, wanted) {
  if (!obj || typeof obj !== 'object') return null;
  for (const k in obj) {
    if (String(k).toLowerCase() === wanted.toLowerCase()) return obj[k];
    const found = findKeyRecursive_(obj[k], wanted);
    if (found !== null) return found;
  }
  return null;
}


function extractWikiSyncDiaryData_(raw) {
  // WikiSync has exposed diary data under more than one response shape over
  // time. Search the full raw payload before/after normalization instead of
  // requiring one exact property name.
  if (!raw || typeof raw !== 'object') return {data:null, path:'', keys:''};

  const wanted = [
    'achievementdiaries',
    'achievementdiary',
    'diaries',
    'diaryprogress',
    'achievementdiaryprogress'
  ];

  const seen = [];
  function normKey_(s) {
    return String(s || '').toLowerCase().replace(/[^a-z0-9]/g,'');
  }

  function looksLikeDiaryMap_(value) {
    if (!value || typeof value !== 'object') return false;
    const txt = JSON.stringify(value).toLowerCase();
    const regions = ['ardougne','desert','falador','fremennik','kandarin','karamja',
      'kourend','lumbridge','morytania','varrock','western','wilderness'];
    const tiers = ['easy','medium','hard','elite'];
    return regions.some(r => txt.indexOf(r) >= 0) && tiers.some(t => txt.indexOf(t) >= 0);
  }

  function walk(node, path, depth) {
    if (!node || typeof node !== 'object' || depth > 8) return null;

    if (looksLikeDiaryMap_(node) && path.length) {
      const last = normKey_(path[path.length - 1]);
      if (wanted.indexOf(last) >= 0) {
        return {data:node, path:path.join('.'), keys:Object.keys(node).slice(0,16).join(', ')};
      }
    }

    for (const k in node) {
      const child = node[k];
      const nk = normKey_(k);
      seen.push(path.concat([k]).join('.'));

      if (wanted.indexOf(nk) >= 0 && child && typeof child === 'object') {
        return {data:child, path:path.concat([k]).join('.'), keys:Object.keys(child).slice(0,16).join(', ')};
      }
    }

    for (const k in node) {
      const child = node[k];
      if (child && typeof child === 'object') {
        const found = walk(child, path.concat([k]), depth + 1);
        if (found) return found;
      }
    }
    return null;
  }

  const found = walk(raw, [], 0);
  if (found) return found;

  // Last-resort: if a nested object itself clearly looks like a diary map,
  // accept it even when its parent key is unusual.
  function findShape(node, path, depth) {
    if (!node || typeof node !== 'object' || depth > 8) return null;
    if (path.length && looksLikeDiaryMap_(node)) {
      return {data:node, path:path.join('.'), keys:Object.keys(node).slice(0,16).join(', ')};
    }
    for (const k in node) {
      const child=node[k];
      if (child && typeof child === 'object') {
        const f=findShape(child,path.concat([k]),depth+1);
        if (f) return f;
      }
    }
    return null;
  }

  return findShape(raw, [], 0) || {data:null,path:'',keys:''};
}

function normalizeQuestState_(value) {
  const n = Number(value);
  if (n === 2) return 'Complete';
  if (n === 1) return 'In Progress';
  if (n === 0) return 'Incomplete';
  const s = String(value || '');
  if (/finished|complete/i.test(s)) return 'Complete';
  if (/progress|started/i.test(s)) return 'In Progress';
  if (/not.?started|incomplete/i.test(s)) return 'Incomplete';
  return 'Unknown';
}


function explainDiarySyncRequirement_() {
  SpreadsheetApp.getUi().alert(
    'Achievement Diary tracking is not synced yet.\n\n' +
    'In RuneLite:\n' +
    '1. Open WikiSync plugin settings.\n' +
    '2. Enable Achievement Diary / Diary tracking.\n' +
    '3. Open the in-game Achievement Diary interface once while logged in.\n' +
    '4. Run RuneScape GE → Sync WikiSync in Browser again.\n\n' +
    'WikiSync can provide diary tier data, but it must be enabled/synced first.'
  );
}

function refreshDiaries_(player, levels, syncData, syncMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Diary_Tracker');
  const rows = sh.getRange('A10:H57').getValues();
  const diaryInfo=extractWikiSyncDiaryData_(syncData);
  const diaryData = syncData && (syncData.achievement_diaries || syncData.achievementDiaries)
    ? (syncData.achievement_diaries || syncData.achievementDiaries)
    : diaryInfo.data;
  const diarySource = syncData && syncData.achievement_diaries_source
    ? syncData.achievement_diaries_source
    : diaryInfo.path;

  // Diagnostic cells: distinguish missing WikiSync diary data from parser issues.
  sh.getRange('J2').setValue('WikiSync Diary Data');
  sh.getRange('K2').setValue(diaryData ? 'Loaded' : 'Not present');
  sh.getRange('J3').setValue('Detected keys');
  sh.getRange('K3').setValue(diaryData && typeof diaryData === 'object'
    ? Object.keys(diaryData).slice(0,12).join(', ')
    : '');
  sh.getRange('J4').setValue('Diary source path');
  sh.getRange('K4').setValue(diarySource || (diaryData ? '' : 'Enable WikiSync diary tracking, open Achievement Diary once, then browser-sync again.'));


  const pageCache = {};
  const uniquePages = {};
  rows.forEach(r => uniquePages[r[0]] = true);

  Object.keys(uniquePages).forEach(region => {
    const page = diaryPageName_(region);
    try {
      const resp = UrlFetchApp.fetch(wikiApiUrl_(page), {
        muteHttpExceptions:true,
        headers:{'User-Agent':'OSRS-Training-GoogleSheet/10.14.9 - personal Google Sheets tool'}
      });
      if (resp.getResponseCode() === 200) {
        const data = JSON.parse(resp.getContentText());
        pageCache[region] = data.parse && data.parse.wikitext ? data.parse.wikitext['*'] : '';
      }
    } catch(e) {}
  });

  const out = rows.map(r => {
    const region = r[0], tier = r[1];
    const status = diaryStatus_(diaryData, region, tier);
    const wt = pageCache[region] || '';
    const tierText = extractTierSection_(wt, tier);
    const missing = missingSkills_(tierText, levels);
    const readiness = tierText ? (missing ? 'Missing' : 'Ready') : 'Unknown';
    const notes = tierText ? formatDiaryTier_(tierText) : syncMessage;
    const wiki = 'https://oldschool.runescape.wiki/w/' + encodeURIComponent(diaryPageName_(region).replace(/ /g,'_'));
    return [region,tier,status,readiness,missing,notes,wiki,new Date()];
  });

  sh.getRange(10,1,out.length,8).setValues(out);
  sh.getRange(10,8,out.length,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}


function stripNestedWikiTemplates_(text) {
  // Remove nested {{...}} templates safely by repeatedly stripping the
  // innermost template. This avoids raw template markup leaking into cells.
  let s = String(text || '');
  let prior = '';
  let guard = 0;
  while (s !== prior && guard++ < 30) {
    prior = s;
    s = s.replace(/\{\{[^{}]*\}\}/g, ' ');
  }
  return s;
}

function cleanDiaryCell_(text) {
  let s = String(text || '');

  // Remove comments/references.
  s = s
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<ref[^>]*>[\s\S]*?<\/ref>/gi, ' ')
    .replace(/<ref[^\/>]*\/>/gi, ' ');

  // Convert common requirement/item templates to readable text BEFORE
  // removing the remaining templates.
  s = s
    .replace(/\{\{\s*(?:SCP|Skillreq|Skill requirement|SkillReq)\s*\|\s*([^|}]+)\s*\|\s*(\d{1,3})[^}]*\}\}/gi,
             (_, skill, lvl) => String(lvl).trim() + ' ' + String(skill).trim())
    .replace(/\{\{\s*(?:Questreq|Quest requirement|Qreq)\s*\|\s*([^|}]+)[^}]*\}\}/gi,
             (_, quest) => String(quest).trim())
    .replace(/\{\{\s*(?:plink|item|inv|npc|monster|spell|location)\s*\|\s*([^|}]+)[^}]*\}\}/gi,
             (_, value) => String(value).trim())
    .replace(/\{\{\s*Coins?\s*\|\s*([^|}]+)[^}]*\}\}/gi,
             (_, value) => String(value).trim() + ' coins');

  s = stripNestedWikiTemplates_(s);

  // Convert links and remove HTML/wiki styling.
  s = s
    .replace(/\[\[(?:File|Image):[^\]]+\]\]/gi, ' ')
    .replace(/\[\[([^|\]#]+)(?:#[^|\]]*)?\|([^\]]+)\]\]/g, '$2')
    .replace(/\[\[([^\]#]+)(?:#[^\]]*)?\]\]/g, '$1')
    .replace(/\[https?:\/\/[^\s\]]+\s*([^\]]*)\]/g, '$1')
    .replace(/<br\s*\/?>/gi, '; ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/'{2,}/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&ndash;/gi, '–')
    .replace(/&mdash;/gi, '—')
    .replace(/&quot;/gi, '"')
    .replace(/\s+/g, ' ')
    .trim();

  return s;
}

function extractDiaryTaskRows_(tierText) {
  let text = String(tierText || '');

  // The first section on diary pages is often a collapsible {{Map|...}}
  // table. It is useful on the Wiki but unreadable in a spreadsheet, so
  // discard the entire collapsible block.
  text = text.replace(
    /\{\|\s*class\s*=\s*["'][^"']*mw-collapsible[^"']*["'][\s\S]*?\n\|\}\s*/gi,
    ''
  );

  // Prefer the actual achievement-diary task table.
  const tableMatch = text.match(
    /\{\|[^\n]*(?:diary-table|wikitable[^\n]*data-diary)[\s\S]*?\n\|\}/i
  );
  const table = tableMatch ? tableMatch[0] : text;

  // Drop table declaration and header rows, then split into task rows.
  const pieces = table.split(/\n\|-\s*(?:\n|$)/);
  const out = [];

  pieces.forEach(piece => {
    let p = String(piece || '').trim();
    if (!p || p.startsWith('{|')) {
      // A declaration chunk can also contain header lines; keep only content
      // after the last header if a task somehow shares this chunk.
      p = p.replace(/^\{\|[^\n]*\n?/, '');
    }

    // Remove ! header cells and table-closing syntax.
    p = p.replace(/^\s*![\s\S]*?(?=\n\|)/, '').replace(/\n\|\}\s*$/, '').trim();
    if (!p || p.startsWith('!')) return;

    // Diary tables normally have two cells: Task and Requirements.
    let cells = p.split(/\n\|\s*/).filter(Boolean);
    if (cells.length < 2) {
      cells = p.split(/\|\|/).filter(Boolean);
    }
    if (!cells.length) return;

    const task = cleanDiaryCell_(cells[0].replace(/^\|+/, ''));
    const req = cells.length > 1 ? cleanDiaryCell_(cells.slice(1).join(' ')) : '';

    // Ignore declarations/header leftovers and obvious non-task content.
    if (!task || /^(task|requirements?)$/i.test(task)) return;
    if (/^(class|style|data-diary)/i.test(task)) return;

    out.push({task:task, req:req});
  });

  return out;
}

function formatDiaryTier_(tierText) {
  const rows = extractDiaryTaskRows_(tierText);

  if (rows.length) {
    const lines = rows.map((row, i) => {
      let task = row.task
        .replace(/^\s*\d+\.\s*/, '')
        .trim();
      const req = row.req && !/^(none|n\/a|—|-)?$/i.test(row.req)
        ? ' — Req: ' + row.req
        : '';
      return (i + 1) + '. ' + task + req;
    });

    // Google Sheets cell limit is much larger, but keep diary notes compact
    // and easy to scan.
    return lines.join('\n').slice(0, 4500);
  }

  // Fallback for an unexpected Wiki layout: aggressively clean markup
  // rather than showing raw {|, {{Map}}, and template code.
  let fallback = String(tierText || '')
    .replace(/\{\|\s*class\s*=\s*["'][^"']*mw-collapsible[^"']*["'][\s\S]*?\n\|\}\s*/gi, ' ')
    .replace(/\{\|[\s\S]*?\|\}/g, ' ');
  fallback = cleanDiaryCell_(fallback);
  return fallback.slice(0, 4500);
}

function normalizeDiaryState_(value) {
  // WikiSync diary values are not quest-state integers. Handle common
  // booleans, numeric completion flags, strings, and wrapped objects safely.
  if (value === true) return 'Complete';
  if (value === false) return 'Incomplete';

  if (value && typeof value === 'object') {
    if (value.completed !== undefined) return normalizeDiaryState_(value.completed);
    if (value.complete !== undefined) return normalizeDiaryState_(value.complete);
    if (value.finished !== undefined) return normalizeDiaryState_(value.finished);
    if (value.state !== undefined) return normalizeDiaryState_(value.state);
    if (value.status !== undefined) return normalizeDiaryState_(value.status);
    if (value.value !== undefined) return normalizeDiaryState_(value.value);
  }

  const s = String(value == null ? '' : value).trim().toLowerCase();
  if (!s) return 'Unknown';

  if (['true','complete','completed','finished','done','yes'].includes(s)) return 'Complete';
  if (['false','incomplete','not started','not_started','notstarted','no'].includes(s)) return 'Incomplete';
  if (['in progress','in_progress','started','partial'].includes(s)) return 'In Progress';

  // Some APIs encode a completed tier as 1 and incomplete as 0.
  if (s === '1') return 'Complete';
  if (s === '0') return 'Incomplete';

  return 'Unknown';
}

function diaryStatus_(data, region, tier) {
  if (!data || typeof data !== 'object') return 'Unknown';

  const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9]/g,'');
  const targetRegion = norm(region);
  const targetTier = norm(tier);

  const regionAliases = {
    'ardougne':['ardougne','ardy'],
    'desert':['desert'],
    'falador':['falador','fally'],
    'fremennik':['fremennik'],
    'kandarin':['kandarin'],
    'karamja':['karamja'],
    'kourendkebos':['kourendkebos','kourendandkebos','kourend','kebos'],
    'lumbridgedraynor':['lumbridgedraynor','lumbridgeanddraynor','lumbridge','draynor'],
    'morytania':['morytania'],
    'varrock':['varrock'],
    'westernprovinces':['westernprovinces','western'],
    'wilderness':['wilderness','wildy']
  };

  const aliases = regionAliases[targetRegion] || [targetRegion];

  function pathMatches(path) {
    const compact = norm(path.join(' '));
    const hasRegion = aliases.some(a => compact.includes(a));
    const hasTier = compact.includes(targetTier);
    return hasRegion && hasTier;
  }

  function walk(value, path) {
    // If the accumulated path already identifies this exact diary tier,
    // decode the value at this node.
    if (pathMatches(path)) {
      const decoded = normalizeDiaryState_(value);
      if (decoded !== 'Unknown') return decoded;
    }

    if (Array.isArray(value)) {
      for (const item of value) {
        if (item && typeof item === 'object') {
          const label = item.name || item.area || item.region || item.diary || '';
          const diff = item.tier || item.difficulty || item.level || '';
          const nextPath = path.concat([label, diff].filter(Boolean));
          if (pathMatches(nextPath)) {
            const raw = item.completed !== undefined ? item.completed :
                        item.complete !== undefined ? item.complete :
                        item.finished !== undefined ? item.finished :
                        item.state !== undefined ? item.state :
                        item.status !== undefined ? item.status :
                        item.value !== undefined ? item.value : item;
            const decoded = normalizeDiaryState_(raw);
            if (decoded !== 'Unknown') return decoded;
          }
          const nested = walk(item, nextPath);
          if (nested !== 'Unknown') return nested;
        }
      }
      return 'Unknown';
    }

    if (value && typeof value === 'object') {
      for (const key in value) {
        const nextPath = path.concat([key]);
        const child = value[key];

        if (pathMatches(nextPath)) {
          const decoded = normalizeDiaryState_(child);
          if (decoded !== 'Unknown') return decoded;
        }

        const nested = walk(child, nextPath);
        if (nested !== 'Unknown') return nested;
      }
    }

    return 'Unknown';
  }

  return walk(data, []);
}

function diaryPageName_(region) {
  const map = {
    'Kourend & Kebos':'Kourend & Kebos Diary',
    'Lumbridge & Draynor':'Lumbridge & Draynor Diary',
    'Western Provinces':'Western Provinces Diary'
  };
  return map[region] || region + ' Diary';
}

function wikiApiUrl_(page) {
  return 'https://oldschool.runescape.wiki/api.php?action=parse&format=json&prop=wikitext&page=' + encodeURIComponent(page);
}

function extractSection_(wt, heading) {
  if (!wt) return '';
  const re = new RegExp('==+\\s*' + heading.replace(/[.*+?^${}()|[\]\\]/g,'\\$&') + '\\s*==+','i');
  const m = re.exec(wt);
  if (!m) return '';
  const rest = wt.slice(m.index + m[0].length);
  const next = rest.search(/\n==[^=]/);
  return next >= 0 ? rest.slice(0,next) : rest;
}

function extractTierSection_(wt, tier) {
  if (!wt) return '';
  const re = new RegExp('={2,4}\\s*' + tier + '\\s*={2,4}','i');
  const m = re.exec(wt);
  if (!m) return '';
  const rest = wt.slice(m.index + m[0].length);
  const next = rest.search(/\n={2,4}[^=]/);
  return next >= 0 ? rest.slice(0,next) : rest;
}

function cleanWiki_(s) {
  return String(s || '')
    .replace(/<ref[^>]*>[\s\S]*?<\/ref>/gi,' ')
    .replace(/<ref[^\/>]*\/>/gi,' ')
    .replace(/\{\{[^{}]*\}\}/g,' ')
    .replace(/\[\[(?:[^|\]]*\|)?([^\]]+)\]\]/g,'$1')
    .replace(/\[https?:\/\/[^\s\]]+\s*([^\]]*)\]/g,'$1')
    .replace(/'{2,}/g,'')
    .replace(/[=*#;]+/g,' ')
    .replace(/\s+/g,' ')
    .trim();
}

function extractQuestPrereqs_(s, currentQuest) {
  const text = String(s || '');
  const found = [];

  function add(name) {
    name = String(name || '')
      .replace(/<!--[\s\S]*?-->/g,'')
      .replace(/\[\[|\]\]/g,'')
      .split('|')[0]
      .split('#')[0]
      .trim();

    if (!name) return;
    if (currentQuest && name.toLowerCase() === String(currentQuest).toLowerCase()) return;
    if (/^(quest|quests|quest point|quest points|skill|skills|members)$/i.test(name)) return;
    if (/^(yes|no|none|n\/a)$/i.test(name)) return;
    if (!found.some(x => x.toLowerCase() === name.toLowerCase())) found.push(name);
  }

  const templatePatterns = [
    /\{\{\s*Questreq\s*\|\s*([^|}\n]+)(?:\|[^}]*)?\}\}/gi,
    /\{\{\s*Quest requirement\s*\|\s*([^|}\n]+)(?:\|[^}]*)?\}\}/gi,
    /\{\{\s*Qreq\s*\|\s*([^|}\n]+)(?:\|[^}]*)?\}\}/gi
  ];

  templatePatterns.forEach(rex => {
    let m;
    while ((m = rex.exec(text)) !== null) add(m[1]);
  });

  text.split(/\r?\n/).forEach(line => {
    if (!/quest|complete|completion|finished|requires?/i.test(line)) return;
    const links = line.match(/\[\[([^\]]+)\]\]/g) || [];
    links.forEach(x => {
      const inner=x.slice(2,-2).split('|')[0].split('#')[0].trim();
      if (!/skill|level|quest point|members|combat/i.test(inner)) add(inner);
    });
  });

  return found.join(', ');
}

function parseQuestSkillRequirements_(s) {
  const text = String(s || '');
  const skillNames = [
    'Attack','Defence','Strength','Hitpoints','Ranged','Prayer','Magic',
    'Cooking','Woodcutting','Fletching','Fishing','Firemaking','Crafting','Smithing',
    'Mining','Herblore','Agility','Thieving','Slayer','Farming','Runecraft','Hunter','Construction','Sailing'
  ];
  const reqs = {};

  function add(skill, level) {
    const canonical = skillNames.find(x => x.toLowerCase() === String(skill).toLowerCase());
    level = Number(level) || 0;
    if (!canonical || level < 1 || level > 99) return;
    reqs[canonical] = Math.max(reqs[canonical] || 0, level);
  }

  const templatePatterns = [
    /\{\{\s*SCP\s*\|\s*([^|}\n]+)\s*\|\s*(\d{1,3})/gi,
    /\{\{\s*Skillreq\s*\|\s*([^|}\n]+)\s*\|\s*(\d{1,3})/gi,
    /\{\{\s*Skill requirement\s*\|\s*([^|}\n]+)\s*\|\s*(\d{1,3})/gi,
    /\{\{\s*SkillReq\s*\|\s*([^|}\n]+)\s*\|\s*(\d{1,3})/gi
  ];
  templatePatterns.forEach(rex => {
    let m;
    while ((m = rex.exec(text)) !== null) add(m[1],m[2]);
  });

  skillNames.forEach(skill => {
    const escaped = skill.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
    const patterns = [
      new RegExp('(\\d{1,3})\\s*(?:\\[\\[)?' + escaped + '(?:\\]\\])?', 'gi'),
      new RegExp('(?:\\[\\[)?' + escaped + '(?:\\]\\])?[^\\d\\n]{0,35}(\\d{1,3})', 'gi')
    ];
    patterns.forEach(rex => {
      let m;
      while ((m = rex.exec(text)) !== null) add(skill,m[1]);
    });
  });

  return reqs;
}

function missingSkills_(s, levels) {
  const reqs = parseQuestSkillRequirements_(s);
  const misses = [];

  Object.keys(reqs).sort().forEach(skill => {
    const key = skill.toLowerCase();
    const have = Number(levels[key]) || 0;
    const need = reqs[skill];
    if (need > have) misses.push(skill + ' ' + have + '/' + need);
  });

  return misses.join(', ');
}

function allRequiredSkills_(s) {
  const reqs = parseQuestSkillRequirements_(s);
  return Object.keys(reqs)
    .sort()
    .map(skill => skill + ' ' + reqs[skill])
    .join(', ');
}

function refreshGEPrices() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const priceSheet = ss.getSheetByName('Price_Data');
  const settings = ss.getSheetByName('Settings');
  const base = 'https://prices.runescape.wiki/api/v1/osrs';
  const options = {muteHttpExceptions:true, headers:{'User-Agent':'OSRS-Training-GoogleSheet/10.14.9 - personal Google Sheets tool'}};

  const mappingResp = UrlFetchApp.fetch(base + '/mapping', options);
  const latestResp = UrlFetchApp.fetch(base + '/latest', options);
  if (mappingResp.getResponseCode() !== 200 || latestResp.getResponseCode() !== 200)
    throw new Error('GE API request failed.');

  const mapping = JSON.parse(mappingResp.getContentText());
  const latest = JSON.parse(latestResp.getContentText()).data;
  const rows = [['Item','Item ID','High (Buy)','Low (Sell)','High Time','Low Time','Mid Price']];
  mapping.forEach(item => {
    const p = latest[String(item.id)];
    if (!p) return;
    const high = p.high == null ? '' : p.high;
    const low = p.low == null ? '' : p.low;
    rows.push([item.name,item.id,high,low,p.highTime?new Date(p.highTime*1000):'',p.lowTime?new Date(p.lowTime*1000):'',(high!==''&&low!=='')?(high+low)/2:'']);
  });
  const header = rows.shift();
  rows.sort((a,b)=>String(a[0]).localeCompare(String(b[0])));
  rows.unshift(header);
  const lastPriceRow = Math.max(priceSheet.getLastRow(), rows.length);
  clearColumnsContent_(priceSheet, 1, lastPriceRow, 1, 7);
  priceSheet.getRange(1,1,rows.length,7).setValues(rows);
  priceSheet.getRange(1,1,1,7).setFontWeight('bold').setBackground('#17365D').setFontColor('#FFFFFF');
  settings.getRange('B14').setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}