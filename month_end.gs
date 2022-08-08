var wtyp_s = 'סטודנטים';  var wtyp_m = 'מורים'; var wtyp_h = 'חניכים';
var workers_copied=0;
//log_row=3;
var ma_ss;
var ma_ss_id;
var g_editors=['mlemida.ryam@gmail.com','dudigoo@gmail.com'];
var tete;

var wfolders = [];
var gp={};
gp.dates_tz = "Asia/Jerusalem";
//gp.locale='en-US';
gp.locale='he-IL';

function mergeMonthSheets() {
  collectParams();
  workers_copied=0;
  for (var i in wfolders) {
    copyDirSheets(wfolders[i]);
  }
  checkLog();
  SpreadsheetApp.flush();
}  

function cleanup() {
  var ma_sheet0 = ma_ss.getSheetByName('Sheet1');
  if ( ma_sheet0 != null ) {
    try {
      ma_ss.deleteSheet(ma_sheet0);
    } catch (e) {
      Logger.log('e=' + e);
    }           
  }
  ma_sheet0 = ma_ss.getSheetByName('גיליון1');
  if ( ma_sheet0 != null ) {
    try {
      ma_ss.deleteSheet(ma_sheet0);
    } catch (e) {
      Logger.log('e=' + e);
    }           
  }
}
  
function createMonthlyFile(typ) {
  var prefix = 'נועם מדר'; 
  var name = prefix + ' ' + typ + ' ' + gp.g_month_name;

  Logger.log('create file ' + name);
  let out_folder = DriveApp.getFolderById(gp.out_folder_id);  
  let file=SpreadsheetApp.create(name);
  ma_ss_id = file.getId();
  var copyFile=DriveApp.getFileById(ma_ss_id);
  out_folder.addFile(copyFile);
  DriveApp.getRootFolder().removeFile(copyFile);
  ma_ss = SpreadsheetApp.openById(ma_ss_id);
  Logger.log('monthly file name '+ name);
}

function copyDirSheets(d) {
  Logger.log('work on '+ d[0]);
  Logger.log('mon_nm: '+ gp.g_month_name); 
  createMonthlyFile(d[1]);
  workers_copied=0;
  var folder = DriveApp.getFolderById(d[0]); 
  var files = folder.getFilesByType('application/vnd.google-apps.spreadsheet');
  while (files.hasNext()) {
    file = files.next();
    copyPersonSheet(file.getId());
  }
  if (workers_copied > 0) {
    Logger.log('workers_copied=' + workers_copied);
    cleanup();
    spreadsheetToPDF(d[1]);
  } else {
    writeLog('all empty');
  }
}

function copyPersonSheet(key) {
  var tss = SpreadsheetApp.openById(key);
  var tnm = tss.getName();
  var tabnm=tnm.replace(/^\S+ \S+ \S+ /, '');
  Logger.log('gp.g_month_name='+gp.g_month_name+ ' person='+tabnm+ ' fnm='+tnm);
  var tsheet;
  try {
    var monthnm=gp.g_month_name;
    //Logger.log('month='+ monthnm);
    tsheet = tss.getSheetByName(monthnm);  
    if (tsheet == null) {
      Logger.log('no sheet '+ monthnm);
    } else {
      var trow=findTotalRow(tsheet);
      if (! trow){
        return;
      }
      var tot = tsheet.getRange(trow,5).getValue();
      if (! tot) {
        writeLog('empty:'+tabnm);
        Logger.log('skipping empty:' + tabnm);
        return;
      }
      setTextDirection(tsheet, trow);
      //Logger.log('copying to merged');
      var m_nsheet = tsheet.copyTo(ma_ss);
      Logger.log('copyed to merged');
      if (gp.monthly_thin){
        m_nsheet.getRange(7,8,trow-7,1).clear().setBorder(false, true, false, false, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID).clearDataValidations();
      }
      deleteHReportEmptyRows(m_nsheet, trow);
      //Logger.log('deleteHReportEmptyRows e');
      workers_copied = workers_copied + 1;
      m_nsheet.showSheet();
      m_nsheet.setName(tabnm);
      m_nsheet.getRange(5, 2).setValue(tsheet.getRange(5, 2).getValue());
      m_nsheet.getRange(5, 7).setValue(tsheet.getRange(5, 7).getValue());
      try { 
        m_nsheet.deleteColumns(9, m_nsheet.getMaxColumns()-8);
      } catch (e) {
        writeLog('err delete columns 9-11: e='+e+' name:'+tabnm);
      }
    }
  } catch (e) {
    writeLog('cpEr:'+tabnm);
    Logger.log('e=' + e);
  }
}

function deleteHReportEmptyRows(sh, trow) {
  var wh_ar=sh.getRange(8, 1,trow-8,8).getValues();
  //wh_ar[i].forEach(e => e[4]='');
  let i;
  for (i=wh_ar.length-1; i>=0; i--){
    if (wh_ar[i].join("").length>1){
      break;
    }
  }
  //Logger.log('wh_ar.length='+wh_ar.length+' i=' + i+' trow='+trow);
  if ((trow-8-i) > 2) {
    var frr=i+9;
    var num=trow-frr-1;
    //Logger.log('deleted rows frr=' + frr+' num='+num);
    sh.deleteRows(frr, num);
  }
}

function setTextDirection(sh,trow){
  sh.getRange(8,6,trow-8,3).setTextDirection(SpreadsheetApp.TextDirection.RIGHT_TO_LEFT);
}

function findTotalRow(sh){
  var lr=sh.getLastRow();
  var i=0;
  var trv=sh.getRange(lr-7,1,7,1).getValues();
  for (i=0;i<trv.length;i++){
    if (trv[i][0].substring(0,1)=='ס'){
      break;
    }
  }
  if ( i == trv.length) {
    writeLog('missing total row. name='+sh.getParent().getName()+' sheet='+sh.getName());
    return 0;
  }
  return i+lr-7;
}

function spreadsheetToPDF(wtype) {
// before run, update folder_id (where to save pdf, copy it from folder url) and file2export_id
//  var folder_id = '1yMwJSMXWvCr-8_hYqJ4__u4wqKsNse83'; 
//  var file2export_id = '1CtqXpkr3Yl0tZ-wvnZav3V-1JeLnlqR2qgQzHcv597o';  

//  var Logger.log('rcd='+rcd); ss = SpreadsheetApp.openById(file2export_id);
  Logger.log(ma_ss.getName());
  var url = ma_ss.getUrl();
  url = url.replace(/edit$/, '');
//  var timestamp = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'-'HHmm");
  var name = wtype + ' ' + gp.g_month_name + ' ' + '.pdf';  

  var token = ScriptApp.getOAuthToken();
  var request = { headers: { 'Authorization': 'Bearer ' +  token } };
  var params = '?fitw=true&exportFormat=pdf&format=pdf&size=7&sheetnames=true&scale=2&portrait=true&sheetnames=false&printtitle=false&gridlines=false';
  //var params = '?fitw=true&exportFormat=pdf&format=pdf&size=7&sheetnames=true&scale=2&portrait=true&sheetnames=false&printtitle=false&gridlines=false&r2=60&c1=0&c2=4';  
  var furl = url + "export"+params
  //Logger.log('furl=' + furl + ' out_folder_id=' + out_folder_id); 
  var blob = UrlFetchApp.fetch(furl, request);
  var rcd = blob.getResponseCode()
  
  var pdf = blob.getBlob().setName(name).getAs('application/pdf');
  var folder = DriveApp.getFolderById(gp.out_folder_id);
  folder.createFile(pdf);
  Logger.log('Done pdf');
}

function sheetProtection(sh,type) {//1=editable hour report 2=view only 3=invisible 4=shibutz editable
  var mxc=sh.getMaxColumns();
  var mxr=sh.getMaxRows();
  var lr=sh.getLastRow();
  rmProtections(sh);
  if (type==1) {
    let trow=findTotalRow(sh);
    if (! trow){return;}
    var p=sh.protect();
    let rng1='A8:M'+(trow-1);
    let rng2='H'+trow+':H'+(trow+1);
    var edit_rngs=[sh.getRange(rng1), sh.getRange(rng2)];
    p=p.removeEditors(p.getEditors()).addEditors(g_editors);
    p.setUnprotectedRanges(edit_rngs);
  } else if (type==4) {
    var p=sh.protect();
    var edit_rngs=[sh.getRange('E2:Q'+lr)];
    p=p.removeEditors(p.getEditors()).addEditors(g_editors);
    p.setUnprotectedRanges(edit_rngs);    
  } else if (type==2) {
    var view_rng=sh.getRange(1,1,mxr,mxc);
    var p=view_rng.protect();
    p.removeEditors(p.getEditors()).addEditors(g_editors);
  } else {
    var p=sh.protect();
    p.removeEditors(p.getEditors()).addEditors(g_editors);
  }
  return;
}

function rmProtections(sh){
  var protections = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    var protection = protections[i];
    protection.remove();
  }
  var protection = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if (protection) {
    protection.remove();
  }
}

function hideAllSheetsExcept(ss,sheetName) {
  var sheets=ss.getSheets();
  ss.getSheetByName(sheetName).showSheet();
  for(var i =0;i<sheets.length;i++){
    //Logger.log(i);
    if(sheets[i].getName()!=sheetName){
      sheets[i].hideSheet();
    }
  }
}

function lockMonthMain(){
  g_func2run='lock_month';
  collectParams();
  iterateMain();
}

function lockMonth(ss){
  sh = ss.getSheetByName(gp.g_month_name); 
  sheetProtection(sh,2);
}