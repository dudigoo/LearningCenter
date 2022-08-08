var g_func2run;

function iterateMain() {  
  //collectParams();
  //writeLog('Starting..' + g_func2run);
  for (var i in wfolders) {
    fixDir(wfolders[i]);
  }
  SpreadsheetApp.flush();
  writeLog('End');
} 

function fixDir(d) {
  Logger.log('work on '+ d[0]);
  var folder = DriveApp.getFolderById(d[0]); 
  var files = folder.getFilesByType('application/vnd.google-apps.spreadsheet');
  while (files.hasNext()) {
    file = files.next();
    fix_SS(file,folder);
  } 
}

function fix_SS(file,folder) {
  var key=file.getId();
  var tnm = file.getName();
  Logger.log('worker :'+tnm);
  var tabnm=tnm.replace(/^\S+ \S+ \S+ /, '');
  var ss = SpreadsheetApp.open(file);
  var w=getWorkerByName(tabnm);
  var func2run=g_func2run;
  
  //func2run='hard2';
  
  if (func2run=='fixss'){
    fixss(ss,file,folder,tabnm,tnm);   
  } else if (func2run=='shareWorkerSS') {
    shareWorkerSS(ss,w);
  } else if (func2run=='unshareWorkerSS') {
    unshareWorkerSS(ss,w);
  } else if (func2run=='switchActiveMonth') {
    switchActiveMonth(ss, tabnm);
  } else if (func2run=='delSomeRows') {
    delSomeRows(ss);
  } else if (func2run=='cp2maakav') {
    cp2maakav(file,ss,w);
  } else if (func2run=='add_monthes') {
    addMonthes(ss,w);
  } else if (func2run=='lock_month') {
    lockMonth(ss);
  } else {
    Logger.log("bad function nm:"+g_func_nm);
  }
}
