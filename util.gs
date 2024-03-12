
  async function getAccountId() { /* ... */ }
  async function getAvailableNumbers(accountId) { /* ... */ }
  async function buyPhoneNumber(accountId, number) { /* ... */ }
  async function sendSMS(accountId, phoneId, to, msg) { /* ... */ }
  async function deletePhoneNumber(accountId, phoneNumberId) { /* ... */ }
function tst5() {
let t=[[9,,2,3,4],[9,,2],[9,,2,,4],[9,,2,,4]];
let a2=findEmptyColumns(t,1);
  Logger.log('x='+a2);
  dropColumns(t,a2);
  Logger.log('xd='+JSON.stringify( t));

   return;
}

function tst6() {
  collectParams();
  let a=getFmtDt(gp.shib_dates);

  Logger.log('x1='+a);

  return;
  Logger.log('maad='+maad.getTime());
  if (maad.getTime()==d.getTime()){
    Logger.log('same');
  } else{
    Logger.log('not same');
  }
  y=[1];
  //y=y.concat(x[0]);
  //Logger.log('yl='+y.length);
}

function delSomeRows(ss) {
  var sh=ss.getSheetByName('16.12-15.1');
  sh.deleteRows(20, 10);
  sh=ss.getSheetByName('16.1-15.2');
  sh.deleteRows(20, 10);
}

function fixMain(){
  g_func2run='fixWorkerTypInHourReport';
  collectParams();
  iterateMain();
} 

function fixWorkerTypInHourReportMain(){
  g_func2run='fixWorkerTypInHourReport';
  collectParams();
  iterateMain();
} 

function fixss(ss,file,folder,tabnm,tnm) {
  let sh=ss.getSheetByName('name');
  sh.getRange('C1').setValue('סוג');
  //Logger.log('wrkrsug='+JSON.stringify(getWorkerByName(sh.getRange('B2').getValue())));
  sh.getRange('C2').setValue(getWorkerByName(sh.getRange('B2').getValue()).typ);
  sh=ss.getSheetByName('1.2-31.2');
  sh.getRange('H5').setValue('=name!$C$2');
}  

function removeCopyOfPrefix () {
    var foldremoveCopyOfPrefixer_id = '1DFR2AzqaDAEIOgTYfs-fLSjkJT_MmVGl'; 
  
    var folder = DriveApp.getFolderById(folder_id); 
    var files = folder.getFilesByType('application/vnd.google-apps.spreadsheet');
    while (files.hasNext()){
      var file = files.next();
      var new_nm = file.getName().slice(8);
      file.setName(new_nm);
    }
}
