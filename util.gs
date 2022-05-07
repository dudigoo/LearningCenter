
function tst5() {
  //var lastUpdated = DriveApp.getFileById('1YP5aziOBgpzO1GS35z3yhY2op0qXdz2W').getLastUpdated();
  collectParams();
  let query = deleteObsoleteRecurRows();
  //let a=querySheet2(query,'1yrL132sLyUUzRruG5EzivGOk8uC88p7KPRC9NwAWI6A','groupPupil');
  //    Logger.log('a:'+a);
  //let d=SpreadsheetApp.openById('1aYDdx3zTQ2__HkRJIoK1NVnkzVzUf_KwzjPOtoE05BI').getSheetByName('recur').getRange(2,17,1,2).getValues()[0];

  //let dt=getDtObj('25/4/22');
  //let x=isDtInRng(dt,d[0],d[1]);
  //Logger.log('x='+JSON.stringify(x));
 // d.setDate( d.getDate()+7)
  //Logger.log('d1='+d);
  //Logger.log('dx='+getYMDStr(d));

 // const offset = d.getTimezoneOffset()
//d = new Date(d.getTime() - (offset*60*1000))
  //Logger.log('d2='+d.toISOString().split('T')[0]);

  //let x={'row2add':md,'name':nm,'sh2look':'allDays', 'sts':'ok'};
  //let p=getShibRecurAr();
  let level='ט';
  //let ed=getPupilEducator(pupil,level);
    //let q = getAlfonKids();
    //let query = 'select A,B,C,D,E,F,G,H where A="יא" and D=1 and E="054-558-1233"';

  //let query = 'select A,B where B="09-8664148"';
  //let query = 'select A,B where 1=1';
  //let a=querySheet(query,'115u0pg6db6muE-3raRAHppvsIRNcnUgRKeqxA8NIVKI','schoolTeachers');
  //str.shift();
  //let str='אב-חוה';
  //let a=getPupilsInGroup(str);
  //Logger.log('q='+lastUpdated );
//  let b=a.join(',');
  //Logger.log('a1='+b);


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
  g_func2run='fixss';

  collectParams();
  iterateMain();
} 

function fixss(ss,file,folder,tabnm,tnm) {
  var sh=ss.getSheetByName('lists');
 let ur='=importrange("https://docs.google.com/spreadsheets/d/1yrL132sLyUUzRruG5EzivGOk8uC88p7KPRC9NwAWI6A","low!t1:ac99")';
 sh.getRange(1,1).setValue(ur);
 var namedRanges = ss.getNamedRanges();
for (var i = 0; i < namedRanges.length; i++) {
  if (namedRanges[i].getName() == 'ז'){ 
    let r=ss.getRange("lists!A2:A99");
    namedRanges[i].setRange(r);  
  }
  if (namedRanges[i].getName() == 'ח'){ 
    let r=ss.getRange("lists!B2:B99");
    namedRanges[i].setRange(r);  
  }
  if (namedRanges[i].getName() == 'ט'){ 
    let r=ss.getRange("lists!C2:C99");
    namedRanges[i].setRange(r);  
  }
  if (namedRanges[i].getName() == 'י'){ 
    let r=ss.getRange("lists!D2:D99");
    namedRanges[i].setRange(r);  
  }  
  if (namedRanges[i].getName() == 'יא'){ 
    let r=ss.getRange("lists!E2:E99");
    namedRanges[i].setRange(r);  
  }
  if (namedRanges[i].getName() == 'יב'){ 
    let r=ss.getRange("lists!F2:F99");
    namedRanges[i].setRange(r);  
  }
}
 Logger.log(' tabnm');
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
