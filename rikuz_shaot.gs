
var month_row = { 9 : 3, 10 : 4 , 11 : 5, 12 : 6 , 1 : 7, 2 : 8 , 3 : 9 , 4 : 10 , 5 : 11 , 6 : 12 , 7 : 13, 8 : 14};

function updateRikuz(name, hours, snm, monthnum) {
  //Logger.log('updateRikuz hours='+ hours + ' name='+name + ' snm='+snm + ' monthnum='+monthnum);
  if (! month_row[monthnum]){
    writeLog('invalid month number for: '+name);
    return;
  }
  var found=0;
  var max_wrkrs=gp.rikuz_wrkrs[0].length;
  for (var i=0; i<max_wrkrs; i++){
    var nm=gp.rikuz_wrkrs[0][i];//g_rikuz_sheet[snm].getRange(2, i).getValue();
    if (name == nm) {
      var snum=1;
      if (snm == 'merkaz'){
        snum=0;
        //Logger.log('name='+name+' len='+gp.rikuz_dat[snum].length + ' i='+i + ' hours=' + hours+ ' monthnum=' + monthnum+ ' snum=' + snum);
        //Logger.log( ' month_row[monthnum]=' + month_row[monthnum]);
        var oldval=gp.rikuz_dat[snum][month_row[monthnum]-3][i]; 
        //g_rikuz_sheet[snm].getRange(month_row[monthnum],i).getValue();
        if (! oldval) {oldval=0;}  
        hours += oldval;
      }
      gp.rikuz_dat[snum][month_row[monthnum]-3][i]=hours;
      //Logger.log('updated: name='+name + ' i='+i + ' hours=' + hours);
      found=1;
      break;
    }
  }
  if (found==0){
    if (! gp.rikuz_nms_not_found){
      gp.rikuz_nms_not_found={};
    }
    if (! gp.rikuz_nms_not_found[name]){
      gp.rikuz_nms_not_found[name]=1;
      writeLog('name='+name + ' hours=' + hours + ' not found in rikuz...');
    }
  }
}

function saveRikuzData() {
  for (var i=0;i<2;i++){
    gp.rikuz_dat_rng[i].setValues(gp.rikuz_dat[i]);
    Logger.log('gp.rikuz_dat_keep[i]='+gp.rikuz_dat_keep[i]);
    gp.rikuz_dat_keep_rng[i].setFormulas(gp.rikuz_dat_keep[i]);
  }
}

function loadRikuzData() {
  var rikuz_ss = SpreadsheetApp.openById(gp.rikuz_file_id);
  gp.rikuz_sheets=[];
  gp.rikuz_sheets.push(rikuz_ss.getSheetByName('merkaz'));
  gp.rikuz_sheets.push(rikuz_ss.getSheetByName('accounting'));
  var lcol=gp.rikuz_sheets[0].getLastColumn()-4;
  gp.rikuz_wrkrs=gp.rikuz_sheets[0].getRange(2,3,1,lcol).getValues();
  gp.rikuz_dat=[]; gp.rikuz_manual_dat=[]; gp.rikuz_dat_rng=[]; gp.rikuz_dat_keep=[]; gp.rikuz_dat_keep_rng=[];
  for (var i=0;i<2;i++){
    gp.rikuz_dat_keep_rng.push(gp.rikuz_sheets[i].getRange('CV3:CV14'));// 76 , 86
    gp.rikuz_dat_keep.push(gp.rikuz_dat_keep_rng[i].getFormulas());
    Logger.log('load gp.rikuz_dat_keep[i]='+gp.rikuz_dat_keep[i]);
    gp.rikuz_dat_rng.push(gp.rikuz_sheets[i].getRange(3,3,12,lcol));
    if (gp.g_month_name == 'all'){
      gp.rikuz_dat_rng[i].clear();
    }
    gp.rikuz_dat.push(gp.rikuz_dat_rng[i].getValues());
    if (i==0){
      gp.rikuz_manual_dat.push(gp.rikuz_sheets[i].getRange(48,3,12,lcol).getValues());
      gp.rikuz_manual_dat.push(gp.rikuz_sheets[i].getRange(48,1,12,1).getValues());
    }
  }
  //Logger.log('loadRikuzData: rikuz_dat0='+gp.rikuz_dat[0]);
}

function loadManualReports(ta) {
  man_ar=gp.rikuz_manual_dat[0];
  //Logger.log('gp.rikuz_manual_dat[1]='+JSON. stringify(gp.rikuz_manual_dat[1]));
  for (let i=0; i<man_ar.length;i++){
    for (let j=0; j<man_ar[i].length;j++){
      if (man_ar[i][j]>0){
        //Logger.log('i='+i+' j='+j+' gp.rikuz_manual_dat[1]='+gp.rikuz_manual_dat[1].length);
        //Logger.log('gp.rikuz_manual_dat[1][j][0]='+gp.rikuz_manual_dat[1][i][0]);
        updateRikuz(gp.rikuz_wrkrs[0][j], man_ar[i][j], 'merkaz', gp.rikuz_manual_dat[1][i][0]);
        updateRikuz(gp.rikuz_wrkrs[0][j], man_ar[i][j], 'accounting', gp.rikuz_manual_dat[1][i][0]);
      }
    }
  }
}

function rikuzMain() {
  collectParams();
  loadRikuzData();
  var hfiles=getSubFoldersFiles(gp.top_accounting_dir_id,'rikuz');
  var ta={};
  loadManualReports(ta);
  for (var i=0; i<hfiles.length; i++){
    processHfile(hfiles[i], ta);
  }
  for (const [wrkr, wrkh] of Object.entries(ta)) {
    for (const [month, h] of Object.entries(wrkh)) { 
      //Logger.log('wrkr='+wrkr+' month='+month);
      updateRikuz(wrkr, h, 'merkaz', month);
    }
  }
  saveRikuzData();
  checkLog();
  SpreadsheetApp.flush();
}  

function checkSkip(fo) {
  try {
    let fomon=getMonthFromStr(fo.getName());
    Logger.log('folder name='+fo.getName()+ ' mon0='+fomon[0]+ ' mon1='+fomon[1]);
    //Logger.log('gp.g_month_name='+gp.g_month_name);
    if ( gp.g_month_name != 'all' && ! gp.g_month_name_ar.includes(fomon[0])){
      Logger.log('skip');
      return 1;
    }
  } catch (e) {
    writeLog('skip invalid folder name: '+fo.getName()+ ' error='+e);
    return 1;
  }
  return 0;
}  

function getSubFoldersFiles(folder_id,client) {
  //Logger.log('fid='+folder_id);
  var folder = DriveApp.getFolderById(folder_id);
  var foi=folder.getFolders();
  var hfiles=[];
  while (foi.hasNext()) {
    fo = foi.next();
    if (client == 'rikuz' && checkSkip(fo)){
      continue;
    }
    var files = fo.getFilesByType('application/vnd.google-apps.spreadsheet');
    while (files.hasNext()) {
      var file = files.next();
      hfiles.push(SpreadsheetApp.open(file));
    } 
  }
  return hfiles;
}

function getMonthFromStr(s) {
  var mon_str=s.match(/\d+\.\d+\-\d+\.\d+/)[0];
  if (!mon_str.length) {
    return [0,0];
  }
  var mon_num=mon_str.match(/\d+$/)[0];
  return [mon_str, mon_num];
}  

function processHfile(ss,ta) {
  shts=ss.getSheets();
  var ssnm=ss.getName();
  //Logger.log('processHfile: name='+ssnm)
  var m=getMonthFromStr(ssnm);
  if (! m[0]){
    writeLog('file w/o month in name:'+ssnm);
    return;
  }
  //var acc_mon=m[0];
  var acc_mon_num=m[1];
  for (var i=0; i<shts.length; i++){
    var wrkrnm=shts[i].getName();
//    if (gp.rikuz_wrkrs_filter && ! gp.rikuz_wrkrs_filter_ar.includes(wrkrnm)){ //mmm
//      continue; 
//    }
    //Logger.log('wrkrnm='+wrkrnm);
    if (! ta.hasOwnProperty(wrkrnm)){
      ta[wrkrnm]={};
    }
    var ts=getHTotalHours(shts[i],ta[wrkrnm]);
    //Logger.log( ' ta='+JSON.stringify(ta)) //mmm
    //Logger.log('acc_mon_num='+acc_mon_num+' wrkrnm='+wrkrnm) //mmm
    updateRikuz(wrkrnm, ts, 'accounting', acc_mon_num);
  }
}

function getHTotalHours(tsheet,tw){
  //Logger.log('tsheet='+tsheet.getName());
  var trow=findTotalRow(tsheet);
  var tots=tsheet.getRange(trow, 5,2,1).getValues();
  var total_sheet=tots[0][0];//tsheet.getRange(trow, 5).getValue();
  if (! total_sheet) {
    return 0;
  }
  var total_sheetac=tots[1][0];//tsheet.getRange(trow+1, 5).getValue();
  var ratio=1;
  if (total_sheetac>total_sheet) {
    ratio=total_sheetac/total_sheet;
    total_sheet=total_sheetac;
  }
  let sheet_total=0;
  var prev_date;
  var prev_subj='';
  //Logger.log('getHTotalHours: trow='+ trow +' name='+tsheet.getName()+' total_sheet='+total_sheet+ ' total_sheetac='+total_sheetac);
  var sha=tsheet.getRange(8, 2,trow-8,6).getValues();
  var wrkrnm=tsheet.getName();
  for (var i=0; i < trow-8; i++) {
    
    //prev dt and subj
    if (! sha[i][5]){
      sha[i][5]=prev_subj;
    }
    prev_subj=sha[i][5];
    var dt1=sha[i][0];//tsheet.getRange(i, 2).getValue();
    dt1 = dt1 ? dt1 : prev_date;
    prev_date=dt1;
    var dt= new Date(dt1);
    var mon=(dt.getMonth()+1).toString();
    
    // filter
    if (gp.rikuz_grade_filter_ar && ! gp.rikuz_grade_filter_ar.includes(sha[i][4])){
      continue;
    }
    if (gp.rikuz_subjects ) {
      if (gp.rikuz_subjects_omit != 'y' && ! gp.rikuz_subjects_ar.includes(sha[i][5]) ) {
        if (! gp.rikuz_wrkrs_filter || gp.rikuz_wrkrs_filter_ar.includes(wrkrnm)){
          continue;
        }
      }
      if (gp.rikuz_subjects_omit == 'y' && gp.rikuz_subjects_ar.includes(sha[i][5]) ) {
        if (! gp.rikuz_wrkrs_filter || gp.rikuz_wrkrs_filter_ar.includes(wrkrnm)){
          continue;
        }
      }
    }
    //add hours
    var hrs;
    hrs=sha[i][3];//tsheet.getRange(i, 5).getValue();
    if (hrs){
      if (! tw.hasOwnProperty(mon)){
        tw[mon]=0;
      }
      tw[mon] += hrs*ratio;
      sheet_total+=hrs;
      //Logger.log('hours='+ hrs + ' mon='+mon+ ' row='+i+ ' dt1='+dt1+' prev_date='+prev_date+' tw='+JSON.stringify(tw)); //mmm
    }
    //Logger.log('hours='+ hrs + ' mon='+mon+ ' row='+i+ ' dt1='+dt1+' tw='+JSON.stringify(tw)); //mmm
    //Logger.log('hrs='+ hrs + ' total='+total+ ' row='+i+ ' dt1='+dt1+' month='+mon);
  }
  //Logger.log('getHTotalHours: total_sheet='+ total_sheet + ' tw='+JSON.stringify(tw)); //mmm
  return Math.round(sheet_total*ratio);
  //return total_sheet;
}
 
function rikuz2pikuach() {
  collectParams();
  var ro=7;
  var sh_rik= SpreadsheetApp.openById(gp.nizul_src).getSheetByName('merkaz');
  var sh_niz= SpreadsheetApp.openById(gp.nizul_tgt).getSheetByName('sheet1');
  for (var i=3;i<100;i++){
    var nm=sh_rik.getRange(2,i).getValue();
    var h=sh_rik.getRange(15,i).getValue();
    h=Math.round(h);
    if (nm == '' || !h) {
      continue;
    }
    var w=getWorkerByName(nm);
    if (!w){
      writeLog('no worker:'+nm);
      continue;
    }
    Logger.log('wrkr:'+JSON.stringify(w));
    sh_niz.getRange(ro,1).setValue(nm);
    sh_niz.getRange(ro,2).setValue(w.subj);
    sh_niz.getRange(ro,3).setValue(w.popu);
    sh_niz.getRange(ro,4).setValue(sh_rik.getRange(16,i).getValue());
    sh_niz.getRange(ro,5).setValue(h);
    //sh_niz.getRange(ro,6).setValue(sh_rik.getRange(17,i).getValue() );
    ro++;
    if (nm == 'דותן חן') {
      break;
    }
  }
  checkLog();
}