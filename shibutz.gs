var dowmap={'0':'א', '1':'ב', '2':'ג', '3':'ד', '4':'ה','5':'ו','6':'ז'};
var dowrmap={'א':'0', 'ב':'1', 'ג':'2', 'ד':'3', 'ה':'4','ו':'5','ז':'6'};

//var dowmap2={'0':'1', '1':'2', '2':'3', '3':'4', '4':'5','5':'6','6':'7'};
var gv_wrkr_view_tab_color='#1129e9';
//var svrow=1;svcol=14;

function fillShibutzMain() {// 22/12/20,24/12/20
  //Logger.log('locale 3='+gp.locale);
  collectParams();
  writeLog('Start');
  if (gp.shib_dates == 'all'){
    updateShibSheetsWork(9,1);
  } else if (gp.shib_dates == 'all2'){
    updateShibSheetsWork(9,10);
  } else {
    shibutzDates(getShibutzSS());
  }
  checkLog();
}

function getShibutzSS() {
  if (! gp.shibutz_ss){
    gp.shibutz_ss=SpreadsheetApp.openById(gp.shibutz_file_id);
  }
  return gp.shibutz_ss;
}

function clearShibutzDatesMain() {
  collectParams();
  writeLog('Start');
  clearShibutzDates()
  checkLog();
}

function sendReminderMain() {
  collectParams(3);
  writeLog('Start');
  let sh=getTodayShibSh();
  sendReminders(sh);
  checkLog();
}

function clearShibutzDates() {
  let shib_ss=getShibutzSS();
  getRecurSh().showSheet();
  let dts=gp.shib_dates.split(",");
  let sha=shib_ss.getSheets();
  for (var i=0;i<sha.length;i++){
    let nm= sha[i].getName();
    let rcd=getArchAction(nm,dts);
    Logger.log('check rcd='+rcd+' nm='+nm);
    if (!rcd){
      continue;
    }
    if (rcd=='arch'){
      copyDtSh2Hist(sha[i]);
      //var shn=sha[i].copyTo(ss_hist);
      //try {
      //  shn.setName(nm);
      //} catch (e) {
        //writeLog('error:'+e+ ' sheet already in archive:'+nm);
      //}
      Logger.log('after cp2hist. '+nm);
    }
    Logger.log('delete:'+nm);
    shib_ss.deleteSheet(sha[i]);
  }
}

function copyDtSh2Hist(sh){
  if (! gp.shib_history_sh){
    gp.shib_history_sh=getShibutzSS().getSheetByName('history');
  }
  let rows=sh.getLastRow()-1;
  let cols=sh.getLastColumn();
  let ar=sh.getRange(2,5,rows,cols-4).getValues();
  let hist_ilr=gp.shib_history_sh.getLastRow()+1;
  let hist_lr=hist_ilr;
  for (let i=0;i<ar.length;i++){
    if (! ar[i].every(element => element? false:true)){
      let trng=gp.shib_history_sh.getRange(hist_lr, 2, 1, cols);
      hist_lr+=1;
      //Logger.log('hist_lr='+hist_lr+ ' i+1='+(i+1)+' cols'+cols);
      sh.getRange(i+2,1,1,cols).copyTo(trng);
    }
  }
  let added=hist_lr-hist_ilr;
  if (added) {
    let dt=getDtObjFromTabNm(sh.getName());
    let dt_ar=[];
    dt_ar[added-1]=[dt];
    dt_ar.fill([dt]);
    //Logger.log('dt='+dt+' dt len='+dt_ar.length+ ' tbnm='+sh.getName());
    tdtrng=gp.shib_history_sh.getRange(hist_ilr,1,added,1).setValues(dt_ar);
    SpreadsheetApp.flush;
  }
}

function getArchAction(nm,dts){ //tabname
    var re =  /(\d+\/\d+\/\d+)/m; 
    match = re.exec(nm);
    if (! match){//no dt in name
      return;
    }
    var dt=match[1];
    if (! dts.includes(dt)){//dt not requested
      return;
    }

    if (nm.match(/ T /)){
      return 'del';
    }

    if (nm.match(/pre:/)){
      return 'del';
    }
    return 'arch';
}

function shibutzDates(shib_ss) {
  var dts=chomp(gp.shib_dates).split(",");
  Logger.log('dtsb='+gp.shib_dates);
  //return; //mmm
  for (var i=0;i<dts.length;i++){
    var sh=createDtSh(shib_ss,dts[i]);
    Logger.log('dt='+dts[i]+' sh nm='+sh.getName());
    fillShibDt(dts[i],sh,shib_ss);
  }
  Logger.log('S orderShibSheets');
  orderShibSheets(shib_ss);
  updateAllDatesSheetWork();
  Logger.log('E orderShibSheets');
  deleteObsoleteRecurRows();
}

function orderShibSheets(ss) { 
  var shs=ss.getSheets();
  var shs2s=shs.map(mkSortable);
  shs2s.sort(compareString);
  let re =  /^(\D)/m; 
  for (let i=0;i<shs2s.length;i++){
    let sh=ss.setActiveSheet(shs2s[i][0]);
    let match = re.exec(shs2s[i][1]);
    if (match && sh.getName() != 'recur'){
      sh.hideSheet();
    }
    //var op=sh.getIndex();
    ss.moveActiveSheet(i+1);
    //Logger.log('i='+i+' snm='+shs2s[i][1]);
  }
}

function mkSortable(e) { //tabname
  let nm=e.getName();
  //Logger.log('nm='+nm);
  let d=getDtObjFromTabNm(nm);
  //Logger.log('d='+d);
  if (d){
    let tim=d.getTime();
    if (tim){
      nm=nm.replace(/.*\d+\/\d+\/\d+/,tim);
    }
  } 
  nm = nm.replace(/^(.*:.*)$/,'z$1'); 

  //Logger.log('mkSortable nm='+nm);
  return [e,nm];
}

function compareString(a,b) {
  if (a[1]==b[1]){return 0}
  if (a[1]>b[1]){return 1}
  return -1;
}

function getDtShNm(dt) { 
  let dstr; let d;
  if (typeof dt === "string" ){
    d=getDtObj(dt);
    dstr=dt.replace(/0(\d)/g, "$1");
  } else {
    d=dt;
    dstr=getFmtDtStr(dt);
  }
  let shnm=(d.getDay()+1) + ' ' + dstr;
  return shnm;
}

function convertTZ(date, tzString) {
    return new Date((typeof date === "string" ? new Date(date) : date).toLocaleString("en-US", {timeZone: tzString}));   
}
/*
function zminutTrigger(dows) { 
  //return;
  collectParams();
  writeLog('Start');
  let today1= new Date();
  let today=convertTZ(today1,gp.dates_tz);
  gp.shib_dates=getFmtDtStr(today);
  let todaysh=getShibutzSS().getSheetByName(getDtShNm(gp.shib_dates));
  Logger.log('sh nm='+todaysh.getName()+' today='+today+' today1 day='+today1.getDay());
  // redo today if dows.include(today's dow) and sheet exists
  Logger.log('dows='+dows+' day='+today.getDay());
  if (dows.includes(today.getDay().toString()) && todaysh ) {
    shibutzDates(getShibutzSS());
  }
  checkLog('mail', 'schedule issue',gp.shibutz_mail_to);
}*/

function updateShibSheets2() { 
  collectParams();
  writeLog('Start');
  updateShibSheetsWork(9,10);
  checkLog('mail', 'schedule issue',gp.shibutz_mail_to);
}

function updateShibSheets() { 
  collectParams();
  writeLog('Start');
  updateShibSheetsWork(9,1);
  checkLog('mail', 'schedule issue',gp.shibutz_mail_to);
}

function updateShibSheetsWork(num,start_from) { 
  let dts=getDtsOfSheetsToWorkOn(num,start_from);
  if (dts.length){
    gp.shib_dates=dts.join(',');
    shibutzDates(getShibutzSS());
  }  
}

function getDtsOfSheetsToWorkOn(limit, start_from, ret_sh_nm) { 
  let shs=getShibutzSS().getSheets();
  let dts=[];
  let counter=0;
  for (let i=0;i<shs.length;i++){
    //Logger.log('i='+i+' sh nm='+shs[i].getName());
    let res=isSheet2upadte(shs[i]);
    //Logger.log('i='+i+' res='+res+' nm='+shs[i].getName());
    if (res){
      counter++;
      if (counter < start_from) {continue;}
      if (ret_sh_nm){
        dts.push(res[1]);
      }else{
        dts.push(res[0]);
      }
    }
    //Logger.log(' trig. push dt='+res[0] + ' dow='+res[1]);
    if (limit && dts.length == limit){
      break;
    }
  }
  return dts;
}

function isSheet2upadte(sh) { 
  if (sh.getTabColor() == gv_wrkr_view_tab_color){
    return;
  }
  let nm=sh.getName();
  if (! nm.match(/^\d \d+\/\d+\/\d+$/)){//tabname
    return;
  }
  let dt=nm.substring(2);////tabname
  //Logger.log(' found '+dt+' '+dowrmap[dowl]);
  return [dt, nm];
  //return [dt, dowrmap[dowl],nm];
}

function getShibTmpl(dt) {
  let tmpl_nm='template';
  if (! gp.shibutz_tmplts){
    return tmpl_nm;
  }
  if (! gp.shibutz_tmplts_hash){
    gp.shibutz_tmplts_hash= JSON.parse('{'+gp.shibutz_tmplts+'}');
  }
  let dow2=getDtObj(dt).getDay()+1;
  if (gp.shibutz_tmplts_hash[dow2]){
    tmpl_nm=gp.shibutz_tmplts_hash[dow2];
  }
  if (gp.shibutz_tmplts_hash[dt]){
    tmpl_nm=gp.shibutz_tmplts_hash[dt];
  }
  return tmpl_nm;
}

function createDtSh(ss, dt) { 
  var sh=ss.getSheetByName(dt);
  if (sh){
    ss.deleteSheet(sh);
  }
  var tmpl=getShibTmpl(dt);
  sh=ss.getSheetByName(tmpl).copyTo(ss);
  sh.setName(dt);
  return sh;
}

function getRecurSh() {
  if (! gp.shibutz_recur_sh){
    gp.shibutz_recur_sh=getShibutzSS().getSheetByName('recur');
  }
  return gp.shibutz_recur_sh;
}

function getRecurRowDat(i) {
  if (!gp.shib_recur_dat_rng_ar){
    gp.shib_recur_dat_rng_ar=[];
  }
  if (! gp.shib_recur_dat_rng_ar[i]){
    gp.shib_recur_dat_rng_ar[i]=getRecurSh().getRange(i+2,5,1,12);
  }
  return gp.shib_recur_dat_rng_ar[i];
}

function getShibRecurAr() {
  if (! gp.shib_recur_ar){
    let rows=getRecurSh().getLastRow();
    if (rows>1){
      gp.shib_recur_ar=getRecurSh().getRange(2,1,rows-1,18).getValues();
      gp.shib_recur_dat_rng_ar=[];
      Logger.log('loaded recur');
    } else {
      gp.shib_recur_ar=[];
    }
  }
  return gp.shib_recur_ar;
}

function addRecurring(sh,dt,ss) {
  let dow=dt.getDay();
  //Logger.log('addRecuring dt='+dt);
  let leng=getShibRecurAr().length;
  for (let i=0;i<leng;i++){
    let recur_row=getShibRecurAr()[i];
    //Logger.log('i='+i+ 'shib_recur_ar[i]='+recur_row);
    let recdow=recur_row[0];
    //Logger.log('dow='+ dow+ ' recdow='+recdow+ ' dowmap[dow]='+dowmap[dow]);
    if ((recdow == dowmap[dow]) && isDtInRng(dt,recur_row[16],recur_row[17])){
      //Logger.log('recdow == dowmap[dow]');
      let rka = recur_row.slice(1,4);
      let rd=getRecurRowDat(i);
      setRecurMeet(sh,rka,rd,i+2);
    }
  }
}

function deleteObsoleteRecurRows() {
  let yest=new Date();
  yest.setDate(yest.getDate() - 1);
  let recur_rows2delete=[];
  let leng=getShibRecurAr().length;
  for (let i=0;i<leng;i++){
    let recur_row=getShibRecurAr()[i];
    Logger.log('recdow == dowmap[dow]');
    if (recur_row[17] && yest.getTime()>recur_row[17].getTime()){
      recur_rows2delete.push(i+2);
    }
  }
  Logger.log('recur_rows2delete='+recur_rows2delete);
  for (let i=recur_rows2delete.length -1; i>=0; i--){
    Logger.log('i='+i+ ' recur_rows2delete[i]='+(recur_rows2delete[i]));
    getRecurSh().deleteRow(recur_rows2delete[i]);
  }
}

function isDtInRng(dt,rngb,rnge) {
  //Logger.log('dt='+dt+ ' rngb='+rngb+' rnge='+rnge);
  if (rngb && dt.getTime()<rngb.getTime()){
    return 0;
  }
  if (rnge && dt.getTime()>rnge.getTime()){// period over. remember to delete row
    return 0;
  }
  return 1;
}

function setRecurMeet(sh,rec,rd,rrow) {
  let i=findRecRow(sh,rec);
  if (i){
    //Logger.log('set Recur Meet i='+i+' setRecurMeet rec:' + rec );
    rd.copyTo(sh.getRange(i,5,1,12));
  } else {
    writeLog('cannot set recur. sheet='+sh.getName()+' worker not available. recur row='+rrow+' key='+rec+' pupil='+rd.getValues());
  }
}

function findRecRow(sh,rec) {
  //Logger.log('shnm='+sh.getName());
  if (! gp.shib_sheets_keys){
    gp.shib_sheets_keys={};
  }
  if (! gp.shib_sheets_keys[sh.getName()]){
    let lr=sh.getLastRow();
    gp.shib_sheets_keys[sh.getName()] = sh.getRange(2,1,lr-1,3).getValues();
    //Logger.log('findrecrow loaded rows '+ gp.shib_sheets_keys[sh.getName()]);
  }
  let r=gp.shib_sheets_keys[sh.getName()];
  //Logger.log(' rec[0]='+rec[0]+' rec[1]='+rec[1]+' rec[2]='+rec[2]);
  for (var i=0;i<r.length;i++){
    let s=r[i];
    //Logger.log('i='+i+' s='+s+' rec='+rec);
    //Logger.log(' s[0]='+s[0]+' s[1]='+s[1]+' s[2]='+s[2]);
    if (s[2] == rec[2] && s[0] == rec[0] && s[1] == rec[1]){
      //Logger.log('found row'+ i +' for rec:'+rec );
      return i+2;
    }
  }
  Logger.log('didnt find row for rec:' + rec);
}

function getDtObj(datestr) { 
  let str=datestr.split('/');
  //Logger.log('gtDtObj str='+str+' locale='+gp.locale);
  if (!str || str.length<3) {
    Logger.log('invalid date string datestr='+datestr);
    return
  }
  let d;
  if (gp.locale=='he-IL'){
    d= new Date(str[1]+'/'+str[0]+'/'+str[2]);
    //Logger.log('gtDtObj  d='+d);
  } else {
    d= new Date(datestr);
  }
  //Logger.log('gtDtObj datestr='+datestr+' d='+d);
  return d;
}

function getDtObjFromTabNm(nm) {
  let dtstr=nm.match(/\d+\/\d+\/\d+/, '$1');
  //Logger.log('getDtObjFromTabNm nm='+nm+' dtstr='+dtstr);
  return dtstr ? getDtObj(dtstr[0]) : null;
}

function fillShibDt(date,dt_sh, shib_ss) { 
  if (! dt_sh){
    writeLog('missing sheet:'+date);
    return;
  }
  let dt=getDtObj(date);
  let dow=dt.getDay();
  Logger.log('fillShibDt dow='+dow+' date='+date);
  var shnm=getDtShNm(date);// +dowmap[dow]; //dowmap2  fixfmt
  Logger.log('s fillTmRngs shnm='+shnm);
  fillTmRngs(dt_sh,dow);
  Logger.log('s sheetProtection 4');
  sheetProtection(dt_sh,4);
  Logger.log('s addRecurring');
  addRecurring(dt_sh,dt,shib_ss);
  var old_sh=shib_ss.getSheetByName(shnm);
  if (old_sh){
    Logger.log('s sheetProtection 2');
    sheetProtection(old_sh,2);
    Logger.log('s cpMeetings');
    cpMeetings(old_sh,dt_sh);
  }
  Logger.log('s takeOverDt');
  takeOverDt(shib_ss, dt_sh, shnm, old_sh);
  //Logger.log('s crtWView');
  //crtWView(shib_ss,shnm,dt_sh.getLastRow());
  //Logger.log('e crtWView');
  //unlockDow(dow);
  Logger.log('fillShibDt end  dow='+dow+' date='+date);
}

function takeOverDt(ss,sh,dtndow,old_sh) { 
  if (old_sh){
      var df2=' pre:2';
      var pre2=ss.getSheetByName(dtndow+df2);
      if (pre2){
        ss.deleteSheet(pre2);
      }
      var df1=' pre:1';
      var pre1=ss.getSheetByName(dtndow+df1);
      if (pre1){
        pre1.setName(dtndow+ df2);
      }
      //var oldnm=dtndow+ ' '+df;
      old_sh.setName(dtndow+df1);
      Logger.log('renamed '+ dtndow + ' to '+ dtndow+df1)
      old_sh.hideSheet();
  }
  sh.setName(dtndow);
  sh.showSheet();
}

function crtWView(ss,shnm,rows){
  var wshnm=shnm+' T';
  var sh=ss.getSheetByName(wshnm);
  if (sh){
    ss.deleteSheet(sh);
  }
  sh=ss.getSheetByName('wtmpl').copyTo(ss);
  sh.setName(wshnm);
  Logger.log('s sheetProtection');
  sheetProtection(sh,2);
  Logger.log('s fillRefs');
  fillRefs(sh,shnm,rows);
  Logger.log('s orderCols');
  orderCols(sh,rows);
  Logger.log('s clearDupRows');
  clearDupRows(sh);
  sh.setFrozenColumns(2);
  sh.showSheet();
}

function fillRefs(sh,shnm,rows){
  var cmap=['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P']
  var ar=[];
  var cols=cmap.length;
  for (i=1;i<=rows;i++){
    var ro=[];
    for (j=0;j<cols;j++){
      var f="='"+shnm+"'!"+cmap[j]+i;
      ro.push(f);
    }
    ar.push(ro);
  }
  r=sh.getRange(1,1,rows,cols);
  r.setFormulas(ar);
}

function orderCols(sh){
  var range = sh.getRange("D1");
  sh.hideColumn(range);
  Logger.log('moving col ');
  sh.moveColumns(sh.getRange("C1:C1"), 1);
  sh.getRange(2,1,sh.getLastRow(),1).setTextDirection(SpreadsheetApp.TextDirection.RIGHT_TO_LEFT);
  sortWrkr(sh);
  sh.setTabColor(gv_wrkr_view_tab_color);
}

function sortWrkr(sh){
  //Logger.log('sort ');
  var rows=sh.getLastRow();
  range = sh.getRange("A2:P"+rows);
  range.sort([{column: 1, ascending: true}, {column: 2, ascending: true}]);
  //Logger.log('sorted ');
}

function clearDupRows(sh){
  var last_row=sh.getLastRow();
  var ra=sh.getRange(2,1,last_row-1,1).setFontColor('#999999').getValues();
  //var ra=sh.getRange(2,1,last_row-1,1).setFontColor('#d9d9d9').getValues();
  var prev='';
  //return;
  //Logger.log('s clearDupRows');
  for (var i=0;i<ra.length;i++){
       var cur=ra[i][0];
       if ( cur != prev){ 
         //Logger.log('i='+i+'  cur='+cur);
         sh.getRange(i+2,1).setFontColor('#000000');
         prev=cur;
       }
    }
}

function cpMeetings(old_sh,sh){
  let oldrows=old_sh.getLastRow();
  let oldall=old_sh.getRange(2,1,oldrows-1,16).getValues();
  for (i=0;i<oldall.length;i++){
    let ar=oldall[i].slice(5,11);
    if (ar.every(element => element === "") || oldall[i][15]){//skip empty or recur row. 
      continue;
    }
    let datr=old_sh.getRange(i+2,5,1,12);
    let keyr=old_sh.getRange(i+2,1,1,3);
    let key=keyr.getValues()[0];
    //Logger.log('i='+i+' key='+key);
    let res=addMeeting2newsh(key,datr,sh);
    //Logger.log(' res='+res);
  }
}

function addMeeting2newsh(key,stur,sh){
  var i=findRecRow(sh,key)
  if (i){
    var nstur=sh.getRange(i,5,1,12);
    //if (nstur.isBlank()) {
      stur.copyTo(nstur);
    //} else {
     // writeLog('meeting row occupied. key:'+key + ' with:'+nstur.getValues());
    //}
  } else {
    writeLog('missing row for meeting. sheet='+sh.getName()+' key:'+key +' stu='+stur.getValues());
  }
}

function fillTmRngs(dt_sh,dow){
  var prev_wrkrs_ar=[]; var prev_fr;var prev_to;
  for (var r=dt_sh.getLastRow(); r>1; r--){ //rnge
    var frto= dt_sh.getRange(r,1,1,2).getValues();
    var frtm= frto[0][0];
    var totm= frto[0][1];
    var tmrng=getTmRngVars(frtm,totm);
    //Logger.log('r='+r+' frtm='+frtm+' totm='+totm+' tmrng='+JSON.stringify(tmrng));
    //Logger.log('s getAvailWrkrs');
    var wrkrs=getAvailWrkrs(dow,tmrng);
    //Logger.log('e getAvailWrkrs');
    //Logger.log('frtm='+frtm+' totm='+totm+' avail='+wrkrs);
    if (! wrkrs || wrkrs.length ==0) {
      //writeLog("no workers for date="+dowmap[dow]+' rng='+frtm+'-'+totm);
      dt_sh.deleteRow(r);
      continue;
    }
    //Logger.log('s fillNames');
    fillNames(dt_sh,r,wrkrs,frtm,totm,prev_fr,prev_wrkrs_ar);
    //Logger.log('e fillNames');
    prev_wrkrs_ar = wrkrs;
    prev_to=totm;
    prev_fr=frtm;
  }
}

function getArrWSubj(e){
  var wo=getWorkerByName(e);
  if (! wo) {
    writeLog("no worker by this name="+e);
  }
  return [e, wo.subj]
}

function dupArrRow(a,n){
  var ra=[];
  for (var i=0;i<n;i++){
    ra.push(a);
  }
  return ra;
}

function fillNames(sh,r,wrkrs,frtm,totm,prev_fr,prev_wrkrs_ar){
  //Logger.log(' r='+r+' wrkrs.length='+wrkrs.length+' prev_wrkrs_ar='+prev_wrkrs_ar);
  if (wrkrs.length > 1){
    sh.insertRowsAfter(r, wrkrs.length-1);
    var frtoar=dupArrRow([frtm,totm],wrkrs.length-1);
    sh.getRange(r+1,1,wrkrs.length-1,2).setFontColor('#d9d9d9').setValues(frtoar);
  }
  sh.getRange(r,3,wrkrs.length,2).setValues(wrkrs.map(getArrWSubj));
  
  //for (var i=0;i<wrkrs.length;i+=1){
    //Logger.log('prev_fr='+prev_fr);
  //  if ( totm>prev_fr && prev_wrkrs_ar.includes(wrkrs[i]) ) {
  //    sh.getRange(r+i+1,2).setBackground('#e6a5ca');
  //  }
  //}
}

function getTmRngVars(fr_tm,to_tm){
  var tmrng={};
  tmrng.fr_tm=fr_tm;
  tmrng.to_tm=to_tm;

  tmrng.adt=new Date();
  tmrng.adt.setSeconds(0);
  tmrng.adt.setMilliseconds(0);
   
  tmrng.frtm=new Date(tmrng.adt.getTime());
  var frtm_a=fr_tm.split(':');
  tmrng.frtm.setHours(frtm_a[0]);
  tmrng.frtm.setMinutes(frtm_a[1]);
  
  tmrng.totm=new Date(tmrng.adt.getTime());
  var totm_a=to_tm.split(':');
  tmrng.totm.setHours(totm_a[0]);
  tmrng.totm.setMinutes(totm_a[1]);

  tmrng.frtmgt=tmrng.frtm.getTime();
  tmrng.totmgt=tmrng.totm.getTime();
  return tmrng;
}

function getAvailWrkrs(dow,tmrng) { 
  var col=dow+3;
  var last_row;
  var avail=[];
  loadZminutSh();
  for (var i=0;i<gp.zmin_nms.length;i++){ //wrkrs
    var wnm=gp.zmin_nms[i][0];
    wnm=chomp(wnm.toString());
    if (! wnm){
      continue;
    }
    if (! getWorkerByName(wnm)){
      writeLog('worker not found:'+wnm);
      continue;
    }
    if (testNonRoundHour4Wrkr(wnm,tmrng,dow)){
      continue;
    }
    //Logger.log(' nm='+wnm); 
    var rng=gp.zmin_rngs[i][dow];
    //Logger.log('rng='+rng+' i='+i+' dow='+dow+' nm='+wnm); //mmm
    var zamin=isAvail(rng,tmrng,wnm);
    if (zamin){
      avail.push(wnm);
      //Logger.log('pushing '+wnm+ ' avail='+avail); //mmm
    }
  }
  //Logger.log('ret avail '+avail);
  return avail.sort();
}

function testNonRoundHour4Wrkr(wnm,tmrng,dow) {// return true if non round and wrkr doesnt have recur with this tmrng
  if ( tmrng.fr_tm.match(/:00$/) &&  tmrng.to_tm.match(/:00$/)){
    return;
  }
  let x=getShibRecurAr().find(e => e[0]==dowmap[dow] && e[1]==tmrng.fr_tm && e[2]==tmrng.to_tm && e[3]==wnm);
  //Logger.log('tmrng='+tmrng+' dow='+dow+' nm='+wnm+' x='+x);
  if (x == undefined){
    //Logger.log('tmrng='+tmrng+' dow='+dow+' nm='+wnm);
    return 1;
  }
}

function loadZminutSh() {
  if (! gp.zminut_sh){
    gp.zminut_sh=SpreadsheetApp.openById(gp.zminut_file_id).getSheetByName('Form responses 1');
    last_row=gp.zminut_sh.getLastRow();
    gp.zmin_nms=gp.zminut_sh.getRange(2,2,last_row-1,1).getValues();
    gp.zmin_rngs=gp.zminut_sh.getRange(2,3,last_row-1,6).getValues();
    //Logger.log(' nms='+gp.zmin_nms);
    Logger.log(' getAvailWrkrs init');
  }
}

function isAvail(rngs,tmrng,wnm) { 
  if (! rngs ){
    return;
  }
  if (typeof  rngs != "string"){
    writeLog('range is not a string. range='+rngs);
    return;
  }    
  //Logger.log('rngs='+rngs);
  var rngar=rngs.split(',');
  for (var i=0;i<rngar.length;i++){
    var rngi =rngar[i];
    if (! rngi || rngi.length == 1){
      continue;
    }
    rng=fixTimeRange(rngi);
    //Logger.log('rngi='+rngi+' fix='+rng);
    var rng2=rng.split('-');
    var rng2fr_a=rng2[0].split(':');
    var rtmf=new Date(tmrng.adt.getTime());
    rtmf.setHours(rng2fr_a[0]);
    rtmf.setMinutes(rng2fr_a[1]);
    //Logger.log('rtmf='+rtmf);
    if (! rng2[1]){
      writeLog('invalid range in zminut file (missing "-" ?). worker=' + wnm + ' rngs='+rngs+' rng=' + rng +' rngi=' + rngi  );
      continue;
    }
    var rng2to_a=rng2[1].split(':');
    var rtmt=new Date(tmrng.adt.getTime());
    //Logger.log(' rtmf='+rtmf+' rtmt='+rtmt);
    rtmt.setHours(rng2to_a[0]);
    rtmt.setMinutes(rng2to_a[1]);
    //Logger.log('3rng2fr_a='+rng2fr_a+' rng2to_a='+rng2to_a+' rtmf='+rtmf+' rtmt='+rtmt);
    var rtmfgt=rtmf.getTime();
    var rtmtgt=rtmt.getTime();
    //Logger.log(' rtmf='+rtmf+' rtmt='+rtmt);
    //Logger.log('rng='+rng+' rtmfgt='+rtmfgt+' rtmtgt='+rtmtgt+'  tmrng.frtmgt='+tmrng.frtmgt+' tmrng.totmgt='+tmrng.totmgt);
    if (rtmfgt<=tmrng.frtmgt && rtmtgt>=tmrng.totmgt){
      //Logger.log('match');
      return 1;
    }
  }
}

function fixTimeRange(rng) {
  var orig=rng;
  rng=chomp(rng);
  var regexp; var match;
  rng=rng.replace(/;/g, ':').replace(/,/, '');  
  regexp =  /^(0+)$/m; // 0
  match = regexp.exec(rng);
  if (match){
  rng=rng.replace(regexp,'' );
  }
  regexp =  /^([^-]+$)/m; // hhmm
  match = regexp.exec(rng);
  if (match){
   rng=rng.replace(regexp,match[1]+'-23:00' );
  }
  regexp =  /(-\d\d|-\d)$/m; // hh-
  match = regexp.exec(rng);
  if (match){
  rng=rng.replace(regexp,match[1]+':00' );
  }
  regexp =  /(\d\d)(\d\d)$/m; // -hhmm
  match = regexp.exec(rng);
  if (match){
    rng=rng.replace(regexp,match[1]+':'+match[2] );
  }
  regexp =  /^(\d\d)(\d\d)-/m; // hhmm-
  match = regexp.exec(rng);
  if (match){
   rng=rng.replace(regexp,match[1]+':'+match[2]+'-' );
  }
  regexp =  /^(\d+)-/m; // hh-
  match = regexp.exec(rng);
  if (match){
   rng=rng.replace(regexp,match[1]+':00-' );
  }
  //Logger.log('orig rng='+orig+' rng='+rng);
  return rng;
} 

function shibutsDateProgressMain() {
  collectParams(3);
  writeLog('Start');
  let dts=getPrevNxtDates();
  //return;
  if (dts[2] != 6){
    gp.shib_dates=dts[0];
    clearShibutzDates();
  }
  if (dts[3] != 6  && ! getShibutzSS().getSheetByName(getDtShNm(dts[1]))){
    gp.shib_dates=dts[1];
    let shib_ss=getShibutzSS();
    shibutzDates(shib_ss);
  } 
  checkLog('mail','shibutz date progress');
}

function getPrevNxtDates(current_dt) {
  let cur_dt=  current_dt ? current_dt : new Date();
  //cur_dt.setHours(0,0,0,0);
  let prev_dt=cur_dt;
  prev_dt.setDate(prev_dt.getDate() - 1);
  Logger.log('prev_dt='+prev_dt);
  let prev_dow=prev_dt.getDay();
  //Logger.log('prev_dt2='+prev_dt);
  let prev_dt_formated=getFmtDtStr(prev_dt);
  prev_dt_formated = prev_dt_formated.replace(/0(\d)/g, "$1");
  //Logger.log('prev_dt_formated='+prev_dt_formated);
  let nxt_dt=new Date();
  nxt_dt.setDate(nxt_dt.getDate() + (gp.shib_days_cycle - 1));
  let nxt_dow=nxt_dt.getDay();
  Logger.log('nxt_dt='+nxt_dt);
  let nxt_dt_formated=getFmtDtStr(nxt_dt);
  nxt_dt_formated=nxt_dt_formated.replace(/0(\d)/g, "$1");
  Logger.log('nxt_dt_formated='+nxt_dt_formated);
  Logger.log('prev_dow='+prev_dow+' nxt_dow='+nxt_dow);
  return [prev_dt_formated,nxt_dt_formated,prev_dow,nxt_dow]
}


function expandGroup2members(ar,replace,from,to) {//from: 0 based, to: 0 based, exclusive 
  let gs=getGroupsDict();
  //Logger.log(' ar.length='+ar.length);
  for (let i=0;i<ar.length;i++){
    //Logger.log('i='+i+' ar[i].length='+ar[i].length);
    let merged='';
    for (let c=from;c<to;c++){
      //Logger.log('i='+i+' c='+c);
      //Logger.log('ar[i][c]]='+ar[i][c]);
      if (gs[ar[i][c]]){
        ar[i][c] = replace ? gs[ar[i][c]] : (ar[i][c] + ': ' + gs[ar[i][c]]) ;
      }
      if (ar[i][c]){
        merged= merged ? (merged+', ') : merged;
        merged += ar[i][c];
      }
    }
    ar[i][from] = merged; 
  }
}

function getSchedWrkrRows(nm, hist) {
    Logger.log('nm='+nm+' hist='+hist );
  let query = 'select A, B, C, D, F, G, H, I, J, K, L, M where (F != "" or G !="" or H !="" or I !="" or J !="" or K !="" or L !="" or M !="" or N !="")';
  let shnm='allDays';
  if (hist=='y'){
    query = 'select A, B, C, D, F, G, H, I, J, K, L, M ';
    query += nm ? (' where D = "'+nm+'"') : "";
    shnm='history';
  } else {
    query += nm ? (' and D = "'+nm+'"') : "";
  }
  let values = querySheet(query, gp.shibutz_file_id, shnm);
  //  Logger.log('valx='+JSON.stringify(values) );
  expandGroup2members(values,0,8,12);
  //  Logger.log('valx2='+JSON.stringify(values) );
  values.forEach(e => e.splice(9,99));
  //Logger.log('valx3='+JSON.stringify(values) );
  if (hist=='y'){
    values.splice(0,values.length-1000);
    values.reverse();
  }
  //Logger.log('7='+values[1][7]);
  Logger.log('getSchedWrkrRows values='+values);
  return values;
}

function crtSchedTabl(rows) {
  let ar=[];
  hrs=['3-4','4-5','5-6','6-7','7-8'];
  for (let i=0; i<hrs.length;i++){
    ar[i]=[];
    ar[i]=[hrs[i],'','','','',''];
  }
  let hmap={'15:00':0, '16:00':1, '17:00':2, '18:00':3, '19:00':4 };
  //Logger.log('ar='+JSON.stringify(ar));
  for (let i=0; i<rows.length;i++){
    let day=rows[i][0].substring(0,1);
    let val=rows[i][3]+'/'+rows[i][4];
    val = rows[i][5] ? (val+'/'+rows[i][5]) : val;
    let r=hmap[rows[i][1]];
    //Logger.log('r='+r+' val='+val+' i='+i+' day='+day+' rows[i]='+rows[i]);
    if (!(rows[i][2] in hmap) || !(rows[i][1] in hmap) || ! ['1','2','3','4','5'].includes(day)) {
      Logger.log('time/day not in table. row='+rows[i]);
      if (! ar[5]){
        ar[5]=['','','','','',''];
      }
      r=5;
      val=rows[i][1]+'-'+rows[i][2]+' '+val;
    }
    //ar[r][day]=val;
    //Logger.log('xval='+val+' r='+r+' day='+day+' ar[r][day]='+ar[r][day]);
    if (ar[r][day] != val){
      ar[r][day]= ar[r][day] ? (ar[r][day] +'<br>'+val) : val;
    }
    //Logger.log('val='+val+' r='+r+' day='+day+' ar[r][day]='+ar[r][day]);
  }
  Logger.log('ar2='+JSON.stringify(ar));
  return ar;
}

function getPupRows(nm,targetSheet) {
  let query = 'select A where B = "'+nm+'"';
  //let gs_a=querySheet(query,gp.pupil_alfon_id,'groups').shift();
  let gs_a=querySheet(query,gp.pupil_alfon_id,'groupPupil');
  Logger.log('nm='+nm+' gs_a= '+JSON.stringify(gs_a));
  if (gs_a && gs_a.length){
    let a=gs_a.map(e => e[0]);
    a.push(nm);
    //Logger.log('a= '+a);
    nm=a.join('|');
  }
  Logger.log('nm= '+nm);
  //query = 'select A, B, E, C, F where (I matches "'+nm+'" or J matches "'+nm+'" or K matches "'+nm+'" or L matches "'+nm+'" or M matches "'+nm+'" or N matches "'+nm+'" or O matches "'+nm+'") ';  
  //xquery = 'select A, B, E, C, F where (I matches "'+nm+'" or J matches "'+nm+'" or K matches "'+nm+'" or L matches "'+nm+'")';  
  //query = 'select A, B, C, F, D, G where (J matches "'+nm+'" or K matches "'+nm+'" or L matches "'+nm+'" or M matches "'+nm+'")';  
  //query = 'select A, B, C, F, D, G where (R matches ".*,('+nm+'),.*")';  
  query = 'select A, B, C, F, D, G where (T matches ".*,('+nm+'),.*")';  
  Logger.log('query= '+query);
  let values=[];
  for (let i=0;i<targetSheet.length;i++){
    //Logger.log('targetSheet[i]= '+targetSheet[i]);
    let val=querySheet(query, gp.shibutz_file_id, targetSheet[i],0);
    Logger.log('val= '+val);
    //xval.forEach(e => e.splice(0,0,targetSheet[i]));
    values=values.concat(val)
  }
  Logger.log('getPupRs values='+JSON.stringify(values));
  return values;
}

function getPupilSched(pnm,mode) {
  Logger.log('pnm='+pnm+'mode='+mode);
  out={};
  //out.email = Session.getActiveUser().getEmail();
  out.name=pnm;
  out.found='y';
  let pm=getAllPupilsMap();
  Logger.log('got bymail');
  if (pm[pnm]){
    Logger.log(' nm='+out.name);
    //xlet dts=getDtsOfSheetsToWorkOn(7,1,1);
    let dts=['allDays'];
    Logger.log('dts='+dts);
    out.rows=getPupRows(out.name,dts);
    if (mode=='week'){
      out.rows=crtSchedTabl(out.rows);
    }
  } else {
    out.found='n';
    Logger.log('mail='+out.email+' Invalid');
  }
  return (out);
}

function updateAllDatesSheetWork() {
  let dts=getDtsOfSheetsToWorkOn(10,1,1);
  let ar=[];
  collectDatesMetaAr(dts,ar);
  //Logger.log('ar='+JSON.stringify(ar));
  updateAllDatesSheet(ar);
}

function collectDatesMetaAr(dts, ar) {
  let row2set=2;
  for (let i=0; i<dts.length;i++){
    let sh=getShibutzSS().getSheetByName(dts[i]);
    let rows=sh.getLastRow()-1;//=ARRAYFORMULA('1 27/3/22'!A2:Q)
    let dt=dts[i].replace(/\d /,'');//tabname
    let f1='=ARRAYFORMULA({"'+dts[i]+'"&Y1:Y'+rows+"})";  //=ArrayFormula({"27/3/22"&Y2:Y500})
    let f2="=ARRAYFORMULA('"+dts[i]+"'!A2:R"+(rows+1)+")";
    //Logger.log('collect rows='+rows+' row2set='+row2set+' dt='+dt+' shnm='+sh.getName());
    ar.push([row2set,rows,f1,f2]);
    row2set+=rows;
  }
}

function updateAllDatesSheet(ar) {
  let sh=getShibutzSS().getSheetByName('allDays');
  sh.getRange(2,1,sh.getLastRow(),2).clear({ contentsOnly: true });
  //Logger.log('upd ar='+JSON.stringify(ar));
  let lr=ar[ar.length-1][0];
  let newlr=lr + ar[ar.length-1][1] -1;
  let rng=sh.getRange(2,1,lr-1,2);
  let rng_ar=rng.getValues();
  Logger.log('upd lr='+lr+' rng_ar.length'+rng_ar.length+' rng_ar[0].length='+rng_ar[0].length);
  for (let i=0; i<ar.length;i++){
    //Logger.log('upd i='+i+' ar[i][0]='+ar[i][0]);
    //Logger.log('upd  rng_ar[ar[i][0] - 2]='+rng_ar[ar[i][0] - 2]);
    rng_ar[ar[i][0] - 2][0]=ar[i][2];
    rng_ar[ar[i][0] - 2][1]=ar[i][3];
  }
  rng.setValues(rng_ar);
  let c='=ARRAYFORMULA("," & J2:J'+newlr+' & "," & K2:K'+newlr+' & "," & L2:L'+newlr+' & "," & M2:M'+newlr+' & "," & N2:N'+newlr+' & "," & O2:O'+newlr+' & "," & P2:P'+newlr+' & ",")';
  //=arrayformula("," & J2:J871 & "," & K2:K871 & "," & L2:L871 & "," & M2:M871 & "," & N2:N871 & "," & O2:O871 & "," & P2:P871 & ",")
  sh.getRange('T2').setValue(c);
}

function sendMeetingReminderMain() {
  let ar=getSchedWrkrRows();
  for (let i=0;i<ar.length;i++){

  }


}