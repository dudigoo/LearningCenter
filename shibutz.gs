var dowmap={'0':'א', '1':'ב', '2':'ג', '3':'ד', '4':'ה','5':'ו','6':'ז'};
var dowrmap={'א':'0', 'ב':'1', 'ג':'2', 'ד':'3', 'ה':'4','ו':'5','ז':'6'};

//var dowmap2={'0':'1', '1':'2', '2':'3', '3':'4', '4':'5','5':'6','6':'7'};
var gv_wrkr_view_tab_color='#1129e9';
//var svrow=1;svcol=14;

function fillShibutzMain() {// 22/12/20,24/12/20
  //Logger.log('locale 3='+gp.locale);
  collectParams();
  if (gp.shib_dates == 'all'){
    updateShibSheetsWork(25,1);
  } else if (gp.shib_dates == 'all2'){
    updateShibSheetsWork(25,26);
  } else {
    shibutzDates();
  }
  checkLog('mail', 'schedule issue',gp.shibutz_mail_to);
}

function getShibutzSS() {
  if (! gp.shibutz_ss){
    gp.shibutz_ss=SpreadsheetApp.openById(gp.shibutz_file_id);
  }
  return gp.shibutz_ss;
}

function clearShibutzDatesMain() {
  collectParams();
  clearShibutzDates()
  checkLog('mail', 'schedule issue');
}

function clearShibutzDates() {
  let shib_ss=getShibutzSS();
  getRecurSh().showSheet();
  let dts=gp.shib_dates.split(",");
  let sha=shib_ss.getSheets();
  for (var i=0;i<sha.length;i++){
    let nm= sha[i].getName();
    let rcd=getArchAction(nm,dts);
    //Logger.log('check rcd='+rcd+' nm='+nm);
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
    let filter = sha[i].getFilter();
    if (filter){
      filter.remove();
    }
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

function shibutzDates() {
  var dts=chomp(gp.shib_dates).split(",");
  Logger.log('dtsb='+gp.shib_dates);
  //return; //mmm
  for (var i=0;i<dts.length;i++){
    Logger.log('dt='+dts[i]);
    fillShibDt(dts[i]);
  }
  Logger.log('S orderShibSheets');
  //orderShibSheets();
  Logger.log('S updateAllDatesSheetWork');
  updateAllDatesSheetWork();
  Logger.log('E updateAllDatesSheetWork');
  deleteObsoleteRecurRows();
}

function getShibWorkSh() { 
  if (! gp.shibutz_work_sh){
    gp.shibutz_work_sh=getShibutzSS().getSheetByName('work');
  }
  return gp.shibutz_work_sh;
}

function orderShibSheets() { 
  var shs=getShibutzSS().getSheets();
  var shs2s=shs.map(mkSortable);
  shs2s.sort(compareString);
  let re =  /^(\D)/m; 
  for (let i=0;i<shs2s.length;i++){
    let sh=getShibutzSS().setActiveSheet(shs2s[i][0]);
    let match = re.exec(shs2s[i][1]);
    if (match && sh.getName() != 'recur'){
      sh.hideSheet();
    }
    getShibutzSS().moveActiveSheet(i+1);
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

function updateShibSheetsMain() { 
  collectParams();
  updateShibSheetsWork(25,1);
  checkLog('mail', 'schedule issue',gp.shibutz_mail_to);
}

function updateShibSheetsWork(num,start_from) { 
  let dts=getDtsOfSheetsToWorkOn(num,start_from);
  Logger.log('dts.length='+dts.length);
  if (dts.length){
    gp.shib_dates=dts.join(',');
    shibutzDates();
  }  
}

function updateShibBySheetName(shnm){
  collectParams();
  Logger.log('updateShibBSheetName shnm='+shnm);
  gp.shib_dates=getDtStrFromShNm(shnm);
  shibutzDates();
  checkLog('mail', 'updateShibBySheetNm schedule issue ');
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

function createDtSh(ss, dt, shnm) { 
  var tmpl=getShibTmpl(dt);
  sh=ss.getSheetByName('format').copyTo(ss);
  sh.setName(shnm);
  return sh;
}

function getRecurSh() {
  if (! gp.shibutz_recur_sh){
    gp.shibutz_recur_sh=getShibutzSS().getSheetByName('recur');
  }
  return gp.shibutz_recur_sh;
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

function addRecurring(wrk_a,dt) {
  let dow=dt.getDay();
  //Logger.log('addRecuring dt='+dt);
  let leng=getShibRecurAr().length;
  for (let i=0;i<leng;i++){
    let recur_row=getShibRecurAr()[i];
    //Logger.log('i='+i+ 'shib_recur_ar[i]='+recur_row);
    //Logger.log('dow='+ dow+ ' recur_row[0]='+recur_row[0]+ ' dowmap[dow]='+dowmap[dow]);
    if ((recur_row[0] == dowmap[dow]) && isDtInRng(dt,recur_row[16],recur_row[17])){
      //Logger.log('recur_row[0] == dowmap[dow]');
      setRecurMeet(wrk_a,recur_row,i+2,dt);
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
    //Logger.log('recdow == dowmap[dow]');
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

function setRecurMeet(wrk_a,recur,rrow, date) {
  let i=findRecRow(wrk_a,recur.slice(1,4));
  if (i>-1){
    //Logger.log('set Recur Meet i='+i+' setRecurMeet rec:' + rec );
    Logger.log('wrk_a[i] len b='+wrk_a[i].length + ' wrk_a[i]='+wrk_a[i]);
    //Logger.log('wrk_a[i] len b='+wrk_a[i].length + ' wrk_a[i]='+wrk_a[i]);
    //wrk_a[i]=wrk_a[i].concat(recur.slice(4,16));
    wrk_a[i]=wrk_a[i].slice(0,4).concat(recur.slice(4,16));
    Logger.log('wrk_a[i] len a='+wrk_a[i].length );
  } else if (isWorkerAvailable(recur[3], date, recur[1], recur[2])){
    Logger.log('not in template but available');
    let w=getWorkerByName(recur[3])
    //let z=recur.slice(1,4).concat([w.subj_popu].concat(recur.slice(4,16)));
    //Logger.log('z='+JSON.stringify(z));
    //wrk_a.push(z);
    wrk_a.push(recur.slice(1,4).concat([w.subj_popu].concat(recur.slice(4,16))));
  } else {
    writeLog('* Cannot set recur. sheet='+date.toDateString()+' worker not available. recur row='+rrow+' recur='+recur);
  }
}

function isWorkerAvailable(wnm, dt, frtm, totm) {
  let tmrng=getTmRngVars(frtm,totm);
  let avail=getAvailWrkrs(dt,tmrng, [wnm]);
  //Logger.log('isWorkerAvailabl avail='+avail);
  return avail.length;
}

function findRecRow(wrk_a,key) {
  //Logger.log('shnm='+sh.getName());
  for (let i=0;i<wrk_a.length;i++){
    let wi=wrk_a[i];
    //Logger.log('i='+i+' wi[0]='+wi[0]+' key[0]='+key[0]);
    //Logger.log('i='+i+' wi[1]='+wi[1]+' key[1]='+key[1]);
    if (wi[0] == key[0] && wi[1] == key[1] && wi[2] == key[2]){
      //Logger.log('found row'+ i +' for rec:'+key );
      return i;
    }
  }
  Logger.log('didnt find row for key:' + key);
  return -1;
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

function getDtStrFromShNm(nm) {
  return nm.replace(/^.\W+/,''); //tabname
}

function getDtObjFromTabNm(nm) {
  let dtstr=nm.match(/\d+\/\d+\/\d+/, '$1');
  //Logger.log('getDtObjFromTabNm nm='+nm+' dtstr='+dtstr);
  return dtstr ? getDtObj(dtstr[0]) : null;
}

function fillShibDt(date) { 
  let dt=getDtObj(date);
  let dow=dt.getDay();
  Logger.log('fillShibDt dow='+dow+' date='+date);
  let shnm=getDtShNm(date);
  let dtsh_exist=1;
  let dt_sh=getShibutzSS().getSheetByName(shnm);
  let wrk_a=[];
  if (!dt_sh){
    dtsh_exist=0;
    dt_sh = createDtSh(getShibutzSS(),date, shnm);
  }
  Logger.log('s fillTmRngs shnm='+shnm);
  fillTmRngs(wrk_a,dt, date);
  Logger.log('s addRecurring');
  addRecurring(wrk_a,dt);
  if (dtsh_exist){
    Logger.log('s cpMeetings');
    cpMeetings(dt_sh,wrk_a,dt);
    Logger.log('e cpMeetings');
  }
  //Logger.log('wrk_a ='+JSON.stringify(wrk_a));
  wrkAr2DtSh(wrk_a, dt_sh, dtsh_exist);
  dt_sh.showSheet();
  Logger.log('fillShibDt end  dow='+dow+' date='+date);
}

function wrkAr2DtSh(wrk_a,dt_sh, dtsh_exist) {
  //Logger.log('s wrkAr2DtSh  wrk_a.length='+wrk_a.length+' dt_sh.getLastRow()='+dt_sh.getLastRow()+' dt_sh.getMaxRows()='+dt_sh.getMaxRows());
  wrk_a.sort((a, b) => { 
  if (a[0]>b[0]){ return 1;}				
  if (a[0]<b[0]){ return -1;}
  if (a[1]>b[1]){ return 1;}				
  if (a[1]<b[1]){ return -1;} 
  if (a[2]>b[2]){ return 1;}				
  if (a[2]<b[2]){ return -1;}
  return 0;
} );  
  
  let rows2add;
  if (dtsh_exist){
    rows2add=wrk_a.length-dt_sh.getMaxRows()+21;
    if (dt_sh.getLastRow()>1){
      dt_sh.getRange(2,1,dt_sh.getLastRow()-1, dt_sh.getLastColumn()).clearContent();
    }
    //Logger.log('dtsh_existh  rows2add='+rows2add );
  } else {
    rows2add=wrk_a.length;
    //Logger.log('dtsh_existh=0  rows2add='+rows2add);
  }
  if (rows2add>0){
    dt_sh.insertRowsAfter(dt_sh.getMaxRows()-1, rows2add);
  }
  //dt_sh.insertRowsAfter(3, wrk_a.length)
  //Logger.log('wrk_a='+JSON.stringify(wrk_a));
  dt_sh.getRange(2,1,wrk_a.length, wrk_a[0].length).setValues(wrk_a);
  Logger.log('e wrkAr2DtSh');
  dt_sh.getRange(2, 1, wrk_a.length, 2).setFontColor('#999999');
  let prev1;
  let prev2;
  for (let i=0; i<wrk_a.length;i++){
    if (prev1 != wrk_a[i][0] || prev2 != wrk_a[i][1]){
      prev1=wrk_a[i][0];
      prev2=wrk_a[i][1];
      dt_sh.getRange(i+2,1,1,2).setFontColor('TEXT');
    }
  }
}

function getTmplHrs(dt) { 
  var tmpl=getShibTmpl(dt);
  //Logger.log('tmpl='+tmpl+' dt='+dt);
  let sh=getShibutzSS().getSheetByName(tmpl);
  let hrs=sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
  return hrs;
}

function cpMeetings(old_sh,wrk_a,dt){
  let oldrows=old_sh.getLastRow();
  if (oldrows<2){return}
  let oldall=old_sh.getRange(2,1,oldrows-1,16).getValues();
  for (i=0;i<oldall.length;i++){
    let ar=oldall[i].slice(5,11);
    if (ar.every(element => element === "") || oldall[i][15]){//skip empty or recur row. 
      continue;
    }
    //Logger.log('i='+i);
    let res=addExistingMeeting(oldall[i],wrk_a, old_sh,dt);
    //Logger.log(' res='+res);
  }
}

function addExistingMeeting(meeting, wrk_a, sh,date){
  var i=findRecRow(wrk_a, meeting.slice(0,3))
  if (i>-1){
    wrk_a[i]=meeting;
  } else if (isWorkerAvailable(meeting[2], date, meeting[0], meeting[1])){
    let w=getWorkerByName(meeting[2]);
    wrk_a.push(meeting.slice(0,16));
  } else {
    writeLog('* Missing row for meeting. sheet='+sh.getName()+' meeting:'+meeting);
  }
}
/*
  } else if (isWorkerAvailable(recur[3], date, recur[1], recur[2])){
    //Logger.log('not in template but available');
    let w=getWorkerByName(recur[3])
    //let z=recur.slice(1,4).concat([w.subj_popu].concat(recur.slice(4,16)));
    //Logger.log('z='+JSON.stringify(z));
    //wrk_a.push(z);
    wrk_a.push(recur.slice(1,4).concat([w.subj_popu].concat(recur.slice(4,16))));

  */
function fillTmRngs(wrk_a,dt, dt_str){
  let hrs= getTmplHrs(dt_str);
  for (let r=0;r<hrs.length;r++){
    var frtm= hrs[r][0];
    var totm= hrs[r][1];
    var tmrng=getTmRngVars(frtm,totm);
    //Logger.log('s getAvailWrkrs');
    //Logger.log('r='+r+' frtm='+frtm+' totm='+totm+' tmrng='+JSON.stringify(tmrng));
    let wrkrs=getAvailWrkrs(dt,tmrng);
    //Logger.log('e getAvailWrkrs');
    //Logger.log('frtm='+frtm+' totm='+totm+' avail='+wrkrs);
    if (! wrkrs || wrkrs.length ==0) {
      //writeLog("no workers for date="+dowmap[dow]+' rng='+frtm+'-'+totm);
      continue;
    }
    //Logger.log('s fillNames');
    fillNames(wrk_a,r,wrkrs,frtm,totm);
    //Logger.log('e fillNames');
  }
}

function getArrWSubj(e){
  var wo=getWorkerByName(e);
  if (! wo) {
    writeLog("no worker by this name="+e);
  }
  return [e, wo.subj_popu]
}

function fillNames(wrk_a,r,wrkrs,frtm,totm){
  //Logger.log(' r='+r+' wrkrs.length='+wrkrs.length);
  wrkrs.forEach( e => {
    let ws=getArrWSubj(e);
    wrk_a.push([frtm,totm,ws[0], ws[1] ,'','','','','','','','','','','',''])
  });
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

function getAvailWrkrs(dt,tmrng,specific_workers) {
  let dow=dt.getDay();
  var col=dow+3;
  var last_row;
  var avail=[];
  loadZminutSh();
  for (var i=0;i<gp.zmin_nms.length;i++){ //wrkrs
    var wnm=gp.zmin_nms[i][0];
    wnm=chomp(wnm.toString());
    //Logger.log('i='+i+' nm='+wnm); //xx
    if (specific_workers && ! specific_workers.includes( wnm)){
      continue;
    }
    if (! wnm){
      continue;
    }
    if (! getWorkerByName(wnm)){
      writeLog('worker not found:'+wnm);
      continue;
    }
    if (isDtInRange(dt,gp.zmin_rngs[i][6],gp.zmin_rngs[i][7],gp.zmin_rngs[i][8])){
      //Logger.log("dt="+dt+" wnm="+wnm +" isDtInRange=1  gp.zmin_rngs[i][8]"+gp.zmin_rngs[i][8] ); //xx
      continue;
    }
    /*if (testNonRoundHour4Wrkr(wnm,tmrng,dow)){
      continue;
    }*/
    var rng=gp.zmin_rngs[i][dow];
    //Logger.log('rng='+rng+' i='+i+' dow='+dow+' nm='+wnm); //mmm
    var zamin=isAvail(rng,tmrng,wnm);
    if (zamin){
      avail.push(wnm);
      //Logger.log('pushing '+wnm+ ' avail='+avail); //mmm
    }
  }
  //Logger.log('ret avail '+avail);
  //return avail;
  return avail.sort();
}

function tstIsDtInRange(){
  let dt= new Date('2022-09-20');
  let gp={};
    gp.zminut_sh=SpreadsheetApp.openById('1SXIOfVl9K4eagWJ8HgcXAR2GAnRIh5IqT_4ptkJAHO4').getSheetByName('Form responses 1');
    let tdts=gp.zminut_sh.getRange(19,9,1,2).getValues();
    let res=isDtInRange(dt,tdts[0][0],tdts[0][1]);
    //Logger.log('res='+res);  

}

function isDtInRange(dt,frdt,todt,dts){
  if (frdt && todt){// important: set correct timezone of calling project 
    //Logger.log('date range dt= '+dt+ ' fr='+frdt+ ' to='+todt);
    if (dt.getTime()>=frdt.getTime() && dt.getTime()<=todt.getTime()) {
      //Logger.log('date in range dt= '+dt+ ' fr='+frdt+ ' to='+todt);
      //Logger.log('dt gettime()= '+dt.getTime()+ ' fr='+frdt.getTime()+ ' to='+todt.getTime());
      return 1;
    }
  }
  if (dts){
    //let dts_a=dts.split(',');
    //for (let i=0;i<dts_a.length;i++){
    for (let i=0;i<dts.length;i++){
      //Logger.log('dts_a[i]= '+dts_a[i]+ ' getDtObj(dts_a[i])='+getDtObj(dts_a[i])+ ' dt='+dt);
      //Logger.log('dts_a[i]= '+dts_a[i]+ ' dt='+dt+ ' dt.getTime()='+dt.getTime());
      //Logger.log('getDtObj(dts_a[i])= '+getDtObj(dts_a[i]));
      //if (getDtObj(dts_a[i]).getTime() == dt.getTime()){
      //  Logger.log('getDtObj(dts_a[i]).getTime() == dt.getTime()');
      if (dts[i] == dt.getTime()){
        return 1;
      }
    }
  }
}

function testNonRoundHour4Wrkr(wnm,tmrng,dow) {// return true if non round and wrkr doesnt have recur with this tmrng
// sigal wanted to ignore template non-round hours if no such recur meeting exists for the teacher
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
    gp.zmin_rngs=gp.zminut_sh.getRange(2,3,last_row-1,9).getValues();
    for (let i=0;i<gp.zmin_rngs.length;i++){
      //Logger.log(' gp.zmin_nms[i][0] ='+gp.zmin_nms[i][0]);
      if (gp.zmin_rngs[i][8]){
        //Logger.log(' gp.zmin_nms[i][0] ='+gp.zmin_nms[i][0]+' gp.zmin_rngs[i][8] ='+gp.zmin_rngs[i][8]);//xx
        let ar=gp.zmin_rngs[i][8].split(',');
        let time_ar=[];
        for (let j=0;j<ar.length;j++){
          //Logger.log(' ar[j] ='+ar[j]);
          if (ar[j].length<6){
            writeLog('invalid unavail date='+ar[j]+ ' full='+gp.zmin_rngs[i][8]);
          } else {
            time_ar.push(getDtObj(ar[j]).getTime());
          }
        }
        gp.zmin_rngs[i][8]=time_ar;
        //Logger.log(' gp.zmin_rngs[i][8] ='+gp.zmin_rngs[i][8]);//xx
      }
    }
    //Logger.log(' nms='+gp.zmin_nms);
    Logger.log('e loadZminutSh init');
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
  let dts=getPrevNxtDates();
  Logger.log('prev nxt='+dts);
  if (dts[2] != 6){
    gp.shib_dates=dts[0];
    Logger.log('rm gp.shib_dates='+gp.shib_dates);
    clearShibutzDates();
  }
  if (dts[3] != 6  && ! getShibutzSS().getSheetByName(getDtShNm(dts[1]))){
    gp.shib_dates=dts[1];
    Logger.log('add gp.shib_dates='+gp.shib_dates);
    shibutzDates();
  } 
  checkLog();
  //checkLog('mail','shibutz date progress');
}

function getPrevNxtDates(current_dt) {
  let cur_dt=  current_dt ? current_dt : new Date();
  //cur_dt.setHours(0,0,0,0);
  let prev_dt=cur_dt;
  prev_dt.setDate(prev_dt.getDate() - 1);
  Logger.log('prev_dt='+prev_dt);
  let prev_dow=prev_dt.getDay();
  Logger.log('prev_dt2='+prev_dt);
  let prev_dt_formated=getFmtDtStr(prev_dt);
  prev_dt_formated = prev_dt_formated.replace(/0(\d)/g, "$1");
  //Logger.log('prev_dt_formated='+prev_dt_formated);
  let nxt_dt=new Date();
  nxt_dt.setDate(nxt_dt.getDate() + (gp.shib_days_cycle - 1));
  let nxt_dow=nxt_dt.getDay();
  //Logger.log('nxt_dt='+nxt_dt);
  let nxt_dt_formated=getFmtDtStr(nxt_dt);
  nxt_dt_formated=nxt_dt_formated.replace(/0(\d)/g, "$1");
  Logger.log('nxt_dt_formated='+nxt_dt_formated);
  Logger.log('prev_dow='+prev_dow+' nxt_dow='+nxt_dow);
  return [prev_dt_formated,nxt_dt_formated,prev_dow,nxt_dow]
}


function expandGroup2members(ar,replace,from,to) {//from: 0 based, to: 0 based, exclusive 
// ar example=[["אב-חוה","חוה-מתמ","","","","",""]]
  let gs=getGroupsDict();
  //Logger.log(' ar.length='+ar.length);
  for (let i=0;i<ar.length;i++){
    //Logger.log('i='+i+' ar[i].length='+ar[i].length);
    let merged='';
    for (let c=from;c<to;c++){
      //Logger.log('i='+i+' c='+c);
      //Logger.log('ar[i][c]]='+ar[i][c]);
      //Logger.log('gs[ar[i][c]]='+gs[ar[i][c]]);
      if (ar[i][c] in gs){
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

function getSchedWrkrRows(nm, hist, group, lvl) {
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
  if (lvl){
    query += ' and H = "'+lvl+'"';
  }
  if (group){
    query += ' and G contains "'+group+'"';
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
  //Logger.log('getSchedWrkrRows values='+values);
  return values;
}


function crtSchedTablRowPerDate(rows) {
  let ar=[['day', ' ','8-9','9-10','10-11','11-12','12-13','13-14','14-15','15-16','16-17','17-18','18-19','19-20','20-21','21-22']];
  let trow=1;
  let pdate=rows[0][0];
  for (let i=0; i<rows.length;i++){
    if (pdate != rows[i][0]){
      trow++;
      pdate=rows[i][0];
    }
    if (! ar[trow]){
      ar[trow]=[rows[i][0], '','','','','','','','','','','','','','',''];
    }
    let val=rows[i][3]+'/'+rows[i][4];
    let hr=rows[i][1];
    let hfr=parseInt(rows[i][1]);
    let hto=parseInt(rows[i][2])-1;
    let col;
    if (hr.match(/:00/) && rows[i][2].match(/:00/) && hfr==hto){
    //if (hr.match(/:00/)){
      col=parseInt(hr)-6;
    } else {
      val=hr+'-'+rows[i][2]+  ' : '+val;
      //val=hr+  ' : '+val;
      col=1;
    }
    ar[trow][col]=ar[trow][col] ? (ar[trow][col] + ', ' + val) : val;
  }
  //Logger.log('b femptycols ar='+JSON.stringify(ar));
  let dc=findEmptyColumns(ar,1);
  dropColumns(ar,dc);
  return ar;
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
    //Logger.log('val= '+val);
    //xval.forEach(e => e.splice(0,0,targetSheet[i]));
    values=values.concat(val)
  }
  //Logger.log('getPupRs values='+JSON.stringify(values));
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
    } else if (mode=='hrs'){
      out.rows=crtSchedTablRowPerDate(out.rows);
    }
  } else {
    out.found='n';
    Logger.log('mail='+out.email+' Invalid');
  }
  return (out);
}

function updateAllDatesSheetMain() {
  collectParams();
  updateAllDatesSheetWork();
  //checkLog();
}
function updateAllDatesSheetWork() {
  let dts=getDtsOfSheetsToWorkOn(30,1,1);
  let ar=[];
  let addpday=15;
  collectDatesMetaAr(dts,ar,addpday);
  //Logger.log('ar='+JSON.stringify(ar));
  updateAllDatesSheet(ar,addpday);
}

function collectDatesMetaAr(dts, ar,addpday) {
  let row2set=2;
  for (let i=0; i<dts.length;i++){
    let sh=getShibutzSS().getSheetByName(dts[i]);
    let rows=sh.getLastRow()-1;//=ARRAYFORMULA('1 27/3/22'!A2:Q)
    let dt=dts[i].replace(/\d /,'');//tabname
    let f1='=ARRAYFORMULA({"'+dts[i]+'"&Y1:Y'+(rows+addpday)+"})";  //=ArrayFormula({"27/3/22"&Y2:Y500})
    let f2="=ARRAYFORMULA('"+dts[i]+"'!A2:R"+(rows+1)+")";
    //Logger.log('collect rows='+rows+' row2set='+row2set+' dt='+dt+' shnm='+sh.getName());
    ar.push([row2set,rows,f1,f2]);
    row2set+=rows;
  }
}

function updateAllDatesSheet(ar,addpday) {
  let sh=getShibutzSS().getSheetByName('allDays');
  sh.getRange(2,1,sh.getLastRow(),2).clear({ contentsOnly: true });
  //Logger.log('upd ar='+JSON.stringify(ar));
  //let lr=ar[ar.length-1][0];
  let lr=ar[ar.length-1][0]+addpday*(ar.length-1);
  let newlr=lr + ar[ar.length-1][1] -1;
  let rng=sh.getRange(2,1,lr-1,2);
  let rng_ar=rng.getValues();
  //Logger.log('upd lr='+lr+' rng_ar.length'+rng_ar.length+' rng_ar[0].length='+rng_ar[0].length);
  let added=0;
  for (let i=0; i<ar.length;i++){
    //Logger.log('upd i='+i+' ar[i][0]='+ar[i][0]);
    //Logger.log('upd  rng_ar[ar[i][0] - 2]='+rng_ar[ar[i][0] - 2]);
    rng_ar[ar[i][0] - 2 +added][0]=ar[i][2];
    rng_ar[ar[i][0] - 2 +added][1]=ar[i][3];
    added+=addpday;
  }
  rng.setValues(rng_ar);
  let c='=ARRAYFORMULA("," & J2:J'+newlr+' & "," & K2:K'+newlr+' & "," & L2:L'+newlr+' & "," & M2:M'+newlr+' & "," & N2:N'+newlr+' & "," & O2:O'+newlr+' & "," & P2:P'+newlr+' & ",")';
  //=arrayformula("," & J2:J871 & "," & K2:K871 & "," & L2:L871 & "," & M2:M871 & "," & N2:N871 & "," & O2:O871 & "," & P2:P871 & ",")
  sh.getRange('T2').setValue(c);
}


function sendMeetingReminderMain() {
  collectParams(3);
  let meet_ar=getMeetingsWithRemind();
  for (let i=0;i<meet_ar.length;i++){
    let res=isMeetingReminderDue(meet_ar[i]);
    if (res){
      remindMeeting(meet_ar[i], res);
    }
  }
  checkLog('mail','shibutz reminders');
}

function getMeetingsWithRemind() {
  let today1= new Date('5/25/22 12:55');
  //let today1= new Date();
  gp.shib_remind_today=convertTZ(today1,gp.dates_tz);
  //Logger.log("now="+gp.shib_remind_today);
  gp.shib_remind_today_shnm=getDtShNm(gp.shib_remind_today);
  let tomor1= new Date();
  gp.shib_remind_tomor=convertTZ(tomor1,gp.dates_tz);
  gp.shib_remind_tomor.setDate(gp.shib_remind_tomor.getDate() + 1);
  gp.shib_remind_tomor_shnm=getDtShNm(gp.shib_remind_tomor);
  //let qry='select * where R != ""';
  //let qry='select * where (A = "'+gp.shib_remind_today_shnm+'" or A = "'+gp.shib_remind_tomor_shnm+'") and (R matches ".+")';
  let qry='select * where (A = "'+gp.shib_remind_today_shnm+'" or A = "'+gp.shib_remind_tomor_shnm+'") and (R is not null)';
  //let qry="select * where  (R is not null)";
  let meet_ar=querySheet(qry, gp.shibutz_file_id, 'allDays', 1);
  //Logger.log("meet_ar="+JSON.stringify(meet_ar));
  return meet_ar;
}

function isMeetingReminderDue(meet) {//(now+remind hours) - meetingTime in [0 - 0.5)
// m3 r1 1n 1:15n 1.29n 1.30n -- 1:31y 1.40y 2y -- 2.05n
  let hm=meet[1].split(':');
  let mdt=getDtObjFromTabNm(meet[0]);
  mdt.setHours(hm[0],hm[1],0,0);
  let rems=meet[17].split(',');
  //Logger.log("rems="+rems+' len='+rems.length);
  for (let i=0;i<rems.length;i++){
    //Logger.log("i="+i);
    if (isNaN(rems[i])){
      writeLog('reminder hours is not a number: '+ JSON.stringify(meet));
      continue;
    }
    let diff=mdt.getTime()-(gp.shib_remind_today.getTime() + rems[i]*60*60*1000) ;
    //Logger.log('rems[i]='+rems[i]+" mdt="+mdt+' diff='+diff+ ' hm='+hm + ' rems[i]='+rems[i]);
    //if (Math.abs(diff) < 30*60*1000){
    if (diff >= 0 && diff < 30*60*1000){
      return rems[i];
    }
  }
  return 0;
}

function sendMailReminder(mail,msg) {
  let em={
        to: mail,
        subject: msg,
        htmlBody: "<p dir=RTL>"+msg+'</p>'
  }
  MailApp.sendEmail(em);
}

function remindMeeting(meet_ar, hours) {
  for (let i=9;i<16;i++){
    if (! meet_ar[i]){ continue  }
    let stu=getStuAr(meet_ar[i]);
    if (! stu[4]){ continue  }
    let msg='הי '+stu+', תזכורת להגיע לתגבור ב '+meet_ar[1] + ' עם ' + meet_ar[3];
    if (hours.length>1){
      msg='תזכורת: תגבור ביום '+meet_ar[0]+' בשעה ' + meet_ar[1] + ' עם ' + meet_ar[3];
    }
    Logger.log('msg= '+msg);
    if (gp.shib_reminder_type != 'mail'){
      if (stu[5]){
        Logger.log('mail mail='+stu[5]+' msg= '+msg);
        //sendMailReminder('dudigoo@gmail.com',msg);
        //sendMailReminder(stu[5],msg);
      } else {
        Logger('no mail for '+stu)
      }
    } else {
      if (stu[4]){
        Logger.log('sms phone='+stu[4]+' msg= '+msg);
        //sendSms(stu[4],msg);
      } else {
        Logger('no phone for '+stu);
      }
    }
  }
}

function findAllGroupsSchedMistakesMain(){
  collectParams();
  let dts=getDtsOfSheetsToWorkOn(6, 1, 1);
  writeLog('* find missing or double pupil');
  for (let i=0;i<dts.length;i++){
    findDtSheetMistakes(dts[i]);
  }
  let qry="select C,D,count(D)  group by C,D";
  let groups=querySheet(qry,gp.shib_arrival_order_file_id,"groupArrivalOrder",1);  
  writeLog("* find pupils w/o Math");
  let subj='מתמטיקה';
  for (let i=0;i<groups.length;i++){
    gp.shib_alldays_query="select T where H='"+groups[i][0]+"' and F='"+subj+"'";
    gp.shib_alfon_query="select B where A='" +groups[i][0] + "' and D="+groups[i][1];
    findPupilWithoutSubj();
  }
  findOvelappingMeetings();
  checkLog('mail', 'schedule mistakes',gp.shibutz_mail_to);
}

function findOvelappingMeetings(){
  writeLog('* check overlaps');
  let meet_a=querySheet("select A,B,C,D,F,G,H,J,K,L where D != '' order by D,A,B,C",gp.shibutz_file_id,"allDays",1);
  for (let i=1;i<meet_a.length;i++){
    if (meet_a[i][3] != meet_a[i-1][3] || meet_a[i][0] != meet_a[i-1][0]){ continue}
    if (meet_a[i][1] < meet_a[i-1][2]) {
      writeLog('overlapping meeting: date='+ meet_a[i][0]+ ' worker='+meet_a[i][3]+" from="+meet_a[i-1][1]+' to='+meet_a[i-1][2]+" from2="+meet_a[i][1]+' to2='+meet_a[i][2]);
    }
  }
}

function findDtSheetMistakes(shnm){
  Logger.log('shnm='+shnm);
  let qry="select B,C,D where A='"+(getDtObjFromTabNm(shnm).getDay()+1+"'");
  let groups=querySheet(qry,gp.shib_arrival_order_file_id,"groupArrivalOrder",1);
  Logger.log('shnm='+shnm+' groups.length='+groups.length);
  Logger.log(' groups='+groups);
  for (let i=0;i<groups.length;i++){
    findDtSheetGroupMistakes(shnm,groups[i][0],groups[i][1],groups[i][2]);
  }
}

function findDtSheetGroupMistakes(shib_sh_nm,time,clas,group){
  Logger.log('shib_sh_nm='+shib_sh_nm+' time'+time+' clas'+clas+' group'+group);
  let alfon_qry="select B where A='" +clas + "' and D="+group;
  shib_dt_qry="select I,J,K,L,M where G='"+clas+"' and A='"+time+"'";
  pupilSchedSheetMistakes(alfon_qry, shib_dt_qry, shib_sh_nm);
}


function findPupilWithoutMathMain() {
  collectParams();
  writeLog('starting');
  findPupilWithoutSubj();
  checkLog();
}

function findPupilWithoutSubj() {
  let ar1=querySheet(gp.shib_alfon_query,gp.pupil_alfon_id,"pupils",0);
  let ar2=querySheet(gp.shib_alldays_query,gp.shibutz_file_id,"allDays",0);
  let ar3 =ar2.flat();
  let ar4=[];
  ar3.forEach(e => ar4.push(e.split(',')));
  //Logger.log("ar4="+JSON.stringify(ar4));
  let res=arItemsNotInAnotherAr(ar1.flat(), ar4.flat(), 0);
}

function rangeItemsNotInAnotherRangeMain() {
  collectParams();
  writeLog('starting');
  pupilSchedSheetMistakes(gp.shib_alfon_query, gp.shib_dt_query, gp.shib_dt_sheet_nm);
  checkLog();
}

function pupilSchedSheetMistakes(alfon_qry,shib_dt_qry,shib_sh_nm) {
  let a1=[alfon_qry, gp.pupil_alfon_id, "pupils"];
  let a2=[shib_dt_qry, gp.shibutz_file_id, shib_sh_nm];
  let ar1=querySheet(a1[0],a1[1],a1[2],0);
  let ar2=querySheet(a2[0],a2[1],a2[2],0);
  let res=arItemsNotInAnotherAr(ar1.flat(), ar2.flat(), 1);
}

function updateShibCurrSheet(){
  let shnm=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  Logger.log('updateShibCurrSheet shnm='+shnm);
  collectParams();
  gp.shib_dates=getDtStrFromShNm(shnm);
  Logger.log('gp.shib_dates='+gp.shib_dates);
  shibutzDates();
  checkLog('mail', 'schedule issue',gp.shibutz_mail_to);
}

function createScheduleWindowsReportMain(){
  let sh=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  //Logger.log('updateShibCurrSheet shnm='+shnm);
  collectParams();
  let win_ss=SpreadsheetApp.openById('16Pig1bcyguk_yJlFZM9Jon8ygskxjYR4Y8jrdUdmhDM');
  let url=createScheduleWindowsReport(win_ss,sh);
  checkLog();
}

function tstHoles(){
  let win_ss=SpreadsheetApp.openById('16Pig1bcyguk_yJlFZM9Jon8ygskxjYR4Y8jrdUdmhDM');
  let sh=SpreadsheetApp.openById('1-VZOYe-sITfOgkDFEWAewEKZKGlnY9yXRCFoRaw-v_g').getSheetByName('2 7/12/22');
  createScheduleWindowsReport(win_ss,sh);
}

function createScheduleWindowsReport(win_ss,sh){
  let shnm=sh.getName();
  let win_sh=win_ss.getSheetByName(shnm);
  if (win_sh){
    if (! win_sh.getRange(2,5,win_sh.getMaxRows()-1,win_sh.getMaxColumns()-4).isBlank()){
      ui = SpreadsheetApp.getUi();
      ui.alert("Meeting exists in this window sheet. Please clear it first");
      return;
    }
    win_sh.getRange(2,1,win_sh.getMaxRows(),win_sh.getMaxColumns()).clear();
  } else {
    win_sh=win_ss.getSheetByName('tmpl').copyTo(win_ss);
    win_sh.setName(shnm);
  }
  win_sh.showSheet();
  //win_ss.setActiveSheet(win_sh);
  //win_ss.moveActiveSheet(0);
  let rows=sh.getRange(2,1,sh.getLastRow()+1,sh.getLastColumn()).getValues();
  //let row1=[ ... rows[0]];
  //Logger.log('rows[0]='+rows[0]);
  rows.sort((a, b) => {
      if (a[2]>b[2]){ return 1;}
      if (a[2]<b[2]){ return -1;}
      if (a[0]>b[0]){ return 1;}				
      if (a[0]<b[0]){ return -1;}
      if (a[1]>b[1]){ return 1;}				
      if (a[1]<b[1]){ return -1;}      
      return 0;
    }  
  );
  let res=findHolesInScheduleRows(rows);
  let row2add=res[0];
  let color=res[1];
  //row2add.splice(0,0,row1);
  //Logger.log('row2add[0].length'+row2add[0].length);
  win_sh.getRange(2,1,row2add.length,row2add[0].length).setValues(row2add);
  for (let i=0;i<res[1].length;i++){
    win_sh.getRange(color[i]+2,1,1,20).setBackground('#aad1f3');
  }
  win_sh.getRange(2,1,row2add.length,win_sh.getMaxColumns()).sort([{column: 1, ascending: true}, {column: 2, ascending: true}, {column: 3, ascending: true}]);
}

  function findHolesInScheduleRows(rows){
    let windows=[];
    let rows2add=[];
    let next_empty;
    let prev_empty;
    let cur_empty=rows[0].slice(4,13).every(element => element === "");
    for (let i=0;i<rows.length-1;i++){
      next_empty=rows[i+1].slice(4,13).every(element => element === "");
      //Logger.log('i='+i+' rows[i][2]='+rows[i][2]+' next_empty='+next_empty+' prev_empty='+prev_empty+' cur_empty='+cur_empty);
      if (cur_empty){ 
        let hole=false;
        if (! next_empty && rows[i][2] == rows[i+1][2]){
          rows2add.push(rows[i]);
          if (! prev_empty && i && rows[i][2] == rows[i-1][2]){
            windows.push(rows2add.length - 1);
            hole=true;
          }
        }
        if (! hole && ! prev_empty && i && rows[i][2] == rows[i-1][2]){
          rows2add.push(rows[i]);
        }
        //Logger.log('rows[i].length='+rows[i].length+ ' i='+i)
      }
      prev_empty=cur_empty;
      cur_empty=next_empty;
    }
    return [rows2add, windows];
  }
