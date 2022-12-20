
var fnm_pref='שעות עבודה';
var shared_dir_id='1XgtB0ofSqt6IPl-IPDblvjzOYJM4Me6q';
var g_mipui_dir;

var wrkrsq_file_id='';
var wtpm='מ'; wtps='ס'; var wtph='ח';
var wtyp2master_id = {}; 
var wtyp2fol_id = {};
var do_mipui;
var max_wrkrs=120;
var absent_str='לא הגיע';

function addw() {
  collectParams();
  var res=confirm_popup();
  if (res=='NO') { return; }
  let worker_rows=getWorkersRows();
  for (var i in worker_rows) {
    var wrkr=getWorker(worker_rows[i]);
    addWorker(wrkr);
  }
  checkLog();
}

function sendNewWrkrMail(worker_rows) {
    var perrs=["שלום "+ tabnm.split(' ')[0]+",\n"].concat(werrs);
    perrs.push('<a href="https://tinyurl.com/ya7ptvoq">הסבר לדיווח שעות</a>');
    perrs.push('<a href="'+furl+'">דיווח שעות שלי</a>');
    var msg="<p dir=RTL>"+ perrs.join("<br>") + '</p>';
    MailApp.sendEmail('mlemida.ryam@gmail.com','מורים חדשים שנוספו למרכז למידה', msg, {htmlBody: msg});
}

function getWorkersRows() {
  if (! gp.worker_rows){
    let result = parseNumbers(gp.wrkrs_row_str);
    if (result[0]) {
      writeLog("Error in rows field");
    } else {
      gp.worker_rows = result[1];
    }
  }
  return gp.worker_rows;
}

function delw() {
  init();
  collectParams();
  let worker_rows=getWorkersRows();
  for (var i in worker_rows) {
    var wrkr=getWorker(worker_rows[i]);
    delWorker(wrkr);
  }
  SpreadsheetApp.flush();
  Logger.log('Done delete');
}

function getAllWorkers() {
  ws={}; ws_by_mail={};
  var wa=getWorkersSh().getRange(2,1,getWorkersSh().getLastRow()-1,getWorkersSh().getLastColumn()).getValues();
  for (var i=0;i<wa.length;i++) {
    ws[wa[i][0]]=getWorker(0,wa[i]);
    ws_by_mail[wa[i][2]]=ws[wa[i][0]];
  }
  gp.all_wrkrs=ws;
  gp.all_wrkrs_by_mail=ws_by_mail;
}

function getWorkerByName(name){
  var wrkr={};
  if (! gp.all_wrkrs){
    getAllWorkers();
  }
  return gp.all_wrkrs[name];
}
 
function getWorkerByMail(mail){
  var wrkr={};
  if (! gp.all_wrkrs){
    getAllWorkers();
  }
  //Logger.log('getWorkerByMail mail='+mail);
  //Logger.log('all_wrkrs_by_mail='+JSON.stringify(gp.all_wrkrs_by_mail));
  return gp.all_wrkrs_by_mail[mail];
}
 
function getWorkersSh() {
  if (! gp.workers_sh){
    gp.workers_sh=SpreadsheetApp.openById(gp.wrkrs_ss_id).getSheetByName('עובדי מרכז');
  }
  return gp.workers_sh;
}

function getWorker(rown,wa) {
  var wrkr={};
  if (!wa){
    wa=getWorkersSh().getRange(rown,1,1,15).getValues()[0];
  }
  wrkr['name']=wa[0];
  wrkr['phone']=wa[1];
  wrkr['mail']=wa[2];
  wrkr['tz']=wa[4];
  wrkr['typ']=wa[5];
  wrkr['subj']=wa[6];
  wrkr['popu']=wa[7];
  //wrkr['share']=wa[8];
  wrkr['subj_popu']=wrkr['subj'] + ' ' + wrkr['popu'];
  if (wrkr.subj.match(/-([חטי]|יא|יב)$/)){
    wrkr['subj_popu']=wrkr['subj']
  }
  wrkr['win_formula']=wa[10];
  //writeLog('worker:name='+wrkr.name+' type='+wrkr.typ + ' mail='+wrkr['mail']);
  return wrkr;
}

function getHoursMasterFile() {
  if (! gp.masterFile) {
    gp.masterFile=DriveApp.getFileById(gp.hours_master_id);
  }
  return gp.masterFile;
}

function getHoursMasterFileSS() {
  if (! gp.master_ss) {
    gp.master_ss = SpreadsheetApp.openById(gp.hours_master_id);
  }
  return gp.master_ss;
}

function addWorker(wrkr) {
  wfname=fnm_pref + ' ' + gp.heb_year + ' ' + wrkr.name;
  Logger.log('adding ' + wrkr.name);
  var wtp=wrkr.typ.substring(0,1);
  if (wtp in wtyp2fol_id) {
    typ_folder = DriveApp.getFolderById(wtyp2fol_id[wtp]);
    file=getHoursMasterFile();
    //try {
      var wfile=file.makeCopy(typ_folder);
      wfile.setName(wfname);
      wfile.addEditor(wrkr.mail);
      var ss = SpreadsheetApp.open(wfile);
      //ScriptApp.newTrigger('addAbsent').forSpreadsheet(ss).onEdit().create();

      var sheet = ss.getSheetByName('name');
      sheet.getRange('B2').setValue(wrkr.name);
      sheet.getRange('A2').setValue(wrkr.tz);
      rmProtections(sheet);
      sheetProtection(sheet, 3);

      sheet = ss.getSheetByName('lists');
      rmProtections(sheet);
      sheetProtection(sheet, 3);

      ss.getSheetByName('template').setName(gp.sheet2show);
      addMonthes(ss,wrkr,1);
      Logger.log('created shaot file for '+ wrkr.name);
    //} catch (e) {
    //  writeLog('makeCopy err. e=' + e);
    //}
  } else {
    writeLog('no hour report for worker type '+wtp);
  } 
  if (wtp != 'ח'){
    try {
      Logger.log('viewer mail='+ wrkr.mail + ' fname='+get_shared_dir().getName());
      get_shared_dir().addViewer(wrkr.mail);   
      Logger.log('added viewer '+ wrkr.name);
    } catch (e) {
      writeLog('addViewer err. e=' + e);
    }
  }
  writeLog('Done='+wrkr.name);
}

function get_shared_dir(){
  if (! gp.shared_dir){
    gp.shared_dir=DriveApp.getFolderById(shared_dir_id);
  }
  return gp.shared_dir;
}

function delWorker(wrkr){
  try {
    get_shared_dir().removeViewer(wrkr.mail);
    writeLog('removed viewer '+ wrkr.name);
  } catch (e) {
    writeLog('rmViewer err. e=' + e);
  }
  if (wrkr.mipui == 'y'){
    try {
      g_mipui_dir.removeEditor(wrkr.mail);
      writeLog('removed editor '+ wrkr.name);
    } catch (e) {
      writeLog('removeEditor err. e=' + e);
    }
  } 
}  

function parseRange(rng) {
  writeLog('parseRng:rng='+rng);
  var ar2 = rng.split("-");
  var nwar = [];
  for (var i=Number(ar2[0]); i<= Number(ar2[1]); i++){
    nwar.push(i);
  }
  //writeLog('nwar='+nwar);
  return nwar;
}  
  
function parseNumbers(str) {
  var regNum = new RegExp("^[0-9]+$");
  var regRng = new RegExp("^[0-9]+[-][0-9]+$");
  var ar = str.split(",");
  for(var i = 0; i < ar.length; i++){
    //writeLog("i="+i+" ari="+ar[i]);
    var testNum = regNum.test(ar[i]);
    //writeLog("testNum="+testNum);
    if (testNum){
      //writeLog("just number");
    } else { // range 
      var testRng = regRng.test(ar[i]);
      //writeLog("testRng="+testRng);
      if (testRng){
        //writeLog("is range");
        var rngAr = parseRange(ar[i]);
        if (rngAr.length<2){
          return [2];
        }
        ar.splice(i,1);
        for (j=0;j<rngAr.length;j++){
          ar.splice(i+j,0,rngAr[j]);
        }
        i = i -1; //i = i + rngAr.length - 1
      } else { 
        writeLog("invalid...");
        return [1];
      }
    }
  }
  return [0,ar];
}

function rmUnauthWorkers(){
  collectParams();
  var viewers= get_shared_dir().getViewers();
  //var editors= get_shared_dir().getEditors();
  for (var i=0; i<viewers.length; i++){
    checkUserAuth(viewers[i],'v');
  }
  var editors= g_mipui_dir.getEditors();
  for (var i=0; i<editors.length; i++){
    checkUserAuth(editors[i],'e');
  }
  writeLog("End");
}

function checkUserAuth(usr, auth){
  console.info('usr='+usr.getEmail);
  var found=0;
  for (var i=2; i<max_wrkrs; i++) {
    var wrkr_mail=getWorkersSh().getRange(i,3).getValue().toLowerCase();
    var wrkr_mail2=getWorkersSh().getRange(i,4).getValue().toLowerCase();
    var wrkr_share=getWorkersSh().getRange(i,9).getValue();
    var wrkr_mipui=getWorkersSh().getRange(i,8).getValue();
    var usrmail=usr.getEmail().toLowerCase();
    if (getWorkersSh().getRange(i,1).getValue()== ''){
      break;
    }
    if (auth=='v' && (usrmail==wrkr_mail || usrmail==wrkr_mail2)  && wrkr_share == 'y') {
      found=1;
      break;
    }
    if (auth=='e' && (usrmail==wrkr_mail || usrmail==wrkr_mail2) && wrkr_mipui == 'y') {
      found=1;
      break;
    }
  }
  if (found == 0) {
    if (auth=='v'){
      get_shared_dir().removeViewer(usr.getEmail());
      writeLog("Removed shared dir viewer: "+usr.getEmail());
    }
    if (auth=='e'){
      g_mipui_dir.removeEditor(usr.getEmail());
      writeLog("Removed mipui dir editor: "+usr.getEmail());
    }
  }
}

function cpMasterSheetToWorkers(){
  g_func2run='add_monthes';
  var res=confirm_popup();
  if (res=='NO') { return; }
  collectParams();
  iterateMain();
}

function addMonthes(ss,w,addw){
  var tmpl_sh = getHoursMasterFileSS().getSheetByName('template');
  for (var i=0;i<gp.g_month_name_ar.length;i++){
    var new_sh=ss.getSheetByName(gp.g_month_name_ar[i]);
    if (! new_sh){
      new_sh=tmpl_sh.copyTo(ss);
      new_sh.setName(gp.g_month_name_ar[i]);
    }
    try {
      new_sh.hideSheet();
    } catch (e) {
    }
    if (! addw || gp.sheet2show != gp.g_month_name_ar[i]){
      sheetProtection(new_sh, 3);
    }
    clearAcademic(new_sh, w.typ);
    Logger.log('created sheet=' + gp.g_month_name_ar[i]);
  }
  if (addw && gp.sheet2show != ''){
    showsh = ss.getSheetByName(gp.sheet2show);
    sheetProtection(showsh, 1);
    hideAllSheetsExcept(ss, gp.sheet2show);
  }

}

function clearAcademic(sh,typ) {
  if (typ.substring(0,1) != 'מ'){
    sh.getRange('A41:E41').clearContent().clearFormat();
  }
}

function unshareWorkerSSMain(){
  collectParams();
  g_func2run='unshareWorkerSS';
  iterateMain();
}

function unshareWorkerSS(ss,w) {
  let cur_sh=ss.getSheetByName(gp.g_month_name);
  try {
    ss.removeEditor(w.mail);
  } catch (err) {
    Logger.log('removeEditor error:'+err);
  }
}


function shareWorkerSSMain(){
  collectParams();
  g_func2run='shareWorkerSS';
  iterateMain();
}

function shareWorkerSS(ss,w) {
  try {
    ss.addEditor(w.mail);
  } catch (err) {
    Logger.log('addEditor error:'+err);
  }
}

function switchActiveMonthMain(){
  collectParams();
  g_func2run='switchActiveMonth';
  iterateMain();
}

function switchActiveMonth(ss,nm) {
  let cur_sh=ss.getSheetByName(gp.g_month_name);
  let w=getWorkerByName(nm);
  //Logger.log('nm='+nm+' gp.g_month_name'+gp.g_month_name);
  //Logger.log('sheetProtection s cur sh nm='+cur_sh.getName());
  if (cur_sh){
    sheetProtection(cur_sh, 2);
  } else {
    writeLog('missing sheet '+gp.g_month_name+ ' in file ' + ss.getName());
  }
  //Logger.log('nm='+nm+' gp.g_month_name'+gp.g_month_name);
  showsh = ss.getSheetByName(gp.sheet2show);
  if (showsh){
    //Logger.log('gp.sheet2show='+gp.sheet2show);
    sheetProtection(showsh, 1);
    //Logger.log('hideAllSheetsExcept s ');
    hideAllSheetsExcept(ss, gp.sheet2show);
  } else {
    writeLog('missing sheet '+gp.sheet2show+ ' in file ' + ss.getName());
  }
  //Logger.log('hideAllSheetsExcept e ');
}

function addSomeRows(sh,wnm) {
  var rn=findTotalRow(sh);
  //Logger.log('rn='+rn);
  var re=sh.getRange(rn-19, 1, 18, 4);
  var re2=sh.getRange(rn-19, 6, 18, 3);
  if (! re.isBlank() || ! re2.isBlank() ){
    Logger.log('rn='+rn+' tabnm='+wnm);
    try {
      SpreadsheetApp.flush();
      sh.insertRows(rn-1, 30);
      Logger.log('added rows to '+ wnm + ' '+ gp.g_month_name);
    } catch(e) {
      writeLog('failed add rows to:'+wnm+' month='+sh.getName()+ ' e='+e);
    }
  }
}

function addAbsentEvent(e){
  wrkrEditEvent(e);
}

function wrkrEditEvent(e){
  let sh = e.source.getActiveSheet();
  let ac = sh.getActiveCell();
  let acol = ac.getColumn();
  let arow = ac.getRow();
  let aval = ac.getValue();
  //Logger.log('ac.getValue='+ac.getValue());
  if(arow<8 || arow >sh.getLastRow()-4){
    return;
  }
  if(acol == 6){
    //Logger.log('grade selected row='+aRow);
    var trange = sh.getRange(arow, acol + 2);
    var sourceRange = e.source.getRangeByName(aval);
    var rule = SpreadsheetApp.newDataValidation().requireValueInRange(sourceRange, true).build();
    trange.setDataValidation(rule);
    //Logger.log('end');
  }

  if (acol == 7){
    if ( aval == 'חשב חיסור'){
      addAbsentStuRow();
    } else if ( aval == 'הוסף שורה'){
      Logger.log('insert row after '+arow);
      if (arow>7 && (arow < (sh.getLastRow()-7))){
        Logger.log('insert row after '+arow+ 'oldval='+e.oldValue);
        ac.setValue(e.oldValue);
        sh.insertRowAfter(arow);
      }
    }
  }

  if (acol == 8) {
    let newValue=e.value;
    let oldValue=e.oldValue;
    //Logger.log('keys selected. old='+oldValue+' new='+newValue);
    if (e.value) {
      if (!e.oldValue) {
        ac.setValue(newValue);
      } else {
        ac.setValue(oldValue+','+newValue);
      }
    }
  }
}

function addAbsentStuRow(){
  collectParams();
  let meeting_details=getSelectedMeetingDetails();
  if (meeting_details.sts=='er'){
    writeMsgInActivity(meeting_details, meeting_details.msg);
    //SpreadsheetApp.getUi().alert(meeting_details.msg);
    return;
  }
  //Logger.log('meeting_details'+JSON.stringify(meeting_details));
  let reg=getStuRegis2Meeting(meeting_details);
  //Logger.log('reg='+JSON.stringify(reg));
  if (reg.sts=='er'){
    writeMsgInActivity(meeting_details, reg.msg);
    //SpreadsheetApp.getUi().alert(reg.msg);
    return;
  }
  // all - reported
  let diff = reg.names.filter(x => !meeting_details.reported.includes(x));
  if (diff.length){
    meeting_details.row2add[7] = diff.join(',');
    Logger.log('diff='+JSON.stringify(diff));
    meeting_details.sh.getRange(meeting_details.row_num2upd,1,1,11 ).setValues([meeting_details.row2add]);
  } else {
    writeMsgInActivity(meeting_details,'no one absent');
    return;
  }
}


function writeMsgInActivity(meeting_details,msg){
  Logger.log('msg='+msg);
  meeting_details.sh.getRange(meeting_details.row_num2upd, 9 ).setValue(msg);
}

function getStuRegis2Meeting(md){
  let q;
  if (md.sh2look=='history'){
    q='select J,K,L,M,N,O,P where D="'+md.tname+'" and B="' + md.row2add[2] +'" and A = date "'+ getYMDStr(md.row2add[1]) + '"';
    //q='select J,K,L,M,N,O,P where D="'+md.tname+'" and B="' + md.row2add[2] +'" and C="' + md.row2add[3] +'" and A = date "'+ getYMDStr(md.row2add[1]) + '"';
  }else{
    let dow=md.row2add[1].getDay()+1;
    let str=dow + ' ' + getFmtDtStr(md.row2add[1]);
    q='select J,K,L,M,N,O,P where D="'+md.tname+'" and B="' + md.row2add[2] +'" and A = "'+ str + '"';
    //q='select J,K,L,M,N,O,P where D="'+md.tname+'" and B="' + md.row2add[2] +'" and C="' + md.row2add[3] +'" and A = "'+ str + '"';
  }
  let vals=querySheet(q, gp.shibutz_file_id, md.sh2look);
  //Logger.log('sched='+JSON.stringify(vals));
  if (! vals.length){
    return {'sts':'er','msg':'Lesson from above row not found in schedule'}
  }
  expandGroup2members(vals,1,0,7);
  //Logger.log('sched2='+JSON.stringify(vals));
  return {sts:'ok', 'names': chomp(vals[0][0]).split(',')};
}

function getSelectedMeetingDetails(){
  let ret={};
  md=[];
  ret.tname=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('name').getRange(2,2).getValue();
  //ret.tname='חוה העוגן';
  ret.sh=SpreadsheetApp.getActiveSheet();
  //ret.sh=SpreadsheetApp.openById('1z-nFGcmf9_ip6n-GawTYlF8g_GujrOQdaHGIZuLWENw').getSheetByName('16.4-15.5');
  let selection = ret.sh.getSelection();
  let srng =  selection.getActiveRangeList().getRanges()[0];
  //Logger.log('srng.getRow()='+srng.getRow()+' shnm'+ret.sh.getName());
  ret.row_num2upd=srng.getRow();
  //ret.row_num2upd=26;
  if (ret.row_num2upd<9){
    return {'sts':'er','msg':'invalid cell selected. Select a cell in the row to fill'}
  }
  let w_ar=ret.sh.getRange(8,1,ret.row_num2upd-8,8).getValues();
  md[8]='*חיסור מחושב';
  md[10]=absent_str;
  let reported='';
  
  for (let i=1; i< w_ar.length;i++){// fill empty cell with prev row value
    for (let c=0;c<4;c++){
      if (!w_ar[i][c] ){
        w_ar[i][c]=w_ar[i-1][c]
      }
    }
  }
  for (let i=w_ar.length-1;i>=0;i--){// while date/time empty or changes
    //Logger.log('i='+i+' w_ar[i]='+JSON.stringify(w_ar[i]));
    //Logger.log(' md='+JSON.stringify(md));
    if ( (md[1] && (w_ar[i][1].getTime() != md[1].getTime())) ||  (md[2] && (w_ar[i][2] != md[2])) ){
      Logger.log('break');
      break;
    }
    if (!md[0] && w_ar[i][0]){
      md[0]=w_ar[i][0];
    }
    if (!md[1] && w_ar[i][1]){
      md[1]=w_ar[i][1];
    }
    if (!md[2] && w_ar[i][2]){
      md[2]=w_ar[i][2];
    }
    if (!md[3] && w_ar[i][3]){
      md[3]=w_ar[i][3];
    }
    if (!md[5] && w_ar[i][5]){
      md[5]=w_ar[i][5];
    }
    if (!md[6] && w_ar[i][6]){
      md[6]=w_ar[i][6];
    }                    
    reported += (','+ w_ar[i][7]); 
  }
  ret.row2add=md;
  ret.reported=chomp(reported).split(',');
  ret.sts='ok';
  let today = new Date();
  today.setHours(0,0,0,0);
  ret.sh2look= (today>md[1]) ? 'history' : 'allDays';
  return ret;
}
