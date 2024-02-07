
var is_eq_str;
var is_eq_col;

function getStuAr(s) {
  if (!gp.all_pupils_map){
    gp.all_pupils_map=getAllPupilsMap();
  }
  if (!gp.all_pupils_map[s]){
    //writeLog('unknown student:'+ s);
    return;
  }
  //Logger.log('**********stu='+s);
  return gp.all_pupils_map[s];
}


function addNewPupilsMain() {
  collectParams();
  let new_pupils_ss_id='1hZEe-Yb-a_m2od_Osorq-njdQw9z8kjJdRYyfee8r38';
  new_pupils_ss=SpreadsheetApp.openById(new_pupils_ss_id);
  var sheets=new_pupils_ss.getSheets();
  let new_rows=[];
  //for (let i=0;i<1;i++){
  for (let i=0;i<sheets.length;i++){
    let sh_nm=sheets[i].getName();
    let level=sheets[i].getName().slice(-1);
    Logger.log("level="+level);
    let data=querySheet('select B,C,D,E,F,G,H,I,J,K,L,M', new_pupils_ss_id, sh_nm, 1);
    if (! data.length){
      continue
    }
    data.forEach((row) => {
      if (row[11]){
        level=row[11];
      }
      new_rows.push([level, row[0], "", 8, row[1], "","","","",row[6],row[8],row[2],row[3],row[4],row[5]])
    });
  }
  let alfon_sh=getAlfonSS().getSheetByName('pupils');
  alfon_sh.getRange(alfon_sh.getLastRow()+1, 1, new_rows.length, new_rows[0].length).setValues(new_rows);
  checkLog();
}


function addNewPupilsSubjLevelsMain() {
  collectParams();
  let all_pupils=getAlfonKids();
  let mipui_first_row=230; //set line number of first new student, it will run untill the end
  let maakav_sh=getMaakavSS().getSheetByName('mipuiNewKids21');
  new_pupil_levels=maakav_sh.getRange(mipui_first_row,1,maakav_sh.getLastRow()-mipui_first_row+1,maakav_sh.getLastColumn()).getValues();
  for (let i=0;i<new_pupil_levels.length;i++){
    //if (! ['ז','ח','ט'].includes(new_pupil_levels[i][0] )){
    //  continue;
    //}
    let the_pupil=all_pupils.find((e) => e[1]==new_pupil_levels[i][2]);
    if (! the_pupil) {
      writeLog(new_pupil_levels[i][2] + ' not found in alfon');
      continue;
    }
    Logger.log("the_pupil="+the_pupil[1]);
    the_pupil[15]='רמה:' + new_pupil_levels[i][4].substring(0,1);
    let eng1=new_pupil_levels[i][6].substring(0,1);
    if (['3','4','5'].includes(eng1)){
      the_pupil[16]='הק:'+eng1;
    } else if (['א','ב','ג'].includes(eng1)){
      the_pupil[16]='רמה:'+eng1;
    } else {
      the_pupil[16]='רמה:' + ((eng1 =='A') ? 'א' : ((eng1=='B') ? 'ב' : 'ג'));
    }
    let heb1=new_pupil_levels[i][13];
    the_pupil[17]='רמה:' + ((heb1>=3.5) ? 'א' : ((heb1>2)? 'ב' : 'ג'));
    //if (i==1){break}
  }
  updateAllKids(all_pupils);
  checkLog();
}

function updateAllKids(kids) {
  let alfon_sh=getAlfonSS().getSheetByName('pupils');
  alfon_sh.getRange(2,1,alfon_sh.getLastRow()-1,alfon_sh.getLastColumn()).setValues(kids);
}

function getHistRows(stu) {//stu: str  , subj: []
  if (!gp.all_maakav_rows){
    let sh=getMaakavSS().getSheetByName('all');
    gp.all_maakav_rows=sh.getRange(6564,1,sh.getLastRow()-6563,13).getValues();
  }
  let rows=[];
    for (let i=0;i<gp.all_maakav_rows.length;i++){
    let e=gp.all_maakav_rows[i];
    try{
      let x=e[0].getTime();
      if (x>=gp.ab_sun_dt.getTime() && x<gp.ab_last_dt.getTime() && ab_subjects.includes(e[1]) && stu==e[5]){
        rows.push(e);
      }
    }catch(er){
       writeLog('ERROR er='+er+' e='+e);
      let g=e.getTime();
    }
  }
  //Logger.log('getHistRow rows='+rows);
  return rows;
}

function getMashovScoresMain() {
  collectParams();
  var folder = DriveApp.getFolderById(gp.mashov_scores_dir_id);
  convertXlsx2sheets(folder);
  var add_rows=[];

  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    var file = files.next();
    //Logger.log('ss nm='+ss.getName());
    let sh=SpreadsheetApp.open(file).getSheets()[0];
    Logger.log('g fnm='+file.getName()+' sh nm='+sh.getName());
    getScoresFromMashovFile(sh, add_rows);
  }
  addMashovScoreRows(add_rows);
  checkLog();
}

function updatePupilEmailFromMashovMain() {
  collectParams();
  var folder = DriveApp.getFolderById(gp.mashov_scores_dir_id);
  convertXlsx2sheets(folder);
  var add_rows=[];
  
  let kids_ar=getAlfonKids();
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    var file = files.next();
    //Logger.log('ss nm='+ss.getName());
    let sh=SpreadsheetApp.open(file).getSheets()[0];
    Logger.log('g fnm='+file.getName()+' sh nm='+sh.getName());
    updatePupilsArEmail(sh);
  }
  Logger.log('updating...');
  gp.pupilAlfonRange.setValues(gp.pupilAlfonAr);
  checkLog();
}

function updatePupilsArEmail(sh) {
  let sh_vals= sh.getRange(4,1,sh.getLastRow()-2,sh.getLastColumn()-2).getValues();
  //Logger.log('sh_vals.length='+sh_vals.length+' sh.getLastRow()'+sh.getLastRow());
  //Logger.log('sh_vals[0]='+sh_vals[0]);
  for (let i=1;i<sh_vals.length;i++){
    //Logger.log('i='+i+' 0='+sh_vals[i][3]+' '+sh_vals[i][2]);
    
    if (sh_vals[i][13]){
      let nm=getKidMerkazNm(sh_vals[i][2]+' '+sh_vals[i][3], sh_vals[i][4]);
      if (nm){
        //Logger.log('nm='+nm);
        let alfmail=getStuAr(nm)[5];
        if (alfmail && alfmail != sh_vals[i][13] ){
          writeLog('different nm='+nm+' alfon mail='+alfmail+' mashov mail='+sh_vals[i][13]);
        } else {
          if (alfmail == sh_vals[i][13]){
            writeLog('same  nm='+nm+' alfon mail='+alfmail+' mashov mail='+sh_vals[i][13]);
          } else {
            writeLog('new  nm='+nm+' alfon mail='+alfmail+' mashov mail='+sh_vals[i][13]);
            getStuAr(nm)[5]=sh_vals[i][13].toLowerCase();
          }
        }
      }

    }
  }
}

function addMashovScoreRows(add_rows) {
  let grades_sh = getMaakavSS().getSheetByName('schoolGrades');
  //let grades_sh = SpreadsheetApp.openById('1bXCdG6Vyo6RpIhH9S6RoEieivfp_iz8kaNfWgWQkRH8').getSheetByName('schoolGrades');
  if (grades_sh.getLastRow()>1){
    grades_sh.getRange(2,1,grades_sh.getLastRow()-1,4).clear();
  }
  grades_sh.getRange(2,1,add_rows.length,4).setValues(add_rows);
}

function getScoresFromMashovFile(sh, add_rows) {
  let sh_vals= sh.getRange(3,3,sh.getLastRow()-2,sh.getLastColumn()-2).getValues();
  Logger.log('sh_vals.length='+sh_vals.length+' sh.getLastRow()'+sh.getLastRow());
  //Logger.log('sh_vals[0]='+sh_vals[0]);
  for (let i=1;i<sh_vals.length;i++){
    //Logger.log('i='+i+' nm='+sh_vals[i][0]);
    getStudentMashovScores(sh_vals[0], sh_vals[i], add_rows,i);
  }
}

function getStudentMashovScores(sh_vals0, stud_vals, add_rows,i) {
  if (! stud_vals[0] && ! stud_vals[1]){
    return;
  }
  let merkaz_nm=getKidMerkazNm(stud_vals[0], stud_vals[1]);
  if (!merkaz_nm){
    writeLog(" invalid. row="+ (i+4) +" level="+stud_vals[1]+' name='+stud_vals[0]);
    return;
  }
  for (let j=3;j<stud_vals.length;j++){
    if (! stud_vals[j]) {continue;}
    if ( sh_vals0[j].startsWith('חנ"ג')) {continue;}
    add_rows.push([merkaz_nm, stud_vals[1], sh_vals0[j], stud_vals[j]]);
  }
}

function getAlfonSS() {
  if (!gp.alfon_ss){
    gp.alfon_ss=SpreadsheetApp.openById(gp.pupil_alfon_id);
  }
  return gp.alfon_ss;
}

function getAlfonKids(just_name) {
  if (! gp.pupilAlfonAr){
    let alfon_sh = getAlfonSS().getSheetByName('pupils');
    //let alfon_sh = SpreadsheetApp.openById('1yrL132sLyUUzRruG5EzivGOk8uC88p7KPRC9NwAWI6A').getSheetByName('pupils');
    gp.pupilAlfonRange = alfon_sh.getRange(2,1,alfon_sh.getLastRow()-1,alfon_sh.getLastColumn());
    gp.pupilAlfonAr = gp.pupilAlfonRange.getValues();
  }
  let result = just_name ? gp.pupilAlfonAr.map(e => e[1]) : gp.pupilAlfonAr;

  return(result);
}

function getKidMerkazNm(str,level) {
  let rnm='';
  let is_eq_str=str;
  let inx=getAlfonKids().findIndex(e => e[0]===level && e[2]===str);
  //Logger.log('str='+str+' inx='+inx+' level='+level);

  if (inx>-1){
    rnm=getAlfonKids()[inx][1];
  } else {
    let nma=str.split(' ');
    if (nma.length>2){
      writeLog('long name missing in alfon:'+ str+' level='+level);
      rnm = str;
    } else {
      is_eq_str=nma[1] +' ' + nma[0];
      let inx2=getAlfonKids().findIndex(e => e[0]===level && e[1]=== is_eq_str);
      if (inx2 == -1) {
        writeLog(" not in alfon. name="+str+' level='+level+ ' swapped='+is_eq_str);
      } else {
        rnm=is_eq_str;
      }
    }
  }
  return(rnm);
}

// for quiz triggers
/*function setQuiztriggersMain() {
  collectParams();
  var files=getSubFoldersFiles('1YP5aziOBgpzO1GS35z3yhY2op0qXdz2W');
  for (let i=0;i<files.length;i++){
    let file_rows=setSubmitTrigger(files[i]);
    
  }
}


function setSubmitTrigger(ss) {
  Logger.log('file='+ss.getName());
  try {
    let trigs=ScriptApp.getUserTriggers(ss);
    if (trigs.length){
      Logger.log('trigger exists' );
      return;
    }
    ScriptApp.newTrigger("handleQuizResponse")
  .forSpreadsheet(ss)
  .onFormSubmit()
  .create();
  } catch(err) {
    Logger.log('err='+err);
  }
}*/

function handleQuizResponse(e) {
  // editable response, add code and lib to sheet
  collectParams();

  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let formsh=ss.getSheets()[0];
  let rownum=e.range.getRow();
  let vals=formsh.getRange(rownum,1,1,3).getDisplayValues()[0];
  let vals2=formsh.getRange(rownum,1,1,3).getValues()[0];
  //Logger.log('vals='+vals);
  let furl=ss.getFormUrl();
  let form = FormApp.openByUrl(furl);
  //Logger.log('furl='+furl+' title='+form.getTitle());
  let form_nm = ss.getName().replace(/\(.*\)$/,'');
  if (! e.values[1] && ! e.values[2]){
    Logger.log('submit but score and name havent changed so skipping. form_nm='+form_nm+ ' rownum='+rownum+ ' name='+vals[2]);
    return;
  }
  let responses = form.getResponses();
  //let resp_url=responses[responses.length - 1].getEditResponseUrl();
  let resp_url=findResponseUrl(responses,vals2[0].setMilliseconds(0));
  let row=[vals[2], vals[0], form_nm, vals[1], resp_url, furl ]
  //Logger.log('row='+row);
  sh=getMaakavSS().getSheetByName('allQuiz');
  sh.appendRow(row);
}

function findResponseUrl(responses,times) {
  for (let i = responses.length-1; i >=0; i--) {
    //Logger.log('i='+i+' responses[i].getTimestamp()='+responses[i].getTimestamp()+' times='+times);
    if (responses[i].getTimestamp().setMilliseconds(0) == times){
      return responses[i].getEditResponseUrl();
    }
  }
  Logger.log('didnt find response for vals='+vals);
}

function getPupilByMail(mail) {
  let query = 'select A,B,C,D,E,F,G,H where F = "'+mail+'"';
  let res=querySheet(query,gp.pupil_alfon_id,'pupils');
  Logger.log('resx='+res);
  if (!res || res.length<1){
    Logger.log('failed query');
    return;
  }
  return res;
}

function getAllPupilsMap() {
  if (!gp.all_pupils_map){
    let res=getAlfonKids();
    //let query = 'select *';
    //let res=querySheet(query,gp.pupil_alfon_id,'pupils');
    //if (!res || res.length<2){
    //  Logger.log('failed query');
    //  return;
    //}
    gp.all_pupils_map={};
    res.forEach(e => gp.all_pupils_map[e[1]]=e);
  }
  return gp.all_pupils_map;
}

function getGroupsDict() {
  let query = 'select A, B';
  let res=querySheet(query,gp.pupil_alfon_id,'group');
  if (!res || res.length<2){
    Logger.log('failed query');
    return {};
  }
  map={};
  for (let i=0;i<res.length;i++){
    map[res[i][0]]=res[i][1];
  }
  return map;
}

function findInvalidMipuiNames() {
  collectParams();
  let query2="select A, B, C, D ";
  let quizs=querySheet(query2,gp.maakav_file_id,'mipuiNewKids21',1);
  for (let i=0;i<quizs.length;i++){
    if (! getAllPupilsMap()[quizs[i][2]]){
      writeLog('invalid: i='+(i+2)+' name='+quizs[i][2]);
    }
  }
  checkLog();
}