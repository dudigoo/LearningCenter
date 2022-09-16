var quiz_dir_id='1YP5aziOBgpzO1GS35z3yhY2op0qXdz2W';

function findNoQuizPupilsMain() {
  collectParams();
  writeLog('Start find pupils');
  let query="select A, B, D where A='ז' or A='ח' or A='ט'";
  let pps=querySheet(query,gp.pupil_alfon_id,'pupils',1);
  let query2="select A, B, C ";
  let quizs=querySheet(query2,gp.maakav_file_id,'allQuiz',1);
  quizs.forEach(e => {e[0]=chomp(e[0])})
  let res;
  res=findNoQuizPupilsWord(pps,quizs,'English');
  writeLog('english:'+res.length);
  writeLog(res.join('\n'));
  res=findNoQuizPupilsWord(pps,quizs,'מבדק');
  writeLog('math:'+res.length);
  writeLog(res.join('\n'));
  checkLog();
}



function findQuizInvalidNamesMain() {
  collectParams();
  let query2="select A, B, C, D ";
  let quizs=querySheet(query2,gp.maakav_file_id,'allQuiz',1);
  for (let i=0;i<quizs.length;i++){
    if (! getAllPupilsMap()[quizs[i][0]]){
      writeLog('invalid: i='+(i+2)+' name='+quizs[i][0]+' dt='+quizs[i][1]+' subj='+quizs[i][2]);
    }
  }
  checkLog();
}

function findNoQuizPupilsWord(pps,quizs,pat) {
  const regex = new RegExp(pat);
  found_ar=[];
  for (let i=0;i<pps.length;i++){
  //for (let i=12;i<13;i++){
    let f;
    Logger.log('pps[i][1]='+pps[i][1]);
    //for (let j=66;j<72;j++){
    for (let j=0;j<quizs.length;j++){
      if (quizs[j][0]==pps[i][1]){
        Logger.log('quizs[j][0]='+quizs[j][0]+' pps[i][1]='+pps[i][1]);
        if (quizs[j][2].match(regex)){
          f=1;
          break;
        }
      }
    }
    if (!f){
      found_ar.push(pps[i][0]+' '+ pps[i][2]+' '+ pps[i][1]);
    }
    Logger.log('i='+i+'');
  }
  found_ar.sort();
  return found_ar;
}

function updateQuizResultsMain() {
  collectParams();
  let days_changed=1;
  //quiz_dir_id='1IPmlBkL_f5V7RxuewRH7Z50_SpaXLABK';
  let files_ar=[];
  let cur_all_quiz_sh=getMaakavSS().getSheetByName('allQuiz');
  let cur_all_quiz_ar=cur_all_quiz_sh.getDataRange().getValues();
  let res=getFolderIdFilesRecursivly(quiz_dir_id,'application/vnd.google-apps.spreadsheet' , files_ar, 100);
  //Logger.log('files_ar '+files_ar);
  let dt= new Date();
  dt.setHours(dt.getHours()-26*days_changed);
  Logger.log('ignore files modified earlier then '+dt);
  for (let i=0;i<files_ar.length;i++){
    if (files_ar[i].getLastUpdated()<dt){
      continue;
    }
    let scores_ar=getScoresFromFile(SpreadsheetApp.open(files_ar[i]));
    if (! scores_ar){
      continue;
    }    
    //Logger.log('fixing file '+files_ar[i].getName()+ ' scores_ar='+JSON.stringify(scores_ar));
    cur_all_quiz_ar=replaceOldScoresWithNewScores(cur_all_quiz_ar,scores_ar,2);
  }
  cur_all_quiz_sh.getRange(1,1,cur_all_quiz_ar.length,cur_all_quiz_ar[0].length).setValues(cur_all_quiz_ar);
}

function replaceOldScoresWithNewScores(full_ar,replace_ar, comp_pos) {
  //Logger.log('comp_pos='+comp_pos+' replace_ar.length='+replace_ar.length);
  let match_val=replace_ar[0][comp_pos];
  let p1=0;let p2;
  for (let i=1;i<full_ar.length;i++){
    //Logger.log('full_ar[i]='+JSON.stringify(full_ar[i]));
    //Logger.log('replace_ar[i]='+JSON.stringify(replace_ar[i]));
    if (full_ar[i][comp_pos] == match_val){
      full_ar.splice(i--,1);
      p1++;
    }
  }
  Logger.log('match_val='+match_val+'deleted='+p1+ ' added='+replace_ar.length);
  
  return full_ar.concat(replace_ar);
}

function getEditUrls(sheet,form) {
  var data = sheet.getDataRange().getValues();
  var responses = form.getResponses();
  var timestamps = [], urls = [], resultUrls = [];
  
  for (var i = 0; i < responses.length; i++) {
    timestamps.push(responses[i].getTimestamp().setMilliseconds(0));
    urls.push(responses[i].getEditResponseUrl());
  }
  for (var j = 1; j < data.length; j++) {

    resultUrls.push([data[j][0]?urls[timestamps.indexOf(data[j][0].setMilliseconds(0))]:'']);
  }
  return resultUrls;  
}

function getScoresFromFile(ss) {
  let f_sh=ss.getSheets()[0];
  if (f_sh.getLastRow()==1){
    return;
  }
  let abc=f_sh.getRange(2,1,f_sh.getLastRow()-1,3).getDisplayValues();
  let dts=f_sh.getRange(2,1,f_sh.getLastRow()-1,1).getValues();
  gp.quiz_subj_nm=ss.getName().replace(/\([^\(]+\)$/,'');
  let furl=ss.getFormUrl();
  //Logger.log('subj_nm='+gp.quiz_subj_nm + ' furl='+furl);
  let rurls;
  if (furl){
    let form = FormApp.openByUrl(furl);
    rurls=getEditUrls(f_sh,form);
  }
  let rows=[];
  for (let i=0;i<abc.length;i++){
    rows.push([chomp(abc[i][2]), dts[i][0], gp.quiz_subj_nm, abc[i][1], furl?rurls[i] : '', furl ])
  }
  return rows;
}

