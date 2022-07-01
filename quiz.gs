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


function collectQuizResultsMain() {
  collectParams();
  let scoreFiles=getSubFoldersFiles(quiz_dir_id);
  Logger.log('scoreFiles='+scoreFiles.length);
  let scores_ar=getScoresFromFiles(scoreFiles);
  saveQuizData(scores_ar,getQuizSaveSh());
}

function getQuizSaveSh() {
  if (! gp.quiz_result_sh){
    gp.quiz_result_sh=getMaakavSS().getSheetByName('allQuiz');
  }
  //Logger.log('shnm='+gp.quiz_result_sh.getName());
  return gp.quiz_result_sh;
}

function saveQuizData(scores_ar,sh) {
  sh.getRange(2,1,sh.getLastRow(),6).clear();
  if (! scores_ar.length){return}
  sh.getRange(2,1,scores_ar.length,6).setValues(scores_ar);
}

function getScoresFromFiles(scoreFiles) {
  let rows=[];
  for (let i=0;i<scoreFiles.length;i++){
    let file_rows=getScoresFromFile(scoreFiles[i]);
    if (file_rows){
      rows=rows.concat(file_rows);
    }
  }
  return rows;
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
    rows.push([abc[i][2], dts[i][0], gp.quiz_subj_nm, abc[i][1], furl?rurls[i] : '', furl ])
  }
  return rows;
}

