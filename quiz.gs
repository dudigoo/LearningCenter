var quiz_dir_id='1YP5aziOBgpzO1GS35z3yhY2op0qXdz2W';
//var quiz_dir_id='1id3DC527c-svm9HTVREvvkAbPEFQTOAq';

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
/*
function tstCloneResponse(){
  //cloneResponse2('1FAIpQLScTtF5kbW-XS8oQG4_qkloL20cOoQ25gXPr8rX5934gtmjKHg', '2_ABaOnuf2dB7NJT-YvG2DJX5DjSzjmH85tJC0bL4qUKmgRm5I5UUMSelr4kzZQkzfsg');
  cloneResponse2('1yFiPMlSJwZ7G49CAQeqreMoaASTI19OPTzCc8L3CW6A', '2_ABaOnufaJ_HwshPqh0zd2cyKwy5rSKmJGtK_QI7Spj_TpjC6PempEDK57Feaozm3YIsZevE');
  //cloneResponse2('1zJTU0I25oiU9qzUHgLRFn-TYzVH58IWYVoc3dRf65_I', '2_ABaOnues44KnVyiHVe2LZoZH0Men-4dtuBX4OhLQ2WN_yIjbHJiZztXosog2vKtyCBe49_0');
}

function cloneResponse2(form_id, response_id){
  let form = FormApp.openById(form_id);
  formResponse = form.getResponse(response_id);
  let newFormResponse = form.createResponse();
  let itemResponses = formResponse.getItemResponses();
  Logger.log('itemResponses.length='+itemResponses.length);
  for (let j = 0; j < itemResponses.length; j++) {
    let itemResponse = itemResponses[j];
    //Logger.log('item type='+ itemResponse.getItem().getType() );
    newFormResponse.withItemResponse(itemResponse);
  }
  form_res=newFormResponse.submit();
  url=form_res.getEditResponseUrl();
  Logger.log('url='+url);
  return url;
}
*/
function findQuizInvalidNamesMain() {
  collectParams();
  let query2="select A, B, C, D ";
  let quizs=querySheet(query2,gp.maakav_file_id,'allQuiz',1);
  for (let i=0;i<quizs.length;i++){
    if (! getAllPupilsMap()[quizs[i][0]]){
      if (!gp.quiz_find_invalid_name_exam_only || (gp.quiz_find_invalid_name_exam_only && quizs[i][2].match(/מבדק|English/))){
        writeLog('invalid: i='+(i+2)+' name='+quizs[i][0]+' dt='+quizs[i][1]+' subj='+quizs[i][2]);
      }
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
  let changed_in_past_days=20;
  //quiz_dir_id='1IPmlBkL_f5V7RxuewRH7Z50_SpaXLABK';
  let files_ar=[];
  let cur_all_quiz_sh=getMaakavSS().getSheetByName('allQuiz');
  let cur_all_quiz_ar=cur_all_quiz_sh.getRange(2,1,cur_all_quiz_sh.getLastRow()-1,6).getValues();
  cur_all_quiz_ar.sort((a, b) => {
  if (a[2] < b[2]) {
    return -1;
  }
  if (a[2] > b[2]) {
    return 1;
  }
    return 0;
  });
  //getFolderIdFilesRecursivly(quiz_dir_id,'application/vnd.google-apps.spreadsheet' , files_ar, 100); // sheet last modofied not changed on submit so using form files 
  getFolderIdFilesRecursivly(quiz_dir_id,'application/vnd.google-apps.form' , files_ar, 100);
  //Logger.log('files_ar '+files_ar);
  let dt= new Date();
  dt.setHours(dt.getHours()-26*changed_in_past_days);
  Logger.log('ignore files modified earlier then '+dt);
  for (let i=0;i<files_ar.length;i++){
    if (files_ar[i].getLastUpdated()<dt){
      continue;
    }
    //xlet scores_ar=getScoresFromFile(SpreadsheetApp.open([i]));
    let scores_ar=getScoresFromFile(SpreadsheetApp.openById(FormApp.openById(files_ar[i].getId()).getDestinationId()));
    if (! scores_ar){
      continue;
    }    
    Logger.log('fixing file '+files_ar[i].getName());
    //Logger.log('fixing file '+files_ar[i].getName()+ ' scores_ar='+JSON.stringify(scores_ar));
    cur_all_quiz_ar=replaceOldScoresWithNewScores(cur_all_quiz_ar,scores_ar,2);
    Logger.log('post replaceOldScoresWithNewScores cur_all_quiz_ar.length='+cur_all_quiz_ar.length);
  }
  Logger.log('cur_all_quiz_ar.length='+cur_all_quiz_ar.length+' cur_all_quiz_sh.getLastRow()='+cur_all_quiz_sh.getLastRow());
  for (let i=cur_all_quiz_ar.length; i<(cur_all_quiz_sh.getLastRow() -1); i++){
    cur_all_quiz_ar.push(['','','','','','']);
    Logger.log('adding empty row');
  }
  cur_all_quiz_sh.getRange(2,1,cur_all_quiz_ar.length,cur_all_quiz_ar[0].length).setValues(cur_all_quiz_ar);
}

function replaceOldScoresWithNewScores(full_ar, replace_ar, comp_pos) {
  //Logger.log('comp_pos='+comp_pos+' replace_ar.length='+replace_ar.length);
  let match_val=replace_ar[0][comp_pos];
  let group_first_position=-1;let group_last_position=0;
  for (let i=0;i<full_ar.length;i++){
    //Logger.log('full_ar[i]='+JSON.stringify(full_ar[i]));
    //Logger.log('replace_ar[i]='+JSON.stringify(replace_ar[i]));
    
    if (full_ar[i][comp_pos] == match_val){
      if (group_first_position== -1) {
        group_first_position=i;
      }
    } else {
      if (group_first_position != -1 && ! group_last_position){
        group_last_position=i-1;
        break;
      }
    }
  }
  Logger.log('group_first_position='+group_first_position+' group_last_position='+ group_last_position);
  if (! group_last_position){
    group_last_position=full_ar.length-1
  }
  let new_ar;
  Logger.log('match_val='+match_val+'deleted='+group_first_position+ ' to:'+group_last_position+ ' added='+replace_ar.length);
  if (group_first_position==0){
    Logger.log('group_first_position=zero group_last_position='+ group_last_position);
    new_ar=replace_ar.concat(full_ar.slice(group_last_position+1));
  } else if (group_last_position==(full_ar.length-1)) { 
    new_ar=full_ar.slice(0,group_first_position).concat(replace_ar);
    Logger.log('group_last_position=full_ar.length group_last_position='+group_last_position);
  } else { // middle
    Logger.log('middle  group_first_position='+group_first_position+' group_last_position='+group_last_position);
    new_ar=full_ar.slice(0,group_first_position).concat(replace_ar).concat(full_ar.slice(group_last_position+1));
  }
  return new_ar;
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
  let regex=/\/([^\/]+)\/viewform$/;
  const found = furl.match(regex);
  Logger.log('subj_nm='+gp.quiz_subj_nm + ' furl='+furl);
  let rurls;
  if (furl){
    let form = FormApp.openByUrl(furl);
    rurls=getEditUrls(f_sh,form);
  }
  let rows=[];
  for (let i=0;i<abc.length;i++){
    //rows.push([chomp(abc[i][2]), dts[i][0], gp.quiz_subj_nm, abc[i][1], furl?rurls[i] : '', furl ])
    rows.push([chomp(abc[i][2]), dts[i][0], gp.quiz_subj_nm, abc[i][1], furl?rurls[i] : '' , furl?found[1]:''])
  }
  return rows;
}

