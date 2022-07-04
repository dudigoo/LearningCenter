
function getFmtDtStr(dt) {
  return dt.toLocaleDateString(gp.locale,{year:"2-digit",month:"2-digit", day:"2-digit"}).replace(/\./g, '/').replace(/0(\d)/g, "$1");
}

function getYMDStr(d) {
  const offset = d.getTimezoneOffset();
  let sd = new Date(d.getTime() - (offset*60*1000));
  return sd.toISOString().split('T')[0];
}

function querySheet(query,fid,shname,headers){
  /*if (!gp.query_sh_ids) {
    gp.query_sh_ids={};
  }
  if (!gp.query_ss_ids) {
    gp.query_ss_ids={};
  }
  if (!gp.query_ss_ids[fid]) {
    gp.query_ss_ids[fid]=SpreadsheetApp.openById(fid);
  }
  if (!gp.query_sh_ids[fid+shname]) {
    gp.query_sh_ids[fid+shname]=gp.query_ss_ids[fid].getSheetByName(shname).getSheetId();
  }*/
  
  let hdrs= (headers == undefined) ? 1 : headers;
  let url = "https://docs.google.com/spreadsheets/d/" + fid + "/gviz/tq?sheet=" + shname + "&headers="+hdrs+"&tqx=out:csv&tq=" + encodeURIComponent(query);
  Logger.log('query='+query +' shname='+shname +' hdrs='+hdrs);
  var res = UrlFetchApp.fetch(url, {headers: {Authorization: "Bearer " + ScriptApp.getOAuthToken()}});
  Logger.log('url='+url );
  //Logger.log('res ='+res );
  let vals;
  try {
    vals = Utilities.parseCsv(res.getContentText());
  } catch (error) {
    Logger.log('querySheet error='+error );
    return;
  }
  //Logger.log('qSheet vals='+JSON.stringify(vals) );
  if (hdrs != 0){
    vals.splice(0,hdrs);
    //vals.shift();
  }
  return vals;
}

function confirm_popup(s1,s2) {
  var ui = SpreadsheetApp.getUi(); // Same variations.
  s1 = s1 ? s1 : 'Please confirm';
  s2 = s2 ? s2 : 'Are you sure you want to continue?';
  var result = ui.alert(s1, s2, ui.ButtonSet.YES_NO);
  Logger.log('r='+result);
  return result;
}

function getLogsh() {
  if (! gp.scripts_log_sh){
    gp.scripts_log_sh = SpreadsheetApp.openById(gp.ms_container_id).getSheetByName('log');
    gp.scripts_log_sh.getRange(1,3,gp.scripts_log_sh.getLastRow(),1).clear();
  }
  return gp.scripts_log_sh;
}

function writeLog(msg) {
  var str = 'C' + log_row;
  Logger.log('row=' + str + ' msg=' + msg);
  getLogsh().getRange(str).setValue(msg);  
  log_row=log_row+1;
}

function checkLog(action,subj,to){
  writeLog('End');
  if (! subj){
    subj='Shibutz errors';
  }
  if (! to){
    to='mlemida.ryam@gmail.com';
  } else {
    to=to+',mlemida.ryam@gmail.com'
  }
  if (gp.scripts_log_sh.getRange('C4').getValue() != 'End') {
    Logger.log('has errs' );
    if (action == 'mail'){
      var r=gp.scripts_log_sh.getRange(3, 3, gp.scripts_log_sh.getLastRow()-2, 1);
      var me=r.getValues().join("\n");
      MailApp.sendEmail(to, subj,  me);
      Logger.log('Sent mail=' + me );
    } else {
      Logger.log('setActiveSheet log' );
      var ss=SpreadsheetApp.getActive();
      ss.setActiveSheet(ss.getSheetByName('log'));
    }
  }
}

function getScriptGlobalParams(col){
  let sel = (col==3)? 'C' : 'B';
  let q='select '+ sel;
  gp.ms_container_id = PropertiesService.getScriptProperties().getProperty('container_id');
  let par=querySheet(q, gp.ms_container_id, 'manage', 1);
  //let params=par.map(e => e[0]);
  return par;
}

function onOpen(){
  var id= SpreadsheetApp.getActiveSpreadsheet().getId();
  PropertiesService.getScriptProperties().setProperty('container_id',id);
  Logger.log('onOpen container_id='+id);
}

function collectParams(col) {
  if (! col){col=2};
  //Logger.log('container_id='+gp.ms_container_id);
  let params = getScriptGlobalParams(col);
  gp.heb_year = params[0][0];
  gp.g_month_name = params[1][0];
  gp.monthly_thin = params[2][0];
  let v=params[3][0]; //stu
  if (v != '') {
    wfolders.push([v, wtyp_s]);
  }
  wtyp2fol_id[wtps] = v;
  v=params[4][0]; // morim
  if (v != '') {
    wfolders.push([v, wtyp_m]);
  }
  wtyp2fol_id[wtpm] = v;
  v=params[5][0]; //hanichim
  if (v != '') {
    wfolders.push([v, wtyp_h]);
  }
  wtyp2fol_id[wtph] = v;

  //Logger.log('work folders=' + wfolders);
  gp.zminut_file_id = params[6][0];
  gp.shib_dates = params[7][0];
  gp.sheet2show = params[8][0];
  gp.out_folder_id = params[9][0];
  gp.shibutz_file_id = params[10][0];
  gp.maakav_file_id = params[11][0];
  gp.wrkrs_ss_id = params[12][0];

  gp.rikuz_grade_filter = params[13][0];
  if (gp.rikuz_grade_filter){
    gp.rikuz_grade_filter_ar=gp.rikuz_grade_filter.split(',');
  }
  gp.top_accounting_dir_id = params[14][0];
  //Logger.log('A: g_top_accounting_dir_id='+g_top_accounting_dir_id );
  gp.rikuz_file_id = params[15][0];
  gp.rikuz_wrkrs_filter = params[16][0];
  gp.rikuz_wrkrs_filter_ar=gp.rikuz_wrkrs_filter.split(',');
  gp.g_month_name_ar=gp.g_month_name.split(',');
  //gp.hours_master_id = params[17][0];
  gp.rikuz_subjects = params[17][0];
  //gp.rikuz_subjects = params[27][0];
  if (gp.rikuz_subjects){
    gp.rikuz_subjects_ar=gp.rikuz_subjects.split(',');
  }
  gp.wrkrs_row_str = params[19][0];

  gp.pupil_alfon_id = params[20][0];
  gp.nizul_src = params[21][0];
  gp.nizul_tgt = params[22][0];
  gp.shibutz_tmplts = params[23][0];

  gp.rikuz_subjects_omit = params[18][0];
  gp.shib_days_cycle = params[24][0];
  gp.shibutz_mail_to = params[25][0];
  gp.mashov_scores_dir_id = params[26][0];
  gp.hours_master_id = params[27][0];
  //Logger.log('pms='+params[29]);
  //gp.ab_last_dt = params[29][0];
  //gp.dates_dmy_fmt = (gp.scripts_ss.getSpreadsheetLocale() == 'iw_IL') ? 'y' : '';
 
}

function chomp(raw_text){
  raw_text=raw_text.replace(/ *, */g, ',').replace(/,,/g, ',').replace(/  /g, ' ').replace(/[, ]+$/, '');
  return raw_text.replace(/^\s+/, '');
}

function dropColumns(ar,cols) {
  for (let k=cols.length-1;k>=0;k--){
    for (let i=0;i<ar.length;i++){
      ar[i].splice(cols[k],1);
    }
  }
}

function findEmptyColumns(ar,from_row) {
  let max=0;
  //Logger.log('femptycols ar='+JSON.stringify(ar));
  for (let i=0;i<ar.length;i++){
    //Logger.log('i='+i);
    if (ar[i].length>max){
      max=ar[i].length;
    }
  }
  emp_col=[];
  for (let j=0;j<max;j++){
    trim=1;
    for (let i=from_row;i<ar.length;i++){
      //Logger.log('j='+j+ ' i='+i+ ' ar[i][j]='+ar[i][j]);
      if (ar[i][j]) {
        trim=0;
        //Logger.log('nemp j='+j+ ' i='+i);
        break;
      }
    }
    if (trim){
      emp_col.push(j);
    }
  }
  return emp_col;
}