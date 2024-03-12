var empty_rows; 
var prev={};
var g_maakav_sh;
var g_present_date; var g_min_dt;
var g_dup_hour={};

function cpMain() {
  collectParams(3);
  cpInit();
  let files_ar=[];
  getFilesFromFoldersRecurse(files_ar, gp.w_folders_id_a, 'application/vnd.google-apps.spreadsheet', 1, 500); //125
  Logger.log('files to cp2maakav='+files_ar.length);
  for (let i=0;i<files_ar.length;i++){
    let tnm = files_ar[i].getName();
    Logger.log(' tnm='+tnm);
    let ss = SpreadsheetApp.open(files_ar[i]);
    Logger.log('ss name='+ss.getName());
    let tabnm=tnm.replace(/^\S+ \S+ \S+ /, '');
    let w=getWorkerByName(tabnm);
    cp2maakav(files_ar[i],ss,w);
  }
  Logger.log('logDailyHours start');
  logDailyHours();
  Logger.log('logDailyHours end');
  //compareWorkersDailyHoursSheets(); we report once per month to atid so this function not needed now
  Logger.log('compareWorkersDailyHoursSheets end');
  mailLog('hreports2maakav')
}


function compareWorkersDailyHoursSheetToAtidSheetMain() {
  collectParams();
  compareWorkersDailyHoursSheets('atid');
}

function compareWorkersDailyHoursSheetsTst() {
  collectParams(3);
  compareWorkersDailyHoursSheets();
}

function getAtidDoneItem(worker, date) {
  if (!gp.loaded_atid_done_hours_sheet){
    gp.loaded_atid_done_hours_sheet= loadAtidDoneHoursSheet();
  }
  //Logger.log('w0rk='+worker+' date='+date);
  if ( gp.loaded_atid_done_hours_sheet[worker]){
    //Logger.log('found wrk');
  } else { return ''}
  if (  date in gp.loaded_atid_done_hours_sheet[worker]){
    //Logger.log('found dt');
  }
  if (! gp.loaded_atid_done_hours_sheet[worker] || (! (date in gp.loaded_atid_done_hours_sheet[worker]))){
    return '';
  }
  Logger.log('wrk='+worker+' date='+date);
  Logger.log('gp.loaded_atid_done_hours_sheet[worker][date]='+gp.loaded_atid_done_hours_sheet[worker][date]);
  return gp.loaded_atid_done_hours_sheet[worker][date];
}

function loadAtidDoneHoursSheet() {
  getAllWorkers();
  //Logger.log('all_wrkrs_by_name2='+JSON.stringify(gp.all_wrkrs_by_name2));
  let folder = DriveApp.getFolderById(gp.atid_work_hours_dir_id);
  let files = folder.getFiles();
  let ss_file;
  let file_id;
  while (files.hasNext()) {
    let file = files.next();
    if (file.getMimeType() === 'application/vnd.google-apps.spreadsheet') {
      ss_file=file;
      file_id=ss_file.getId();
      break;
    }
  }
  if (! ss_file){ 
    ss_file=convertXlsx2sheets(folder)[0];
    file_id=ss_file.id;
    Logger.log('ss_file id='+ss_file.id);
  }
  
  let atid_ss = SpreadsheetApp.openById(file_id);
  atid_ss.setSpreadsheetTimeZone("Asia/Jerusalem");
  Logger.log('file name='+atid_ss.getName());
  let atid_sh = atid_ss.getSheets()[0];
  let atid_ar=atid_sh.getDataRange().getValues();
  let atid_loaded_sheet={};
  atid_ar.forEach((el, inx) => {
    if (! inx || ! el[3]) {return}
    Logger.log('inx='+inx+' el3='+el[3]);
    //Logger.log(' nm02='+gp.all_wrkrs_by_name2[el[3]]);
    //Logger.log(' nm2='+gp.all_wrkrs_by_name2[el[3]]['name']);
    if (! atid_loaded_sheet[gp.all_wrkrs_by_name2[el[3]]['name']]){ 
      atid_loaded_sheet[gp.all_wrkrs_by_name2[el[3]]['name']]={};
    }
    //Logger.log('el[8]='+el[8]);
    if (! (el[5] in atid_loaded_sheet[gp.all_wrkrs_by_name2[el[3]]['name']])){
      atid_loaded_sheet[gp.all_wrkrs_by_name2[el[3]]['name']][el[5]]=0;
      //Logger.log('aded key='+el[5]);
    }
    atid_loaded_sheet[gp.all_wrkrs_by_name2[el[3]]['name']][el[5]] += el[8];
  });
  //Logger.log('loaded atid file'+ JSON.stringify(atid_loaded_sheet));
  return atid_loaded_sheet;
}

function compareWorkersDailyHoursSheets(atid_type) {
  let ss=SpreadsheetApp.openById(gp.workers_daily_hours_ss_id);
  let list_sh=ss.getSheetByName('list');
  let list_ar=list_sh.getRange(1,1,list_sh.getLastRow(),list_sh.getLastColumn()).getValues();
  let done_sh; let done_ar; let adaptors;
  if (! atid_type){
    done_sh=ss.getSheetByName('done');
    //Logger.log('did getSheetByName');
    done_ar=done_sh.getRange(1,1,done_sh.getLastRow(),done_sh.getLastColumn()).getValues();
    adaptors=getDoneCorrectInx(list_ar,done_ar);
    Logger.log('adopters0='+JSON.stringify(adaptors[0]))
    Logger.log('adopters1='+JSON.stringify(adaptors[1]))
  }
  for (let i=1;i<list_ar.length;i++){
    for (let j=1;j<list_ar[0].length;j++){
      //Logger.log("i="+i+' j='+j + ' adaptors[1][i]='+adaptors[1][i]+' adaptors[0][j]='+adaptors[0][j])
      let done_item='';
      if (! atid_type){
        if (adaptors[1][i] && adaptors[0][j]){
          done_item=done_ar[adaptors[1][i]][adaptors[0][j]];
          //Logger.log('new done_item='+done_item);
        }
      } else {
        done_item=getAtidDoneItem(list_ar[i][0], list_ar[0][j]);
        Logger.log('2 atid list_ar[i][0]='+list_ar[i][0]+' list_ar[0][j]'+list_ar[0][j]);
        Logger.log('atid done_item='+done_item);
      }
      let new_val=compareListAndDoneCell(list_ar[i][j], done_item);
      //Logger.log('new_val='+new_val + ' i='+i+' j='+j);
      if (new_val) {
        list_ar[i][j]=new_val;
      }
    }
  }
  updateListSheet(list_ar);
}

function updateListSheet(list_ar) {
  let ss=SpreadsheetApp.openById(gp.workers_daily_hours_ss_id);
  ss.getSheetByName('list').getRange(1,1,list_ar.length,list_ar[0].length).setValues(list_ar);

}

function compareListAndDoneCell(list_item,done_item) {
  if (! list_item && ! done_item){
    return
  }
  if (!list_item){list_item=''}
  let list_item_orig=list_item.toString().replace(/#.*$/,'');
  //Logger.log('list_item_orig='+list_item_orig+ ' done_item='+done_item);
  let done_item_orig=done_item.toString().replace(/#.*$/,'');
  //Logger.log(' list_item_orig='+list_item_orig+' done_item_orig='+done_item_orig);
  if ( done_item_orig != list_item_orig) {
    done_item_orig = done_item_orig ? done_item_orig : 0;
    list_item_orig = list_item_orig ? list_item_orig : 0;
    let ret_val=list_item_orig+'#'+(Number(list_item_orig) - Number(done_item_orig));
    //Logger.log(' ret_val='+ret_val);
    return ret_val;
  }
  return list_item_orig;
}

function getDoneCorrectInx(list_ar,done_ar) {
  let done_date_adaptor=[];
  let j=1;
  //Logger.log('list_ar='+JSON.stringify(list_ar));
  for (let i=1;i<list_ar[0].length;i++){
    for (;j<done_ar[0].length; j++){
      //Logger.log('i='+i+' j='+j+ ' done_ar[0][j]='+done_ar[0][j]);
      if (list_ar[0][i].getTime() == done_ar[0][j].getTime()) {
        done_date_adaptor[i]=j
        break;
      } else if (list_ar[0][i].getTime() < done_ar[0][j].getTime()) {
        j--;
        break;
      }
    }
  }
  let done_nm_adaptor=[];
  j=1;
  for (let i=1;i<list_ar.length;i++){
    for (;j<done_ar.length; j++){
      if (list_ar[i][0] == done_ar[j][0]) {
        done_nm_adaptor[i]=j
        break;
      }
      if (list_ar[i][0] < done_ar[j][0]) {
        j--;
        break;
      }
    }
  }  
  return [done_date_adaptor, done_nm_adaptor];
}

function convertObjectTo2DArray(obj, all_dates) {
    let people = Object.keys(obj).sort();
    //Logger.log('all_dates='+JSON.stringify(all_dates));
    let dates = Object.keys(all_dates).map((e) => new Date(e));
    //Logger.log('dates='+JSON.stringify(dates));
    dates.sort((a, b) => { 
      if (a.getTime()>b.getTime()) {return 1}
      if (a.getTime()<b.getTime()) {return -1}
      return 0;
    })
    let array = [[''].concat(dates)];

    for (let person of people) {
        let row = [person];
        for (let date of dates) {
            row.push(obj[person][date] || '');
        }
        array.push(row);
    }

    return array;

  }



function logDailyHours() {
  //Logger.log('obj='+JSON.stringify(gp.workers_days_hours));
  let arr = convertObjectTo2DArray(gp.workers_days_hours, gp.workers_dates);
  //Logger.log('arr='+JSON.stringify(arr));
  let ss=SpreadsheetApp.openById(gp.workers_daily_hours_ss_id);
  let sh=ss.getSheetByName('list');
  if (!sh){
    sh= ss.insertSheet();
    sh.setName('list');
    ss.moveActiveSheet(0);
    sh.insertColumns(10, 40); 
  }
  sh.clear();
  rows=arr.length;
  cols=arr[0].length;
  //Logger.log('rows='+rows+' cols='+cols);
  sh.getRange(1,1,rows,cols).setValues(arr);
  
  //  total
  let sh_tot=ss.getSheetByName('listTotal');
  sh_tot.getRange("A1:CW250").clearContent();
  sh_tot.getRange(1,1,rows,cols).setValues(arr);
}

function cpInit() {
  g_present_date = new Date();
  let g_month_name_frmonth=gp.g_month_name.replace(/^\d+\./,'').replace(/-.*$/,'') - 1;
  g_min_dt = new Date();
  g_min_dt.setMonth(g_month_name_frmonth);
  g_min_dt.setDate(10);
  g_min_dt.setFullYear(g_present_date.getFullYear());
  //g_min_dt.setHours(0,0,0,0);
  g_min_dt.setMonth(g_min_dt.getMonth()-1);
  if (g_present_date.getMonth() <  g_month_name_frmonth){
    Logger.log(' year changed. g_present_date.getMonth()='+g_present_date.getMonth()+ ' g_month_name_frmonth='+g_month_name_frmonth) ;
    g_min_dt.setFullYear(g_present_date.getFullYear()-1);
  }
  //Logger.log(' g_month_name='+gp.g_month_name+' a='+gp.g_month_name.replace(/^.../,'')+' b='+gp.g_month_name.replace(/^.../,'').replace('-.*$',''));
  Logger.log(' g_month_name_frmonth='+g_month_name_frmonth);
  Logger.log(' g_min_dt='+g_min_dt);
  Logger.log('g_present_date='+g_present_date);

  g_maakav_sh=getMaakavSS().getSheetByName('all');
  gp.mail2admin_ar=[];
  gp.mail2educator_ar=[];
  gp.workers_days_hours={};
  gp.workers_dates={};
}

function getMaakavSS() {
  if (! gp.maakav_ss){
    gp.maakav_ss=SpreadsheetApp.openById(gp.maakav_file_id);
  }
  return gp.maakav_ss;
}

function cp2maakav(file,ss,w) {
  //Logger.log(' w='+JSON.stringify(w)+' file='+file.getName()) ;
  w.worker_hours_url=file.getUrl();
  Logger.log(' person='+w.name + ' shnm='+gp.g_month_name) ;
  var sheet = ss.getSheetByName(gp.g_month_name);
  var werrs=[];
  g_dup_hour={};
  Logger.log(' snm='+sheet.getName());
  var rowsnum=sheet.getMaxRows() - 7;
  let sh_ar=sheet.getRange(8,1,rowsnum-7,sheet.getMaxColumns()).getValues();
  let copied_ar = sh_ar.map(x => [x[18]]);
  let copied_pre = copied_ar.join();

  prev.date='';  prev.hrs=''; prev.frtm='';prev.totm='';prev.dow='';
  if (! handleWorkerRows(w,werrs,sh_ar,copied_ar)){
  //Logger.log('cp2makav werrs.length='+werrs.length);
    updateCopiedCol(sheet,copied_pre,copied_ar);
  }
  handleWorkerErrs(werrs,w);
  addSomeRows(sheet,w.name);
}

function handleWorkerRows(w,werrs,sh_ar,copied_ar) { 
  empty_rows=0;
  //Logger.log('sh_ar.length='+sh_ar.length+' sh_ar[0].length='+sh_ar[0].length);
  let wrkr_rows2write=[];  
  for (var i=0; i<sh_ar.length;i++){// w rows
    var lnerrs=[];
    cpRowInfo(w,i,lnerrs,sh_ar[i],copied_ar,wrkr_rows2write);
    //if (lnerrs.length) {Logger.log('ln='+i+' errs='+lnerrs.join("\n"));}
    lnerrs.forEach(el => {werrs.push(el)});
    //Logger.log('werrs.length='+werrs.length+' lnerrs.length='+lnerrs.length);
    if (empty_rows>5){
      break;
    }
  }
  return appendRows2Maakav(wrkr_rows2write);
}
  
function updateCopiedCol(sheet,copied_pre,copied_ar) {
  let copied_post = copied_ar.join();
  if (copied_pre != copied_post){
    sheet.getRange(8,19,copied_ar.length,1).setValues(copied_ar);
  }
}

function handleWorkerErrs(werrs,w) {
  //Logger.log('werrs.length='+werrs.length);
  if (werrs.length){
    var perrs=["שלום "+ w.name.split(' ')[0]+",\n"].concat(werrs);
    perrs.push('<a href="https://tinyurl.com/ya7ptvoq">הסבר לדיווח שעות</a>');
    perrs.push('<a href="'+w.worker_hours_url+'">דיווח שעות שלי</a>');
    var msg="<p dir=RTL>"+ perrs.join("<br>") + '</p>';
    MailApp.sendEmail(w.mail,'בבקשה לתקן בדיווח שעות', msg, {htmlBody: msg});
    //MailApp.sendEmail('mlemida.ryam@gmail.com','בבקשה לתקן בדיווח שעות', msg, {htmlBody: msg});
    //Logger.log('w='+JSON.stringify(w));
    //Logger.log('errs='+perrs.join("\n"));
  }
}

function cpRowInfo(wrkr,rn,lnerrs, wrow, copied_ar, wrkr_rows2write) {
  var vals={};
  wsr2vals(vals,wrow,wrkr.name,rn);
  //Logger.log('dt='+vals.date+' vals0='+JSON.stringify(vals) );
  //Logger.log('rn='+rn+' empty_rows='+empty_rows );
  empty_rows= (! vals.hours && ! vals.level && ! vals.subj) ? empty_rows+1 : 0;

  let hours= vals.hours ? vals.hours : 0;
  prev2vals(vals,prev);
  if (hours){
    addHoursToWorkerDayHours(wrkr.name,vals,hours); // collect hours
  }
  if (chekDupHours(hours,vals,lnerrs)){
    return;
  }
  //Logger.log('pupils='+vals.pupils+' dt='+vals.date+' vals='+JSON.stringify(vals) );
  vals2prev(vals,prev);
  if (! vals.pupils || vals.pupils.substring(0,1) == '-'|| copied_ar[rn]=='y' || vals.subj.substring(0,1) == '-'){
    return;
  }
  notifyEducator(vals,wrkr);
  notifyManager(vals,wrkr);
  findWrkrErrs(vals, lnerrs);
  var kids = vals.pupils.split(",");
  var kidsu = kids.filter(onlyUnique);
  if (kidsu.length != kids.length){ lnerrs.push('שורה '+vals.row+': תלמיד כפול')}
  let rownum_pre=wrkr_rows2write.length;
  for (var i=0;i<kidsu.length;i++){
    let rcd=cpKidInfo(vals,kidsu[i],lnerrs,wrkr_rows2write);
    if (rcd){
      return rcd;
    }
  }
  if (! lnerrs.length && rownum_pre < wrkr_rows2write.length){
    copied_ar[rn][0]='y';
    //Logger.log('copied row='+rn );
  }
}

function addHoursToWorkerDayHours(name,vals,hours){
  if (! (name in gp.workers_days_hours)){
    gp.workers_days_hours[name]={};
  }
  if (! (vals.date in gp.workers_days_hours[name])){
    gp.workers_days_hours[name][vals.date]=0;
  }  
  gp.workers_days_hours[name][vals.date]+=hours;
  gp.workers_dates[vals.date]=1;//collect all dates to sort later
}

function notifyEducator(vals,wrkr){
  //Logger.log('notifyEducator wrkr='+JSON.stringify(wrkr) );
  //Logger.log('vals='+JSON.stringify(vals) );
  if (! vals.note.toString().match(/##/)) {
    return;
  }
  let ps=vals.pupils.split(',');
  for (let i=0;i<ps.length;i++){
    let ed=getPupilEducator(ps[i], vals.level);
    let msg='שלום '+ (ed ? ed[0][2].split(' ')[0] :'') + ',<br><br>';
    msg += 'התלמיד:'+ps[i]+'<br>';
    msg += 'הערה:'+vals.note.replace('##','')+'<br>';
    msg += 'אשמח לדבר על כך'+'<br><br>';
    msg += 'מורה המרכז:'+vals.wrkr+'<br>';
    msg += 'דואל:'+wrkr.mail+'<br>';
    msg += 'טלפון:'+wrkr.phone+'<br>';
    msg += 'תאריך שיעור:'+getFmtDtStr(vals.date)+'<br>';
    msg += 'מקצוע :'+vals.subj+'<br>';
    let em={
        to: ed ? ed[0][3]:'',
        cc: gp.shibutz_mail_to+','+wrkr.mail,
        subject: ps[i]+' - '+ 'דיווח מרכז למידה',
        htmlBody: "<p dir=RTL>"+msg+'</p>'
    }
    gp.mail2educator_ar.push(em);
  }     
}

function getPupilEducator(pupil,level){
  //Logger.log('level='+level + 'pupil='+pupil);
  //Logger.log('getAllPupilsMap()[pupil]='+getAllPupilsMap()[pupil]);
  if (! getAllPupilsMap()[pupil]){
    writeLog('pupil not found:'+pupil);
    return;
  }
  let pclass=getAllPupilsMap()[pupil][8];
  let q='select A,B,C,D where A="'+level+'" and B='+pclass;
  let edu=querySheet(q, gp.wrkrs_ss_id, 'מחנכים');
  Logger.log('edu='+edu);
  return edu;
}

function notifyManager(vals,wrkr){
  if (! vals.note.toString().match(/@/)) {
    return;
  }
  let a='teacher:'+vals.wrkr;
  let b='row:'+vals.row;  
  let c='note:'+vals.note;
  let d='pupil:'+vals.pupils;
  let e='class:'+vals.level;
  let f='<a href="'+wrkr.worker_hours_url+'">hour report</a>';
  gp.mail2admin_ar.push([a,b,c,d,e,f]);
  //Logger.log('mail2admin_ar='+gp.mail2admin_ar);
}

function chekDupHours(hours,vals,lnerrs){
  if (hours){
    if (g_dup_hour[vals.frtm + ' '+ vals.date]){ 
      lnerrs.push('שורה '+vals.row+': טור שעות כבר מולא ליום ושעה זו');
      return 1;
    } else {
      g_dup_hour[vals.frtm + ' '+vals.date] = 1;
    }
  }
  return;
}

function findWrkrErrs(vals, lnerrs) {
  if (vals.props[0] != 'לא הגיע' && vals.impression.length<1 &&  ! ['י','יא','יב'].includes(vals.level)){ lnerrs.push('שורה '+vals.row+': חסר ש.ב.')}
  if (vals.row==8 && vals.date==''){ lnerrs.push('שורה '+vals.row+': חסר תאריך')}
  let dt= new Date(vals.date);
  let day = dowmap[dt.getDay().toString()];
  //Logger.log('day='+day +' vals.dow='+vals.dow );
  //Logger.log('dt='+dt.toString() );
  if (dt.toString()=='Invalid Date'){
    //Logger.log('invalid' );
    lnerrs.push('שורה '+vals.row+': תאריך לא חוקי')
  } else {
    //Logger.log('vali' );
    if (g_present_date.getTime() < dt.getTime()) { lnerrs.push('שורה '+vals.row+': תאריך בעתיד')}
    if (g_min_dt.getTime() > dt.getTime() && vals.note.substring(0,1) != '-' ) { 
      lnerrs.push('שורה '+vals.row+':  תאריך ישן מדי - לאישור הכנס - בתחילת טור פעילות')
    }
    if(day != vals.dow){
      //Logger.log('got the err' );
      lnerrs.push('שורה '+vals.row+':  תאריך לא מתאים ליום בשבוע')
    }
  }
  if (vals.note.length<1 && vals.props[0] != 'לא הגיע'){ lnerrs.push('שורה '+vals.row+': חסרה פעילות')}
  if (vals.level.length<1){ lnerrs.push('שורה '+vals.row+': חסרה שכבה')}
  if (vals.subj.length<1){ lnerrs.push('שורה '+vals.row+': חסר מקצוע')}
}

function onlyUnique(value, index, self) {
  return self.indexOf(value) === index;
}

function cpKidInfo(vals,kid,errs,wrows2add) {
  if (! kid){
    errs.push('שורה '+vals.row +':  שם תלמיד ריק '+ kid);
    return;
  }
  let kid_a=getAllPupilsMap()[kid];
  //Logger.log('cpKidInfo kid_a='+kid_a );  
  if (!kid_a){
    errs.push('שורה '+vals.row +':  שם תלמיד שגוי '+ kid);
    //writeLog('invalid kid. name:'+ kid +': teacher:'+vals.wrkr+ ' row='+vals.row);
  } 
  if (errs.length>0){
    return;
  }
  vals.group=kid_a[3];
  //Logger.log('kid_a='+kid_a ); 
  let rcd=pushVals2add(kid,vals,wrows2add);;
  return rcd;
}

function pushVals2add(kid,vals,wrows2add){
  //Logger.log('adding dt='+vals.date+' vals='+ JSON.stringify(vals));
  let valsar=[vals.date, vals.subj, vals.note, vals.impression, vals.wrkr, 
            kid, vals.level, vals.group, '=ROW()', vals.hours];
  //valsar= valsar.concat(vals.props);
  valsar= valsar.concat(vals.props.slice(0,3).concat([vals.frtm,vals.totm])); // frtmtotm
  if (! valsar[0]){
    writeLog('empty date: valsar='+JSON.stringify(valsar)+' vals='+JSON.stringify(vals));
    return 'empty date';
  }
  wrows2add.push(valsar);
}

function mailLog(subj){
  if (gp.mail2admin_ar.length){
    let x1=gp.mail2admin_ar.map(e => e.join('<br>'));
    //Logger.log('x1='+x1);
    let msg=x1.join('<br><br>');
    //Logger.log('msg='+msg);
    //MailApp.sendEmail('dudigoo@gmail.com','Teachers @ report', msg, {htmlBody: msg});
    MailApp.sendEmail(gp.shibutz_mail_to,'Teachers @ report', msg, {htmlBody: msg});
  }
  if (gp.mail2educator_ar.length){
    for (let i=0; i<gp.mail2educator_ar.length; i++){
      MailApp.sendEmail(gp.mail2educator_ar[i]);
    }
  }
  checkLog('mail',subj);
  return;
  var rows_added=g_maakav_sh.getLastRow()- last_row_num_at_start;
  if (rows_added) {
    //var r=g_maakav_sh.getRange(last_row_num_at_start + 1, 1, rows_added, 8);
    //var me=r.getValues().join("\n");
    //MailApp.sendEmail("dudigoo@gmail.com",'newly added maakav rows',  'rows added:'+rows_added);
    //Logger.log('sent new maakav rows mail');
  }
}

function wsr2vals(vals,r,wrkr,rownum) {
  //Logger.log('wsr2vals r='+r);
  vals.dow=r[0];
  vals.date=r[1];
  vals.frtm=r[2];
  vals.totm=r[3];
  vals.hours=r[4];
  vals.level=r[5];
  vals.subj=r[6];
  vals.pupils=chomp(r[7]);
  vals.wrkr=wrkr;
  vals.row=rownum+8;
  vals.impression=r[9];
  vals.note=r[8];
  vals.props=r.slice(10,18);
  //Logger.log('wsr2vals dt='+vals.date+ ' vls='+JSON.stringify(vals));
}

function vals2prev(vals,prev){
  if (vals.dow) {prev.dow=vals.dow}
  if (vals.date) {prev.date=vals.date}
  if (vals.frtm) {prev.frtm=vals.frtm}
  if (vals.totm) {prev.totm=vals.totm}
  if (vals.hours) {prev.hours=vals.hours}
  if (vals.level) {prev.level=vals.level}
  if (vals.subj) {prev.subj=vals.subj}
}

function prev2vals(vals,prev){
  if (! vals.dow && prev.dow) {vals.dow=prev.dow}
  if (! vals.date && prev.date) {vals.date=prev.date}
  if (! vals.frtm && prev.frtm) {vals.frtm=prev.frtm}
  if (! vals.totm && prev.totm) {vals.totm=prev.totm}
  if (! vals.hours && prev.hours) {vals.hours=prev.hours} //comment this row?
  if (! vals.level && prev.level) {vals.level=prev.level}
  if (! vals.subj && prev.subj) {vals.subj=prev.subj}
  //Logger.log('prev2vals dt='+vals.date+ ' vls='+JSON.stringify(vals));
}

function appendRows2Maakav(rows) {
  if (! rows.length){
    return 'empty';
  }
  if (! rows.every(e => e[0]>1)) {
    writeLog('empty dates rows='+JSON.stringify(rows));
    return 'empty dates';
  }
  let sh=getMaakavSS().getSheetByName('all');
  let p1=sh.getLastRow()+1;
  sh.getRange(p1,1,rows.length,rows[0].length).setValues(rows);
  let dys=sh.getRange(p1,1,rows.length,3).getValues();
  if (! dys.every(e => e[0]>1)) {
    writeLog('empty dates inserted dys='+JSON.stringify(dys));
    return 'empty dates';
  }  
}
