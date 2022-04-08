var ab_subjects=['ספרות','תנך','ניהול עסקי','מתמטיקה','ימאות','אזרחות','הסטוריה','אנגלית'];

/*function addNoShow2MaakavMain() {
  collectParams();
  writeLog('Start');
  gp.ab_sat_dt=getDtPlusDays(gp.ab_sun_dt,6);
  gp.rows_to_add2maakav=[];
  let abclasses_sh=getABSS().getSheetByName('חניכים 6.3');
  //let abclasses_sh=getABSS().getSheetByName('חניכים');
  let ar_cols=abclasses_sh.getLastColumn();
  let sh_ar=abclasses_sh.getRange(3,2,abclasses_sh.getLastRow()-2,ar_cols-1).getValues();
  Logger.log('sh_ar='+JSON.stringify(sh_ar));
  for (let c=0; c<sh_ar[1].length; c++){
    Logger.log('c='+c+' sh_ar[2][c]='+sh_ar[2][c]+' sh_ar[1][c]='+sh_ar[1][c]+' sh_ar[3][c]='+sh_ar[3][c]);
    let teac_dts_subj_ar=aBfindTeacDts(sh_ar[2][c],sh_ar[1][c]); 
    Logger.log('teac_dts_subj_ar='+teac_dts_subj_ar)
    //for (let d=0;d<teac_dts_subj_ar[1].length;d++){
      for (let k=3;k<sh_ar.length;k++){
        if (! sh_ar[k][c]) {
          break;
        }
        let kid=sh_ar[k][c].replace(/\(.*$/,'').replace(/\d.*$/,'').replace(/\s+$/,'');
        //Logger.log('sh_ar[k][c]='+sh_ar[k][c]+' kid='+kid);
        abAddNoShowIfNeeded(kid,teac_dts_subj_ar[0],teac_dts_subj_ar[2],teac_dts_subj_ar[1],sh_ar[0][c]); 
        //aBaddNoShowIfNeeded(sh_ar[k][0],teac_dts_subj_ar[0],teac_dts_subj_ar[2],[gp.ab_sun_dt, gp.ab_sat_dt]); //stu, teac[], subj, dt[][]
      }
    //}
  }
  appendRows2Maakav(gp.rows_to_add2maakav);
  Logger.log('added rows:'+gp.rows_to_add2maakav.length);

  checkLog();
}

function aBfindTeacDts(teach_raw, subj_raw) {
  let ts_raw=teach_raw.split('/');
  //Logger.log('ts_raw='+ts_raw);
  let dows_ar=ts_raw.map(e => e.replace(/^.+\(/,'').replace(/\)$/,'')); //dows
  //Logger.log('dows_ar='+dows_ar);
  let tdts_ar=calcDtsFromDows(dows_ar,gp.ab_sun_dt); //[[d1,d2],[d3,d4]]
  //Logger.log('tdts_ar='+tdts_ar);
  let ts_ar=ts_raw.map(e => e.replace(/\(.*\)/,'')); //e.g ['t1','t2']
  let subj=subj_raw;//.replace(/\s.*$/,'');
  return [ts_ar, tdts_ar, subj]; // [['t1','t2'], [[d1,d2],[d3,d4]], subj]
}

function getDtPlusDays(dt, days) {
  let d = new Date(dt.getTime());
  //Logger.log('plus dt='+dt+' days='+days);
  d.setDate(d.getDate() + days);
  //Logger.log(' d='+d);
  return d;
}

function calcDtsFromDows(dows_ar, sun_dt) {
  let dts_ar=dows_ar.map(e => dtsOfDows(e,sun_dt));
  //Logger.log(' calcDtsFromDows dts_ar='+dts_ar);
  return dts_ar;
}

function dtsOfDows(dows, sun_dt) {
  let dt_ar=dows.split(',').map(e => {
              let d = getDtPlusDays(sun_dt,e -1);
              return d; })
  //Logger.log('dows='+dows+' dt_ar='+dt_ar);
  return dt_ar;
}

function abAddNoShowIfNeeded(stu, teac, subj, dt, mindays) {
  //Logger.log('stu='+stu+' teac='+teac+' subj='+subj+' dt='+ JSON.stringify( dt));
  let rows=getHistRow(dt,teac,stu,subj);
  //Logger.log('rows='+rows+' rows.len='+rows.length+ ' dt[0].length='+dt[0].length);
  let miss=(mindays ? mindays : dt[0].length) - rows.length;
  if (miss < 1) {
    //writeLog('no missed classes:'+stu);
    return;
  }
  //Logger.log('miss='+miss);
  abAddNoShow(miss, rows, stu, teac, subj, dt);
  //Logger.log('gp.rows_to_add2maakav='+JSON.stringify(gp.rows_to_add2maakav));
}

function abAddNoShow(miss, rows, stu, teac, subj, dt) {
  let reported_dts_ar=rows.map(e => e[0].getTime() );
  let clas_dts=dt[0];
  let cls_dts_ar=clas_dts.map(e => e.getTime() );
  let count=0;
  for (let i=0;i<cls_dts_ar.length;i++){
    if ( reported_dts_ar.includes(cls_dts_ar[i])){
      continue;
    }
    let stu_ar=getStuAr(stu);
    if (! stu_ar){
      writeLog('invalid pupil:'+stu);
    }
    let act='אני בגרותי - חיסור אוטומטי'; let lvl= 'יא';  let noshow='לא הגיע';
//    let a= [getFmtDtStr(clas_dts[i]),subj,act,,teac[0], stu, lvl, stu_ar[3], '=ROW()', 1, noshow, , ];
    let a= [clas_dts[i],subj,act,,teac[0], stu, lvl, stu_ar[3], '=ROW()', 1, noshow, , ];
    gp.all_maakav_rows.push(a); 
    gp.rows_to_add2maakav.push(a);
    count++;
    if (count == miss) {
      break;
    }
  }
}*/

function getABSS() {
  if (! gp.AB_SS){
    gp.AB_SS=SpreadsheetApp.openById(gp.ab_file_d);
    //gp.AB_SS=SpreadsheetApp.openById('1orOUpZg2cwj0MoICHbbp0rJny4opOFetugogRB8Rs_k');
  }
  return gp.AB_SS;
}
/*
function updatePupilAttendMain() {
  collectParams();
  writeLog('Start');
  let abatt_sh=getABSS().getSheetByName('ניקוד 6.3');
  //let abatt_sh=getABSS().getSheetByName('ניקוד');
  let miss_sh=getABSS().getSheetByName('חיסורים 6.3');
  //let miss_sh=getABSS().getSheetByName('חיסורים');
  abatt_sh.getRange(6,3,abatt_sh.getLastRow()-5,abatt_sh.getLastColumn()-4).clear({contentsOnly: true});
  miss_sh.getRange(6,3,miss_sh.getLastRow()-5,miss_sh.getLastColumn()-4).clear({contentsOnly: true});
  abSortShByName(abatt_sh);

  //abSortShByName(miss_sh);
  let attend_ar=abatt_sh.getRange(4,1,abatt_sh.getLastRow()-3,abatt_sh.getLastColumn()-2).getValues();
  let miss_ar=miss_sh.getRange(4,1,miss_sh.getLastRow()-3,miss_sh.getLastColumn()-2).getValues();
  attend_ar.forEach((e,i) => {miss_ar[i][1] = e[1]});//same kids and order
  for (let i=2;i<attend_ar.length;i++){
    let stu=attend_ar[i][1];
    Logger.log('stu#='+i+' stu='+stu);
    let stua=getStuAr(stu);
    if (!stua) {continue}
    updAttenStuScore(miss_ar, attend_ar, stu, i);
  }
  abatt_sh.getRange(4,1,abatt_sh.getLastRow()-3,abatt_sh.getLastColumn()-2).setValues(attend_ar);
  miss_sh.getRange(4,1,miss_sh.getLastRow()-3,miss_sh.getLastColumn()-2).setValues(miss_ar);
  abSortShByScore(abatt_sh);
  abSortShByScore(miss_sh);
  checkLog();
}

function abSortShByName(sh) {
  var cols=sh.getLastColumn();
  range = sh.getRange(6,2,sh.getLastRow()-5,cols-2);
  range.sort([{column: 2, ascending: true}]);
}

function abSortShByScore(sh) {
  var cols=sh.getLastColumn();
  Logger.log('cols='+cols);
  range = sh.getRange(6,2,sh.getLastRow()-5,cols-2);
  range.sort([{column: cols-1, ascending: false}]);
}

function updAttenStuScore(miss_ar, attend_ar, stu, i) {
    let attdt;
    for (let c=2;c<attend_ar[i].length-1;c++){
      let subj=attend_ar[1][c];
      if (attend_ar[0][c]) {attdt=attend_ar[0][c]};
      //Logger.log('c='+c+' subj='+subj+' attdt='+attdt+' attend_ar[i].length='+attend_ar[i].length);
      let maa_rows=getHistRow(attdt,'',stu,subj);
      let score=abCalcScore(maa_rows);
      //Logger.log('score='+score);
      attend_ar[i][c]=score[0];
      miss_ar[i][c]=score[1];
    }
}
*/
function updAttenStuScoreInTimeRangeMain() {
    collectParams();
    writeLog('Start');
    let sh=getABSS().getSheetByName('ניקוד');
    //let rng=sh.getRange(5,2,3,2);
    let rng=sh.getRange(5,2,sh.getLastRow()-4,3);
    let ar=rng.getValues();
    let attdt;
    for (let i=0;i<ar.length;i++){
      let maa_rows=getHistRows(ar[i][0]);
      //Logger.log('rows'+maa_rows.length);
      //Logger.log('rows'+JSON.stringify( maa_rows));
      let res=abCalcScore(maa_rows);
      ar[i][1]=res[0];
      ar[i][2]=res[1];
      //Logger.log('ari1='+ar[i][1]);
    }
    ar.sort((a, b) => { if (a[1]==b[1]){return 0}
                        if (a[1]>b[1]){return -1}
                        return 1; } )
    rng.setValues(ar);
    sh.getRange(2,2,1,2).setValues([[gp.ab_sun_dt,gp.ab_last_dt]])
    checkLog();
}

function abCalcScore(rows) {
  if (!rows || rows.length<1){
    return ['',''];
  }
  let score=0;
  let miss=0;
  for (let i=0;i<rows.length;i++){
    //Logger.log('rows[i][10]='+rows[i][10]);
    if (['לא הגיע'].includes(rows[i][10])){
      miss+=1;
      continue; //xxx
      //return [0,miss]; //xxx
      //Logger.log('miss+++='+miss);
    }
    if (['בא בזמן',''].includes(rows[i][10])){
      score+=2;
    }
    if (['איחר'].includes(rows[i][10])){
      score++;
    }
//    Logger.log('2score='+score);
    if (['הביא',''].includes(rows[i][11])){
      score++;
    }
//    Logger.log('3score='+score);
    if (['השתתף',''].includes(rows[i][12])){
      score++;
    }
//    Logger.log('4score='+score);
    if (['הפריע'].includes(rows[i][12])){
      score--;
    }
//    Logger.log('score='+score);
    //return [score,miss]; //xxx
  }
  return [score,miss]; //xxx
}

