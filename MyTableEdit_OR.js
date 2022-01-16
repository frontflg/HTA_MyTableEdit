var tName     = '';          // 選択中テーブル
var strWhere  = '';          // 検索・更新条件文
var aKey      = new Array(); // KEY項目フラグ配列
var tSchema   = 'TESTUSER';  // 対象スキーマ
var maxRow    = '15';        // テーブル項目詳細画面検索最大数
// const tDatSrc = 'Provider=OraOLEDB.Oracle;DSN=XE;User ID=TESTUSER;Password=pass123'; // 環境に合わせて変える
// const tDatSrc = 'Provider=MSDASQL;DSN=XE;User ID=TESTUSER;Password=pass123'; // 環境に合わせて変える
// const tDatSrc = 'Provider=MSDASQL;Data Source=XE;User ID=TESTUSER;Password=pass123'; // 環境に合わせて変える
const tDatSrc = 'Provider=OraOLEDB.Oracle;Data Source=ORACLR_CONNECTION_DATA;User ID=TESTUSER;Password=pass123'; // 環境に合わせて変える
// const tDatSrc = 'Provider=OraOLEDB.Oracle;'
//              + '(DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = NEC-PCuser-PC)(PORT = 1521))'
//              + '(CONNECT_DATA = (SERVER = DEDICATED)(SERVICE_NAME = XE)));'
//              + 'User ID=TESTUSER;Password=Pass123';
//const tDatSrc = 'Provider=OraOLEDB.Oracle;'
//              + '(DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = IPC)(KEY = EXTPROC1521)))'
//              + '(CONNECT_DATA = (SID = CLRExtProc)(PRESENTATION = RO)));'
//              + 'User ID=TESTUSER;Password=Pass123';
// テーブル一覧画面
function setList() {
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  var mySql = "SELECT T.TABLE_NAME,T.OWNER,T.NUM_ROWS,C.COMMENTS FROM ALL_TABLES T";
  mySql += " LEFT JOIN  USER_TAB_COMMENTS C ON T.TABLE_NAME = C.TABLE_NAME"
  mySql += " WHERE T.OWNER = '" + tSchema + "' ORDER BY T.OWNER,T.TABLE_NAME";
  cn.Open(tDatSrc);
  try {
    rs.Open(mySql, cn);
  } catch (e) {
    cn.Close();
    alert('対象テーブル検索不能' + (e.number & 0xFFFF) + ' ' + e.message + ' ' + mySql);
    return;
  }
  if (rs.EOF){
    rs.Close();
    cn.Close();
    rs = null;
    cn = null;
    clrScr();
    $('#tabs').tabs( { active: 1} );
    return;
  }
  var strDoc = '';
  while (!rs.EOF){
    strDoc += '<tr><td style="width:250px;"><a href="#" onClick=colPage("' + rs(0).value + '")>' + rs(0).value + '</a></td>';
    strDoc += '<td width="100">' + rs(1).value + '</a></td>';
    strDoc += '<td width="80" align="RIGHT">' + rs(2).value + '</a></td>';
    strDoc += '<td width="300">' + rs(3).value + '</a></td></tr>';
    rs.MoveNext();
  }
  $('#lst01').replaceWith('<tbody id="lst01">' + strDoc + '</tbody>');
  rs.Close();
  cn.Close();
  rs = null;
  cn = null;
  strDoc = '';
  $('#tabs').tabs( { active: 0} );
  $('#li02').css('visibility','hidden');
  $('#li03').css('visibility','hidden');
}
// テーブル項目詳細画面
function colPage(tName) {
  maxRow = $('#maxRow').val();
  if ( isNaN(maxRow) ) { 
     alert('件数は数字を入力してください！');
     maxRow = "999"  // 入力なしの時の最大値をセット
  }
  var whereRow = $('#whereRow').val();
  // テーブル項目情報の検索
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  var mySql = "SELECT CC.COMMENTS,C.COLUMN_NAME,C.DATA_TYPE,NVL(C.DATA_PRECISION,C.DATA_LENGTH),CK.POSITION"
            + " FROM ALL_TAB_COLUMNS C"
            + " LEFT JOIN ALL_COL_COMMENTS CC ON C.OWNER = CC.OWNER"
            + " AND C.TABLE_NAME = CC.TABLE_NAME AND C.COLUMN_NAME = CC.COLUMN_NAME"
            + " LEFT JOIN (SELECT C3.TABLE_NAME,C3.COLUMN_NAME,C3.POSITION FROM USER_CONS_COLUMNS C3"
            + " INNER JOIN USER_CONSTRAINTS C4 ON C3.CONSTRAINT_NAME = C4.CONSTRAINT_NAME"
            + " WHERE C4.CONSTRAINT_TYPE = 'P') CK"
            + " ON CK.TABLE_NAME = C.TABLE_NAME AND CK.COLUMN_NAME = C.COLUMN_NAME"
            + " WHERE C.OWNER = '" + tSchema + "' AND C.TABLE_NAME =  '" + tName + "' ORDER BY C.COLUMN_ID";
  cn.Open(tDatSrc);
  try {
    rs.Open(mySql, cn);
  } catch (e) {
    cn.Close();
    document.write('対象レコード検索不能' + (e.number & 0xFFFF) + ' ' + e.message + ' ' + mySql);
    alert('対象レコード検索不能');
    return;
  }
  var strDocL = '';
  var strDocR = '';
  var strDoc1 = '';
  var strDoc2 = '';
  var strDoc3 = '';
  var strKey = tName + ' WHERE ';
  aKey = [];
  var cmtFlg = 0;                            // 項目コメント無し
  var colNo = 0;                             // 項目カウンタ
  while (!rs.EOF){
    if (rs(0).value != '') { cmtFlg = 1; }   // 項目コメント有り
    var dtype = rs(2).value;                 // データ型
    var txtNum = 60;                         // 幅
    if (dtype == 'DATE') {
      txtNum = 90;
    } else if (dtype == 'TIME') {
      txtNum = 70;
    } else if (dtype == 'TIMESTAMP(6)') {
      txtNum = 145;
    } else if (dtype == 'TEXT') {
      txtNum = 400;
    } else if (dtype == 'NVARCHAR2') {
      txtNum = rs(3).value * 4 + 10;
      if (txtNum > 400) { txtNum = 400; }
      if (txtNum < 80) { txtNum = 80; }
    } else if (dtype == 'VARCHAR2') {
      txtNum = rs(3).value * 8 + 10;
      if (txtNum > 400) { txtNum = 400; }
      if (txtNum < 80) { txtNum = 80; }
    }
    strDoc1 += '<td style="width:' + txtNum + 'px;">' + rs(0).value + '</td>';
    if (rs(4).value != null) {
      strDoc2 += '<td style="width:' + txtNum + 'px;"><font color="aqua">' + rs(1).value + '</font></td>';
      if (strKey.slice(-6) != 'WHERE ' ) { strKey += ' AND ' }
      strKey += rs(1).value + '★@' + ('0' + colNo).slice(-2);
      aKey[colNo] = 1;
    } else {
      strDoc2 += '<td style="width:' + txtNum + 'px;">' + rs(1).value + '</td>';
      aKey[colNo] = 0;
    }
    strDoc3 += '<td nowrap>' + dtype + '</td>';
    rs.MoveNext();
    colNo += 1;
  }
  if (cmtFlg == 0) {
    strDocL  = '<tr style="display: none;"><td></td></tr><tr class="bg-primary">';
    strDocL += '<td style="width:55px;  height:60px;" rowspan="2" valign="bottom">';
    strDocL += '<input type="button" style="height:27px;" value="新規" onClick="insPage(\'' + tName + '\')" ></td></tr>';
    strDocR  = '<tr style="display: none;">' + strDoc1 + '<td class="dummyColumn"></td></tr>'
  } else {
    strDocL  = '<tr class="bg-primary"><td style="width:55px; height:90px;" rowspan="3" valign="bottom">';
    strDocL += '<input type="button" style="height:27px;" value="新規" onClick="insPage(\'' + tName + '\')" ></td></tr>';
    strDocR  = '<tr class="bg-primary">' + strDoc1 + '<td class="dummyColumn"></td></tr>'
  }
  strDocR += '<tr class="bg-primary">' + strDoc2 + '<td class="dummyColumn"></td></tr>'
  strDocR += '<tr class="bg-primary">' + strDoc3 + '<td class="dummyColumn"></td></tr>';
  $('#hdr02L').replaceWith('<tbody id="hdr02L" style="color:white;">' + strDocL + '</tbody>');
  $('#hdr02R').replaceWith('<tbody id="hdr02R" style="color:white;">' + strDocR + '</tbody>');
  rs.Close();
  cn.Close();
  // テーブルレコードの検索
  rs = null;
  cn = null;
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  mySql = "SELECT * FROM " + tSchema + "." + tName;
  if (whereRow) {
    mySql += " WHERE ROWNUM <= maxRow AND " + whereRow;
  }
  cn.Open(tDatSrc);
  strDocL = '';
  strDocR = '';
  try {
    rs.Open(mySql, cn);
  } catch (e) {
    cn.Close();
    document.write('対象レコード検索不能' + (e.number & 0xFFFF) + ' ' + e.message + ' ' + mySql);
    alert('対象レコード検索不能');
    return;
  }
  while (!rs.EOF){
    strWhere = strKey;
    var strRow = '';
    for ( var i = 0; i < rs.Fields.Count; i++ ) {
      if (rs(i).Type == 133) {
        strRow += '<td style="width:90px;">';
        if (rs(i).Value != null) { strRow += formatDate(rs(i).Value,'YYYY-MM-DD'); }
      } else if (rs(i).Type == 134) {
        strRow += '<td style="width:70px;">';
        if (rs(i).Value != null) { strRow += formatDate(rs(i).Value,'hh:mm:ss'); }
      } else if (rs(i).Type == 135) {
        if (rs(i).Precision == 19)  {
          strRow += '<td style="width:90px;">';
          if (rs(i).Value != null) { strRow += formatDate(rs(i).Value,'YYYY-MM-DD'); }
        } else {
          strRow += '<td style="width:145px;">';
          if (rs(i).Value != null) { strRow += formatDate(rs(i).Value,'YYYY-MM-DD hh:mm:ss'); }
        }
      } else if (rs(i).Type == 203) {
        strRow += '<td style="width:400px;">' + rs(i).Value;
      } else if (rs(i).Type == 202) {
        var txtNum = rs(i).DefinedSize * 8 + 10;
        if (txtNum > 400) { txtNum = 400; }
        if (txtNum < 80) { txtNum = 80; }
        strRow += '<td style="width:' + txtNum + 'px;">' + rs(i).Value;
      } else {
        strRow += '<td style="width:60px;">' + rs(i).Value;
      }
      strRow += '</td>';
      var array = [8,129,133,134,135,201,202,203];
      if (array.indexOf(rs(i).Type) >= 0) {
        strWhere = strWhere.replace('@' + ('0' + i).slice(-2),'※' + rs(i).Value + '※');
      } else {
        strWhere = strWhere.replace('@' + ('0' + i).slice(-2),rs(i).Value);
      }
    }
    strDocL += '<tr><td style="width:55px; height: 30px;" align="center"><input type="button" style="height:27px;" value="編集" onClick="updPage(\'' + strWhere + '\')" ></td></tr>';
    strDocR += '</tr>' + strRow + '</tr>';
    rs.MoveNext();
  }
  $('#tName2').replaceWith('<div id="tName2">' + tName + '</div>');
  $('#tName3').replaceWith('<div id="tName3">' + tName + '</div>');
  $('#reCol').replaceWith('<input type="button" style="height:27px;" value="再検索" onClick="colPage(\'' + tName + '\')">');
  $('#lst02L').replaceWith('<tbody id="lst02L">' + strDocL + '</tbody>');
  $('#lst02R').replaceWith('<tbody id="lst02R">' + strDocR + '</tbody>');
  rs.Close();
  cn.Close();
  rs = null;
  cn = null;
  $('#tabs').tabs( { active: 1} );
  $('#li02').css('visibility','visible');
  $('#li03').css('visibility','hidden');
}
// レコード編集画面
function updPage(updWhere) {
  strWhere = updWhere;
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  // 代替文字　★：イコール、※：￥マーク(文字)
  var mySql = "SELECT * FROM " + updWhere.replace(/★/g, ' = ').replace(/※/g, '\'');
  cn.Open(tDatSrc);
  try {
    rs.Open(mySql, cn);
  } catch (e) {
    cn.Close();
    document.write('対象レコード検索不能' + (e.number & 0xFFFF) + ' ' + e.message + ' ' + mySql);
    alert('対象レコード検索不能');
    return;
  }
  var strDoc = '';
  if (!rs.EOF){
    for ( var i = 0; i < rs.Fields.Count; i++ ) {
      strDoc += '<tr>';
      strDoc += '<td width="150">' + rs(i).Name + '</td><td width="95">';
      
      if (rs(i).Type == 202) { strDoc += 'NVARCHAR2';
      } else if (rs(i).Type == 131) { strDoc += 'NUMBER';
      } else if (rs(i).Type == 133) { strDoc += 'DATE';
      } else if (rs(i).Type == 134) { strDoc += 'TIME';
      } else if (rs(i).Type == 135) {
        if (rs(i).Precision == 19)  { strDoc += 'DATE';
        } else {                      strDoc += 'TIMESTAMP(6)';
        }
      } else if (rs(i).Type == 203) { strDoc += 'VARCHAR2';
      } else if (rs(i).Type ==  16) { strDoc += 'BINARY_DOUBLE';
      } else if (rs(i).Type ==   3) { strDoc += 'BINARY_FLOAT';
      } else { strDoc += rs(i).Type; }
      if (rs(i).Type == 3 || rs(i).Type == 16 || rs(i).Type == 131) {
        var colSize = Math.round(rs(i).Precision * 1.3);   // 数値サイズ
      } else {
        var colSize = Math.round(rs(i).DefinedSize * 1.3); // 文字列サイズ
      }
      if (colSize > 142) { colSize = 142; }
      strDoc += '</td><td width="50">' + colSize + '</td>';
      if (aKey[i] == 1) {                                // KEY項目は表示（入力不可）
        if (rs(i).Value == '') {
          strDoc += '<td></td>';
        } else if (rs(i).Type == 133) {
          strDoc += '<td>' + formatDate(rs(i).Value,'YYYY-MM-DD') + '</td>';
        } else if (rs(i).Type == 134) {
          strDoc += '<td>' + formatDate(rs(i).Value,'hh:mm:ss') + '</td>';
        } else if (rs(i).Type == 135) {
          strDoc += '<td>' + formatDate(rs(i).Value,'YYYY-MM-DD hh:mm') + '</td>';
        } else if (rs(i).Type == 3 || rs(i).Type == 16 || rs(i).Type == 131) {
          strDoc += '<td>0</td>';
        } else {
          strDoc += '<td>' + rs(i).Value + '</td>';
        }
      } else {
        if (rs(i).Value == 0) {
          strDoc += '<td><input type="text" value="0" id="' + rs(i).Name + '" size=5></td>';
        } else if (rs(i).Value == '' || rs(i).Value == null) {
          if (rs(i).Type == 133) { strDoc += '<td><input type="date" ';
          } else if (rs(i).Type == 134) { strDoc += '<td><input type="time" ';
          } else if (rs(i).Type == 135) { strDoc += '<td><input type="datetime" ';
          } else if (rs(i).Type == 3 || rs(i).Type == 16 || rs(i).Type == 131) {
            strDoc += '<td><input type="number" value=0 ';
          } else { strDoc += '<td><input type="text" '; }
          strDoc += 'id="' + rs(i).Name + '"></td>';
        } else if (rs(i).Type == 133) {
          strDoc += '<td><input type="date" id="' + rs(i).Name
                  + '" value="' + formatDate(rs(i).Value,'YYYY-MM-DD') + '"></td>';
        } else if (rs(i).Type == 134) {
          strDoc += '<td><input type="time" id="' + rs(i).Name
                  + '" value="' + formatDate(rs(i).Value,'hh:mm:ss') + '"></td>';
        } else if (rs(i).Type == 135) {
          strDoc += '<td><input type="datetime" id="' + rs(i).Name
                  + '" value="' + formatDate(rs(i).Value,'YYYY-MM-DD hh:mm:ss') + '"></td>';
        } else if (rs(i).Type == 203) {
            var rs2 = new ActiveXObject('ADODB.Recordset');
            var mySql2 = "SELECT SUBSTR(" + rs(i).Name + ",1,255) FROM " + updWhere.replace(/★/g, ' = ').replace(/※/g, '\'');
            rs2.Open(mySql2, cn);
            strDoc += '<td><input type="text" id="' + rs(i).Name
                   + '" value="' + rs2(0).Value + '" size=142" maxlength=255"></td>';
            rs2.Close();
        } else if (rs(i).Type == '3' || rs(i).Type == '16' || rs(i).Type == '131') {
          strDoc += '<td><input type="number" id="' + rs(i).Name
                  + '" value="' + rs(i).Value + '" size="' + colSize
                  + '" maxlength="' + colSize + '"></td>';
        } else {
          strDoc += '<td><input type="text" id="' + rs(i).Name
                  + '" value="' + rs(i).Value + '" size="' + colSize
                  + '" maxlength="' + colSize + '"></td>';
        }
      }
      strDoc += '</tr>';
    }
  }
  $('#lst03').replaceWith('<tbody id="lst03">' + strDoc + '</tbody>');
  rs.Close();
  cn.Close();
  rs = null;
  cn = null;
  $('#insert').hide();
  $('#update').show();
  $('#delete').show();
  $('#tabs').tabs( { active: 2} );
  $('#li03').css('visibility','visible');
}
// レコード新規画面
function insPage(tblName) {
  tName = tblName;
  $('#tName2').replaceWith('<div id="tName2">' + tName + '</div>');
  $('#tName3').replaceWith('<div id="tName3">' + tName + '</div>');
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  var mySql = "SELECT C.COLUMN_ID,C.COLUMN_NAME,C.DATA_TYPE,NVL(C.DATA_PRECISION,C.DATA_LENGTH)"
            + " FROM ALL_TAB_COLUMNS C"
            + " WHERE C.OWNER = '" + tSchema + "' AND C.TABLE_NAME =  '" + tName + "' ORDER BY C.COLUMN_ID";
  cn.Open(tDatSrc);
  try {
    rs.Open(mySql, cn);
  } catch (e) {
    cn.Close();
    document.write('対象レコード検索不能' + (e.number & 0xFFFF) + ' ' + e.message + ' ' + mySql);
    alert('対象レコード検索不能');
    return;
  }
  var strDoc = '';
  var rCnt = 0;
  var rName = '';
  var rTYpe = '';
  var rSize = 0;
  while (!rs.EOF){
    rCnt  = rs(0).Value -1;
    rName = rs(1).Value;
    rType = rs(2).Value;
    rSize = rs(3).Value;
    strDoc += '<tr>';
    if ( aKey[rCnt] == 1 ) {
      strDoc += '<td width="150"><font color="red">' + rName + '</font></td><td width="95">';
    } else {
      strDoc += '<td width="150">' + rName + '</td><td width="95">';
    }
    strDoc += rType + '</td><td width="50">';
    if (rType == 'DATE') {
      strDoc += '<td><input type="date" id="' + rName + '"></td>';
    } else if (rType == 'TIME') {
      strDoc += '<td><input type="time" id="' + rName + '"></td>';
    } else if (rType == 'DATETIME' || rType == 'TIMESTAMP(6)') {
      strDoc += '<td><input type="datetime" id="' + rName + '"></td>';
    } else if (rType == 'CHAR') {
      strDoc += rSize + '<td><input type="longtext" id="' + rName
               + '" size=144" maxlength=255"></td>';
    } else if (rType == 'FLOAT' || rType == 'BINARY_DOUBLE' || rType == 'NUMBER' || 
               rType == 'BINARY_FLOAT' || rType == 'LONG') {
      strDoc += rSize + '<td><input type="number" id="' + rName
             + '" size="' + (rSize * 1.3) + '" maxlength="' + rSize + '"></td>';
    } else {
      strDoc += rSize + '<td><input type="text" id="' + rName;
      if (rSize < 144) {
        strDoc += '" size="' + (rSize * 1.3) + '" maxlength="' + rSize + '"></td>';
      } else {
        strDoc += '" size=144" maxlength=255"></td>';
      }
    }
    strDoc += '</tr>';
    rs.MoveNext();
  }
  $('#lst03').replaceWith('<tbody id="lst03">' + strDoc + '</tbody>');
  rs.Close();
  cn.Close();
  rs = null;
  cn = null;
  $('#insert').show();
  $('#update').hide();
  $('#delete').hide();
  $('#tabs').tabs( { active: 2} );
  $('#li03').css('visibility','visible');
}
// 日付時刻のフォーマット
function formatDate(date, format) {
  var day = new Date(date);
  format = format.replace(/YYYY/, day.getFullYear());
  format = format.replace(/MM/, ('0' + (day.getMonth() + 1)).slice(-2));
  format = format.replace(/DD/, ('0' + day.getDate()).slice(-2));
  format = format.replace(/hh/, ('0' + day.getHours()).slice(-2));
  format = format.replace(/mm/, ('0' + day.getMinutes()).slice(-2));
  format = format.replace(/ss/, ('0' + day.getSeconds()).slice(-2));
  return format;
}
// 更新処理
function updRec() {
  var mySql = "";
  var errFlg = 0;
  tName = $('#tName3').text();
  $('#lst03 input').each(function() {         // ゆくゆくはtextareaも拾いたい
    if (mySql == "") { 
      mySql += "UPDATE " + tName + " SET ";
    } else {
      mySql += ",";
    }
    if ($(this).val() == '') {
      mySql += $(this).attr('id') + " = null";
    } else if ($(this).attr('type') == "number") {
      if ( isNaN($(this).val()) ) { 
        atError ( $(this).attr('id'), '数値を入力してください！');
        errFlg = 1;
        return false;
      }
      mySql += $(this).attr('id') + " = " + $(this).val();
    } else if ($(this).attr('type') == "date") {
      if ( !isDate ( $(this).val()) ) {
        atError ( $(this).attr('id'), '日付形式(YYYY-MM-DD)で入力してください！');
        errFlg = 1;
        return false;
      }
      mySql += $(this).attr('id') + " = '" + $(this).val() + "'";
    } else if ($(this).attr('type') == "time") {
      if ( !isTime ( $(this).val()) ) {
        atError ( $(this).attr('id'), '時刻形式(HH:MM:SS)で入力してください！');
        errFlg = 1;
        return false;
      }
      mySql += $(this).attr('id') + " = '" + $(this).val() + "'";
    } else {
      mySql += $(this).attr('id') + " = '" + $(this).val() + "'";
    }
  });
  if (errFlg != 0) {
    alert('エラーがあります、再入力してください！');
    return;
  }
  mySql += strWhere.slice(strWhere.indexOf(" WHERE")).replace(/★/g, ' = ').replace(/※/g, '\'');
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  cn.Open(tDatSrc);
  try {
    var rs = cn.Execute(mySql);
    alert('対象レコード更新完了');
  } catch (e) {
    cn.Close();
    alert('対象レコード更新失敗 ' + (e.number & 0xFFFF) + ' ' + e.message + ' ' + mySql);
    return;
  }
  cn.Close();
  rs = null;
  cn = null;
  $('#li03').css('visibility','hidden');
  colPage(tName);
}
function insRec() {
  var mySql  = "";
  var mySql2 = "";
  var i = 0;
  var errFlg = 0;
  $('#lst03 input').each(function() {
    if (mySql == "") { 
      mySql  += "INSERT INTO " + tName + " (";
      mySql2 += ") VALUES (";
    } else {
      mySql  += ",";
      mySql2 += ",";
    }
    mySql  += $(this).attr('id');
    if ($(this).val() == '') {
      if ( aKey[i] == 1 ) {
        atError ( $(this).attr('id'), 'KEY項目が入力されていません！');
        errFlg = 1;
        return false;
      }
      mySql2 += "null";
    } else if ($(this).attr('type') == "number") {
      if ( isNaN($(this).val()) ) { 
        atError ( $(this).attr('id'), '数値を入力してください！');
        errFlg = 1;
        return false;
      }
      mySql2 += $(this).val();
    } else if ($(this).attr('type') == "date") {
      if ( !isDate ( $(this).val()) ) {
        atError ( $(this).attr('id'), '日付形式(YYYY-MM-DD)で入力してください！');
        errFlg = 1;
        return false;
      }
      mySql2 += " '" + $(this).val() + "'";
    } else if ($(this).attr('type') == "time") {
      if ( !isTime ( $(this).val()) ) {
        atError ( $(this).attr('id'), '時刻形式(HH:MM:SS)で入力してください！');
        errFlg = 1;
        return false;
      }
      mySql2 += " '" + $(this).val() + "'";
    } else {
      mySql2 += " '" + $(this).val() + "'";
    }
    i = i + 1;
  });
  if (errFlg != 0) {
    alert('エラーがあります、再入力してください！');
    return;
  }
  mySql += mySql2 + ")";
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  cn.Open(tDatSrc);
  try {
    var rs   = cn.Execute(mySql);
    alert('対象レコード登録完了');
  } catch (e) {
    cn.Close();
    if ((e.number & 0xFFFF) == '1062' || (e.number & 0xFFFF) == '3604') {
      alert('対象レコードは、既に登録されています。');
    } else {
      alert('対象レコード登録失敗 ' + (e.number & 0xFFFF) + ' ' + e.message + ' ' + mySql);
    }
    return;
  }
  cn.Close();
  rs = null;
  cn = null;
  $('#li03').css('visibility','hidden');
  colPage(tName);
}
function delRec() {
  var mySql = "DELETE FROM " + strWhere.replace(/★/g, ' = ').replace(/※/g, '\'');
  if( confirm('本当に削除しますか？')) {
  } else {
    alert('削除キャンセルしました！');
    return;
  }
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  cn.Open(tDatSrc);
  try {
    var rs = cn.Execute(mySql);
    alert('対象レコード削除完了');
  } catch (e) {
    cn.Close();
    alert('対象レコード削除失敗 ' + (e.number & 0xFFFF) + ' ' + e.message + ' ' + mySql);
    return;
  }
  cn.Close();
  rs = null;
  cn = null;
  $('#li02').css('visibility','hidden');
  $('#li03').css('visibility','hidden');
  setList();
}
function isDate ( strDate ) {
  if (strDate == '') return true;
  if(!strDate.match(/^\d{4}-\d{1,2}-\d{1,2}$/)){
    return false;
  } 
  var date = new Date(strDate);  
  if(date.getFullYear() !=  strDate.split('-')[0] 
    || date.getMonth() != strDate.split('-')[1] - 1
    || date.getDate() != strDate.split('-')[2]){
    return false;
  } else {
    return true;
  }
}
function isTime ( strTime ) {
  if (strTime == '') return true;
  if(!strTime.match(/^\d{1,2}:\d{1,2}:\d{1,2}$/)){
    if(!strTime.match(/^\d{1,2}:\d{1,2}$/)){
      return false;
    }
  }
  var arrayOfTime = strTime.split(':');
  if (arrayOfTime[0] > 24) { return false; }
  if (arrayOfTime[1] > 60) { return false; }
  if (arrayOfTime.length == 2) { return true; }
  if (arrayOfTime[2] > 60) { return false; }
  if (arrayOfTime.length > 3) { return false; }
  return true;
}
function atError ( str, msg ) {
  alert(msg);
  $('#' + str).focus();
}