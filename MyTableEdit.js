var tName = '';         // 選択中テーブル
var strWhere = '';      // 検索・更新条件文
var aKey = new Array(); // KEY項目フラグ配列
const tSchema = 'test'; // 環境に合わせて変える
const tDatSrc ='Provider=MSDASQL; Data Source=Connector_MariaDB'; // 環境に合わせて変える
// テーブル一覧画面
function setList() {
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  var mySql = "SELECT TABLE_NAME,TABLE_COMMENT,TABLE_ROWS,DATE_FORMAT(CREATE_TIME,'%Y/%m/%d %H:%i:%s')"
            + " FROM information_schema.TABLES WHERE TABLE_SCHEMA = '" + tSchema + "'";
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
    strDoc += '<tr><td style="width:150px;"><a href="#" onClick=colPage("' + rs(0).value + '")>' + rs(0).value + '</a></td>';
    strDoc += '<td width="300">' + rs(1).value + '</a></td>';
    strDoc += '<td width="80" align="RIGHT">' + rs(2).value + '</a></td>';
    strDoc += '<td width="200">' + rs(3).value + '</a></td></tr>';
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
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  var mySql = "SELECT C.COLUMN_COMMENT,C.COLUMN_NAME,C.COLUMN_TYPE,K.ORDINAL_POSITION"
            + " FROM information_schema.`COLUMNS` C"
            + " LEFT OUTER JOIN information_schema.KEY_COLUMN_USAGE K"
            + " ON (K.TABLE_NAME = C.TABLE_NAME"
            + " AND K.COLUMN_NAME = C.COLUMN_NAME)"
            + " WHERE C.TABLE_SCHEMA = '" + tSchema + "' AND C.TABLE_NAME = '" + tName + "'";
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
    if (dtype == 'date') {
      txtNum = 90;
    } else if (dtype == 'time') {
      txtNum = 70;
    } else if (dtype == 'datetime') {
      txtNum = 130;
    } else if (dtype == 'text') {
      txtNum = 400;
    } else if (dtype.slice(0,7) == 'varchar') {
      txtNum = dtype.slice(8,(dtype.length -1)) * 8 + 10;
      if (txtNum > 400) { txtNum = 400; }
      if (txtNum < 80) { txtNum = 80; }
    }
    strDoc1 += '<td style="width:' + txtNum + 'px;">' + rs(0).value + '</td>';
    if (rs(3).value != null) {
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
  rs = null;
  cn = null;
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  mySql = "SELECT * FROM " + tName;
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
        strRow += '<td style="width:130px;">';
        if (rs(i).Value != null) { strRow += formatDate(rs(i).Value,'YYYY-MM-DD hh:mm:ss'); }
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
      strDoc += '<td width="150">' + rs(i).Name + '</td><td width="60">';
      if (rs(i).Type == 202) { strDoc += 'varchar';
      } else if (rs(i).Type == 133) { strDoc += 'date';
      } else if (rs(i).Type == 134) { strDoc += 'time';
      } else if (rs(i).Type == 135) { strDoc += 'datetime';
      } else if (rs(i).Type == 203) { strDoc += 'text';
      } else if (rs(i).Type ==  16) { strDoc += 'tinyint';
      } else if (rs(i).Type ==   3) { strDoc += 'int';
      } else { strDoc += rs(i).Type; }
      strDoc += '</td><td width="50">' + rs(i).DefinedSize + '</td>';
      if (aKey[i] == 1) {                                // KEY項目は表示（入力不可）
        if (rs(i).Value == '') {
          strDoc += '<td></td>';
        } else if (rs(i).Type == 133) {
          strDoc += '<td>' + formatDate(rs(i).Value,'YYYY-MM-DD') + '</td>';
        } else if (rs(i).Type == 134) {
          strDoc += '<td>' + formatDate(rs(i).Value,'hh:mm:ss') + '</td>';
        } else if (rs(i).Type == 135) {
          strDoc += '<td>' + formatDate(rs(i).Value,'YYYY-MM-DD hh:mm') + '</td>';
        } else {
          strDoc += '<td>' + rs(i).Value + '</td>';
        }
      } else {
        if (rs(i).Value == '') {
          if (rs(i).Type == 133) { strDoc += '<td><input type="date" ';
          } else if (rs(i).Type == 134) { strDoc += '<td><input type="time" ';
          } else if (rs(i).Type == 135) { strDoc += '<td><input type="datetime" ';
          } else if (rs(i).Type == 3 || rs(i).Type == 16) { strDoc += '<td><input type="number" ';
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
      //    strDoc += '<td><textarea rows="4" cols="144" id="'
      //            + rs(i).Name + '">' + rs(i).Value + '</textarea></td>';
      // リストだとうまく行くのに個別にSELECTするとTEXT型がとって来れない（valueが常にnull）、何故？
            strDoc += '<td>NAME = ' + rs(i).Name
                    + ' VALUE = ' + rs(i).Value
                    + ' PRECISION = ' + rs(i).Precision
                    + ' SCALE = ' + rs(i).NumericScale
                    + ' TYPE = ' + rs(i).Type
                    + ' DEFSIZE = ' + rs(i).DefinedSize
                    + ' ACTSIZE = ' + rs(i).ActualSize
                    + '</td>';
        } else if (rs(i).Type == 3 || rs(i).Type == 16) {
          strDoc += '<td><input type="number" id="' + rs(i).Name
                  + '" value="' + rs(i).Value + '" size="' + Math.round(rs(i).DefinedSize * 1.3)
                  + '" maxlength="' + rs(i).DefinedSize + '"></td>';
        } else {
          strDoc += '<td><input type="text" id="' + rs(i).Name
                  + '" value="' + rs(i).Value + '" size="' + Math.round(rs(i).DefinedSize * 1.3)
                  + '" maxlength="' + rs(i).DefinedSize + '"></td>';
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
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  var mySql = "SELECT * FROM " + tName + " LIMIT 1";
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
      if ( aKey[i] == 1 ) {
        strDoc += '<td width="150"><font color="red">' + rs(i).Name + '</font></td><td width="60">';
      } else {
        strDoc += '<td width="150">' + rs(i).Name + '</td><td width="60">';
      }
      if (rs(i).Type == 202) { strDoc += 'varchar';
      } else if (rs(i).Type == 133) { strDoc += 'date';
      } else if (rs(i).Type == 134) { strDoc += 'time';
      } else if (rs(i).Type == 135) { strDoc += 'datetime';
      } else if (rs(i).Type == 203) { strDoc += 'text';
      } else if (rs(i).Type ==  16) { strDoc += 'tinyint';
      } else if (rs(i).Type ==   3) { strDoc += 'int';
      } else { strDoc += rs(i).Type; }
      strDoc += '</td><td width="50">' + rs(i).DefinedSize + '';
      if (rs(i).Type == 133) {
        strDoc += '<td><input type="date" id="' + rs(i).Name + '"></td>';
      } else if (rs(i).Type == 134) {
        strDoc += '<td><input type="time" id="' + rs(i).Name + '"></td>';
      } else if (rs(i).Type == 135) {
        strDoc += '<td><input type="datetime" id="' + rs(i).Name + '"></td>';
      } else if (rs(i).Type == 203) {
        strDoc += '<td><textarea rows="4" cols="144" id="' + rs(i).Name + '"></textarea></td>';
      } else if (rs(i).Type == 3 || rs(i).Type == 16) {
        strDoc += '<td><input type="number"   id="' + rs(i).Name
                + '" size="' + Math.round(rs(i).DefinedSize * 1.3)
                + '" maxlength="' + rs(i).DefinedSize + '"></td>';
      } else {
        strDoc += '<td><input type="text" id="' + rs(i).Name
                + '" size="' + Math.round(rs(i).DefinedSize * 1.3)
                + '" maxlength="' + rs(i).DefinedSize + '"></td>';
      }
      strDoc += '</tr>';
    }
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
  tName = strWhere.slice(0,strWhere.indexOf(" WHERE"));
  $('#lst03 input').each(function() {
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
 // alert('エラーがあります、再入力してください！');
    return;
  }
  mySql += strWhere.slice(strWhere.indexOf(" WHERE")).replace(/★/g, ' = ').replace(/※/g, '\'');
//  alert('SQL=' + mySql);
//  return;
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
    // 日付時刻形式(YYYY-MM-DD HH:MM:SS)は未作成
    } else {
      mySql2 += " '" + $(this).val() + "'";
    }
    i = i + 1;
  });
  if (errFlg != 0) {
 // alert('エラーがあります、再入力してください！');
    return;
  }
  mySql += mySql2 + ")";
//  alert('SQL=' + mySql;
//  return;
  var cn = new ActiveXObject('ADODB.Connection');
  var rs = new ActiveXObject('ADODB.Recordset');
  cn.Open(tDatSrc);
  try {
    var rs   = cn.Execute(mySql);
    alert('対象レコード登録完了');
  } catch (e) {
    cn.Close();
    if ((e.number & 0xFFFF) == '1062') {
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
//  alert('削除SQL: ' + mySql);
//  return;
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
  if (arrayOfTime[0] > 60) { return false; }
  if (arrayOfTime[1] > 60) { return false; }
  if (arrayOfTime.length == 2) { return true; }
  if (arrayOfTime[2] > 60) { return false; }
  if (arrayOfTime.length > 3) { return false; }
  return true;
}
// function isDateTime ( strDateTime ) { // 未作成（未対応）
function atError ( str, msg ) {
  alert(msg);
  $('#' + str).focus();
}