//月次でチームのチャンネル作ってメッセージ送信する作業
function doMonthlyTeamActivityJob() {

  //トークンID取得
  const tokenID = getTokenID();

  //既存slackユーザーを「名前_ID対応表」シートに出力
  const menberNameList = getSlackUsers(tokenID);
  
  //「実行用シート」シートからチーム名、メンバーを取得
  let teamList = getTeamsAndMembers(menberNameList,tokenID);

  Logger.log(teamList)

  //エラーが起こった時、実行終了
  if(Array.isArray(teamList)===false){
    SpreadsheetApp.getUi().alert("修正してまた実行してね！")
    return
  }
  
  //１チームずつチャンネル作成しメンバーを招待する
  teamList.forEach(function(team){
    BulkInviteMembersToChannel(team.groupName,team.groupMembers,tokenID);
  });
}

/**
 * シートからトークンID取得
 */
function getTokenID() {
  //シート指定
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("トークンID");
  
  const tokenID = sheet.getRange(1,1).getValue();
  return tokenID
}

/**
 * 全ユーザの名前とIDを取得してスプレッドシートに出力する関数
 */
function getSlackUsers(tokenID) {

  //tokenIDが未定義のとき、取りに行く
  if(typeof tokenID == "undefined"){
    tokenID = getTokenID();
  }

  const url = "https://slack.com/api/users.list";
  const options = {
    "method" : "get",
    "contentType": "application/x-www-form-urlencoded",
    "payload" : { 
      "token": tokenID
    }
  };  
  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response);
  
  //全メンバーを格納
  const members = json.members;
  
  //シートに出力する内容を全て格納する変数（ヘッダ含む）
  let table = [["表示名","登録名", "ユーザーID"]];
  
  for (const member of members) {
    
    //削除済、botユーザー、Slackbotを除く
    if (!member.deleted && !member.is_bot && member.id !== "USLACKBOT") {
      let id = member.id;
      let display_name = member.profile.display_name_normalized; //表示名
      let real_name = member.real_name; //登録名
      table.push([display_name, real_name, id]);
    }
    
  }
  
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("名前_ID対応表");
  
  //スプレッドシートに書き込み
  sheet.getRange(1, 1, sheet.getMaxRows()-1, 3).clearContent();
  sheet.getRange(1, 1, table.length, table[0].length).setValues(table);

  //メンバー名の一覧の配列を返す
  //表示名、登録名の両方を含んだ配列
  const memberNameList = []
  members.forEach(function(member){
    memberNameList.push(member.profile.display_name_normalized)
    memberNameList.push(member.real_name)
  })
  return memberNameList;
}

/**
 * チャンネル一覧取得
 */
function getChannelList(tokenID){
  const url = "https://slack.com/api/conversations.list";
  const options = {
    "method" : "get",
    "contentType": "application/x-www-form-urlencoded",
    "payload" : { 
      "token": tokenID
    }
  };  

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response);

  const channelInfoList = json.channels

  //チャンネル名を配列に格納
  let channelNames = []
  channelInfoList.forEach(function(channelInfo){
    channelNames.push(channelInfo.name)
  })

  return channelNames;
}

/**
 * 各カラムのチーム名とメンバーを取得し配列に格納する関数
 */
function getTeamsAndMembers(menberNameList,tokenID){

  //格納する配列作成
  let teamList = [];

  //同名のチャンネルを許容しないために、既存のチャンネルリストを取得
  const channelNameList = getChannelList(tokenID);
  
  //シート指定
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("実行用シート");
  
  //while文用のフラグ
  let columnFlag=true;
  
  //調査するカラム番号
  let columnNumber = 2;
  
  while(columnFlag===true){
    //グループ名取得
    let groupName = sheet.getRange(1,columnNumber).getValue();
    
    //同名のチャンネルが存在する場合、実行停止
    if(channelNameList.includes(groupName)){
      SpreadsheetApp.getUi().alert(groupName + "はもう存在してるよ")
      return
    }
    
    //メンバーを入れるハコ
    let groupMembers = []
    
    //while文用のフラグ
    let rowFlag=true
    
    //調査するカラム番号
    let rowNumber = 2;
  
    //最後の行までチェック
    while(rowFlag===true){
      //メンバー名取得
      let memberName = sheet.getRange(rowNumber,columnNumber).getValue()

      Logger.log(memberName)
      Logger.log(menberNameList)

      //存在しないメンバー名の場合、実行終了[TODO]
      if(menberNameList.includes(memberName)===false){
        SpreadsheetApp.getUi().alert(memberName + "というユーザーは存在しないよ。名前を確かめてね")
        return
      }

      //メンバー名からID取得して配列にpush
      groupMembers.push(getIdFromName(memberName));
      rowNumber=rowNumber+1;
      
      //次の行が空の場合break
      if(sheet.getRange(rowNumber,columnNumber).isBlank()){
        rowFlag=false;
      }
    }
    
    //チームリストにオブジェクトを追加
    teamList.push({groupName:groupName, groupMembers:groupMembers});
    columnNumber=columnNumber+1;
    
    //次の列が空の場合break
    if(sheet.getRange(1,columnNumber).isBlank()){
      columnFlag=false;
    }
  }
  
  //チーム名の重複チェック
  //チーム名の配列
  let createChannelNameList = [];
  teamList.forEach(function(team){
    createChannelNameList.push(team.groupName)
  })

  if(existsSameValue(createChannelNameList)){
    SpreadsheetApp.getUi().alert("重複してるチーム名があるよ")
    return 
  }
  return teamList;
}

/** 配列内で値が重複してないか調べる **/
function existsSameValue(arr){
  var existsSame = false;
  arr.forEach(function(val){
    /// 配列中で arr[i] が最初に出てくる位置を取得
    var firstIndex = arr.indexOf(val);
    /// 配列中で arr[i] が最後に出てくる位置を取得
    var lastIndex = arr.lastIndexOf(val);
 
    if(firstIndex !== lastIndex){
      /// 重複していたら true を返す
      existsSame = true;
    }
  })
  return existsSame;
}

//メンバーの名前からIDを取得する関数。
//表示名を調べたあと、登録名の一致を判定する
//入力：表示名or登録名、出力：ID
function getIdFromName(userName){
  
  //シート指定
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("名前_ID対応表");
  
  let rowNumber = 2;

  let userID;
  
  //下端まで調査.IDの列が空になったら調査終了
  while (sheet.getRange(rowNumber,3).isBlank()===false){
    Logger.log("------------------------------------------------")
    Logger.log(userName)
    Logger.log(sheet.getRange(rowNumber,1).getValue())
    Logger.log(sheet.getRange(rowNumber,2).getValue())
    if(sheet.getRange(rowNumber,1).getValue()==userName){         //表示名が一致するかの判定
    
      //一致したらID出力      
      userID = sheet.getRange(rowNumber,3).getValue()
      return userID;
      
    }else if(sheet.getRange(rowNumber,2).getValue()==userName){   //登録名が一致するかの判定
      //一致したらID出力      
      userID = sheet.getRange(rowNumber,3).getValue()
      return userID;
    }
    rowNumber=rowNumber+1;
  }
  return ;
}

//指定のチャンネルにメンバーを招待
function BulkInviteMembersToChannel(channelName,members,tokenID){

  //チャンネル作成。resの中にid等含まれている
  const channel_res = createSlackGroups(channelName,tokenID);
  
  //メンバー招待
  members.forEach(function(memberId){
    inviteMember(channel_res["channel"]["id"],memberId,tokenID)
  })

  //チャンネルにひとこと送信
  sendMessegeToTheChannel(channel_res["channel"]["id"],tokenID)
}

//指定のメンバーを指定のチャンネルに招待
function inviteMember(channelID, memberID,tokenID){
  
  const url = "https://slack.com/api/conversations.invite";
  const options = {
    "method" : "post",
    "contentType": "application/x-www-form-urlencoded",
    "payload" : { 
      "token": tokenID,
      "channel": channelID,
      "users": memberID
    }
  }
  
  const response = UrlFetchApp.fetch(url, options);
}

//チャンネル作成
function createSlackGroups(channelName,tokenID){
  const url = "https://slack.com/api/conversations.create";
  const options = {
    "method" : "post",
    "contentType": "application/x-www-form-urlencoded",
    "payload" : { 
      "token": tokenID,
      "name": channelName,
            "is_private": false
    }
  };  
  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  return json;
}

//指定のチャンネルに一言送信する関数
function sendMessegeToTheChannel(channelID, tokenID){

  //シートからメッセージ取得
  const messageText = getMessage();

  //メッセージを指定のチャンネルに送信
  const url = "https://slack.com/api/chat.postMessage";
  const options = {
    "method" : "post",
    "contentType": "application/x-www-form-urlencoded",
    "payload" : { 
      "token": tokenID,
      "channel": channelID,
      "text": messageText
    }
  };  

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
}

//メッセージのシートからの取得
function getMessage() {
  //シート指定
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName("メッセージ");
  
  //メッセージの文面取得（１行のみ）
  const messageText = sheet.getRange(1,1).getValue();
  return messageText
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("自動化作業用")
    .addItem("自動招待実行", "doMonthlyTeamActivityJob")
    .addItem("ユーザー一覧出力", "getSlackUsers")
    .addToUi()
}

function showAlert() {
  SpreadsheetApp.getUi().alert("スクリプトが実行されました！")
}