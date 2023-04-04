function doPost(e) {
  var CHANNEL_ACCESS_TOKEN = '{LINE BOT TOKEN}';
  var msgs = JSON.parse(e.postData.contents);
  console.log(msgs);
  
  // 設定預設值
  var replyToken = msgs.events[0].replyToken;    // 金鑰
  var clientID = msgs.events[0].source.userId;   // 使用者ID
  var userMessage = msgs.events[0].message.text; // 使用者訊息
  var ExcelSheetName = "查膳具";  // 工作表名稱
  var datas; // 撈出來的資料
  var text;  // 輸出訊息
  var msg,count = false;  
  
  // 檢查是否有金鑰
  if (typeof replyToken == 'undefined') {
    return;
  }
  
  // 檢查是否有使用者訊息
  if (userMessage != "") {
    msg = userMessage.split(" ")[1];
    if (userMessage.split(" ")[0] == "德州" && msg != "") {  // 判斷第一段有"德州"且後面的訊息不為空
      switch(msg) {
        case "查食魂":                                    // 查詢食魂資料
          datas = GetData(msg);
          msg = userMessage.split(" ")[2];
          if (msg != "") {
            for (var i in datas) {
              if (datas[i][0].indexOf(msg) >= 0) {
                text = "少主，這是你要查的資料。\n";
                text += datas[0][0] + "：" + datas[i][0] + "\n";
                text += datas[0][1] + "：" + datas[i][1] + "\n";
                text += datas[0][2] + "：" + datas[i][2] + "\n";
                text += datas[0][3] + "：" + datas[i][3] + "\n";
                text += datas[0][4] + "：" + datas[i][4] + "\n";
                text += datas[0][5] + "：" + datas[i][5] + "\n";
                text += datas[0][9] + "：\n" + datas[i][9];
                output(CHANNEL_ACCESS_TOKEN,replyToken,text);
                count = true;
                break;
              }
            }
            if (!count) {
              text = "少主，找不到這食魂的相關資料。";
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
            }
          }
          else 
          {
            text = "......少主，你是不是多打了個空格？";
            output(CHANNEL_ACCESS_TOKEN,replyToken,text);
          }
          break;
          
        case "查膳具":                                       // 查詢膳具資料
          datas = GetData(msg);
          msg = userMessage.split(" ")[2];
          if (msg != "") {
            for (var i in datas) {
              if (datas[i][0].indexOf(msg) >= 0) {
                text = "少主，這是你要查的資料。\n";
                text += datas[0][0] + "：" + datas[i][0] + "\n";
                text += datas[0][1] + "：" + datas[i][1] + "\n";
                text += datas[0][2] + "：" + datas[i][2] + "\n";
                text += datas[0][3] + "：" + datas[i][3];
                output(CHANNEL_ACCESS_TOKEN,replyToken,text);
                count = true;
                break;
              }
            }
            if (!count) {
              text = "抱歉少主，查無此資料...";
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
            }
          }
          else 
          {
            text = "......少主，你是不是多打了個空格？";
            output(CHANNEL_ACCESS_TOKEN,replyToken,text);
          }
          break;
          
        case "查膳具副本":                                  // 查詢膳具副本資料
          var number = 0, datext = "";
          ExcelSheetName = "查膳具";
          datas = GetData(ExcelSheetName);
          msg = userMessage.split(" ")[2];
          if (msg != "") {
            for (var i in datas) {
              if (datas[i][3].indexOf(msg) >= 0) {
                datext += datas[i][0] + "(" + datas[i][1] + ")\n";
                count = true;
                number += 1;
              }
            }
            if(count) {
              text = "少主，我幫你整理好了。\n";       // 標題
              text += msg + "掉落膳具：\n";
              text += datext;
              text += "總計有" + number + "套膳具可在此打到。";
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
            }
            else {
              text = "德州未找到相關資料，很抱歉少主...";
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
            }
          }
          else 
          {
            text = "......少主，你是不是多打了個空格？";
            output(CHANNEL_ACCESS_TOKEN,replyToken,text);
          }
          break;
          
        case "查好感禮物":                                  // 查詢膳具副本資料
          var datext = "";
          ExcelSheetName = "查食魂";
          datas = GetData(ExcelSheetName);
          msg = userMessage.split(" ")[2];
          if (msg != "") {
            for (var i in datas) {
              if (datas[i][2].indexOf(msg) >= 0) {
                datext += datas[i][0] + "喜歡的禮物便是" + datas[i][2] + "。";
                count = true;
              }
            }
            if(count) {
              text = "少主，根據我的調查...\n";
              text += datext;
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
            }
            else {
              text = "少主，空桑內並未找到喜歡此物的食魂。";
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
            }
          }
          else 
          {
            text = "......少主，你是不是多打了個空格？";
            output(CHANNEL_ACCESS_TOKEN,replyToken,text);
          }
          break;
          
        case "查技能":                                  // 查詢膳具副本資料
          ExcelSheetName = "查食魂";
          datas = GetData(ExcelSheetName);
          msg = userMessage.split(" ")[2];
          if (msg != "") {
            for (var i in datas) {
              if (datas[i][0].indexOf(msg) >= 0) {
                text = "少主，這是你要查的資料。\n";
                text += datas[0][0] + "：" + datas[i][0] + "\n";
                text += datas[0][6] + "：" + datas[i][6] + "\n";
                text += datas[0][7] + "：" + datas[i][7] + "\n";
                text += datas[0][8] + "：" + datas[i][8];
                output(CHANNEL_ACCESS_TOKEN,replyToken,text);
                count = true;
                break;
              }
            }
            if (!count) {
              text = "少主，找不到這食魂的相關資料。";
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
            }
          }
          else 
          {
            text = "......少主，你是不是多打了個空格？";
            output(CHANNEL_ACCESS_TOKEN,replyToken,text);
          }
          break;
          
        case "查BUFF":                                       // 查詢技能BUFF
          ExcelSheetName = "BUFF";
          datas = GetData(ExcelSheetName);
          msg = userMessage.split(" ")[2];
          if (msg != "") {
            for (var i in datas) {
              if (datas[i][0].indexOf(msg) >= 0) {
                text = "少主，" + datas[i][0] + "可"  + datas[i][1] + "\n";
                text += "...這些資訊很重要，我認為你應該熟記下來。";
                output(CHANNEL_ACCESS_TOKEN,replyToken,text);
                count = true;
                break;
              }
            }
            if (!count) {
              text = "抱歉少主，這個德州也不清楚。";
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
            }
          }
          else 
          {
            text = "......少主，你是不是多打了個空格？";
            output(CHANNEL_ACCESS_TOKEN,replyToken,text);
          }
          break;
          
        case "查生活技能":                                  // 查詢生活技能
          var datext = "",sortext = "";
          ExcelSheetName = "查食魂";
          datas = GetData(ExcelSheetName);
          msg = userMessage.split(" ")[2];
          if (msg != "") {
            for (var i in datas) {
              if (datas[i][9].indexOf(msg) >= 0) {
                datext += "(" + datas[i][1] + ")" + datas[i][0] + "\n";
                count = true;
              }
            }
            switch(msg) {
              case "經驗":
                sortext = "御品：初始10%，滿等20%\n珍品：初始8%，滿等16%\n尚良品：初始5%，滿等10%";
                break;
              case "烹飪":
                sortext = "御品：初始15%，滿等30%\n珍品：初始10%，滿等20%\n尚良品：初始5%，滿等10%";
                break;
              case "貝幣":
                sortext = "御品：初始10%，滿等20%\n珍品：初始8%，滿等16%\n尚良品：初始5%，滿等10%";
                break;
              case "堅果":
                sortext = "金秋願林中有10%機率額外獲得1個隨機堅果。";
                break;
              default:
                sortext = "御品：初始25%，滿等50%\n珍品：初始15%，滿等30%\n尚良品：初始10%，滿等20%";
                break;
            }
            if(count) {
              text = "少主，我幫你整理好了。\n";       // 標題
              text += sortext + "\n\n";
              text += "合適的人選如下：\n";
              text += datext;
              text += "不知哪位比較合你心意？";
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
            }
            else {
              text = "德州未找到合適人選，很抱歉少主...";
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
            }
          }
          else 
          {
            text = "......少主，你是不是多打了個空格？";
            output(CHANNEL_ACCESS_TOKEN,replyToken,text);
          }
          break;
          
        case "管理員模式":                                  //管理員身分驗證
          var AdminOn = false;
          datas = GetData(msg);
          for (var i in datas) {
            if (datas[i][0] == clientID) {
              text = "身分核對正確，歡迎您少主！";
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
              AdminOn = true;
              SetData(AdminOn,msg,Number(i)+1,2);
            }
          }
          if (!AdminOn) {                          // 使用者身分不為管理員
            text = "抱歉，少主。你好像沒有管理員身分...";
            output(CHANNEL_ACCESS_TOKEN,replyToken,text);
          }
          break;
          
        case "新增詞彙":                           // 新增聊天對白(一般使用者皆可新增)
          var dt = new Date();
          if ((Number(dt.getDay()) == 0 || Number(dt.getDay()) == 6) && Number(dt.getHours()) >= 20) {
            ExcelSheetName = "德州辭庫";
            datas = GetData(ExcelSheetName);
            msg = userMessage.split(" ")[2];
            userMessage = msg.split("/");
            if (userMessage[0].length <= 2) {
              text = "少主，關鍵字少於三個字是不能被登記進資料中的。";
              output(CHANNEL_ACCESS_TOKEN,replyToken,text);
              break;
            }
            else {
              for (var i in datas) {
                if (datas[i][0] == userMessage[0]) {
                  text = "少主，資料中已登記過相同問題。";
                  output(CHANNEL_ACCESS_TOKEN,replyToken,text);
                  count = true;
                  break;
                }
              }
              ExcelSheetName = "德州自動回覆辭庫";
              var datas2 = GetData(ExcelSheetName);
              for (var i in datas2) {
                if (datas2[i][0] == userMessage[0]) {
                  text = "少主，資料中已登記過相同問題。";
                  output(CHANNEL_ACCESS_TOKEN,replyToken,text);
                  count = true;
                  break;
                }
              }
              if (!count) {
                ExcelSheetName = "德州辭庫";
                SetData(userMessage[0],ExcelSheetName,(datas.length+1),1);
                SetData(userMessage[1],ExcelSheetName,(datas.length+1),2);
                text = "好的，已登記到資料中。";
                output(CHANNEL_ACCESS_TOKEN,replyToken,text);
              }
            }
          }
          else {
            text = "少主，此功能僅限每周六日晚間八至十二點開放使用。";
            output(CHANNEL_ACCESS_TOKEN,replyToken,text);
          }
          break;
          
        default:                                          //聊天功能
          if (msg != undefined) {
            ExcelSheetName = "德州自動回覆辭庫";
            datas = GetData(ExcelSheetName);
            for (var i in datas) {
              if (msg == datas[i][0]) {
                count = true;
                break;
              }
            }
            if(!count) {
              ExcelSheetName = "德州辭庫";
              datas = GetData(ExcelSheetName);
              for (var i in datas) {
                // 若該列中包含搜尋項目
                if (userMessage.indexOf(datas[i][0]) >= 0) {
                  text = datas[i][1];
                  output(CHANNEL_ACCESS_TOKEN,replyToken,text);
                  count = true;
                  break;
                }
              }
              if (!count) {
                text = "少主，不清楚你在說什麼。";
                output(CHANNEL_ACCESS_TOKEN,replyToken,text);
              }
            }
          }
          break;
      }
    }
  }
  
  // 管理員模式下指令
  //  if (AdminOn) {
  //    
  //  }
  
}

// 資料撈出試算表
function GetData(ExcelSheetName) {
  var file = SpreadsheetApp.openById("{ Google Excel ID }");
  var sheet = file.getSheetByName(ExcelSheetName);
  var datas = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues(); // 所有資料
  return datas;
}

// 資料儲存進試算表
function SetData(userMessage,ExcelSheetName,rownub,columnnub) {
  var file = SpreadsheetApp.openById("{ Google Excel ID }");
  var sheet = file.getSheetByName(ExcelSheetName);
  sheet.getRange(rownub,columnnub).setValue(userMessage);
}

// 送出訊息
function output(TOKEN,replyToken,text){
  var url = 'https://api.line.me/v2/bot/message/reply';
  UrlFetchApp.fetch(url, {
      'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': replyToken,
      'messages': [{
        'type': 'text',
        'text': text,
      }],
    }),
  });
}