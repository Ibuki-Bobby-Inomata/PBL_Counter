// @ts-nocheck
// 参考: https://qiita.com/ryosuk/items/1a18c663884d2748d4ca
// API_TOKENはワークスペースごとに設定
var API_TOKEN = "xoxp-4141846193926-4182891036818-4316676314433-fae1c262a3a3a5fdb8bc45f17d159aff";

var target_year = '2023-01-18'; // 対象となる期間の最古日付

// ページネーション
var MAX_HISTORY_PAGINATION = 1; // デバッグ用の定数になった
var HISTORY_COUNT_PER_PAGE = 200; // 対象年分とってくる。

var stamps = {};
var counts = [];
var ss = SpreadsheetApp.getActiveSpreadsheet();
var timezone = ss.getSpreadsheetTimeZone();

var slack_members = {};
var all_post_user = {};

function main() {
  var logger = new SlackChannelHistoryLogger();
  logger.run(target_year);
  
  for (var _i = 1, a = 9; _i < a; _i++) {
    updateTeamSheet(_i);
  };
  fullSheet();

  function updateTeamSheet(num) {
    team_name = "team"+String(num)
    var sheet = ss.getSheetByName("team"+String(num));
    sheet.clear();

    var data = [];
    data.push(['なまえ', '発言数', 'リアクション', 'Personal score']);
    Object.entries(all_post_user[team_name]).forEach(function([key, value]){
      data.push([slack_members[key], value[0], value[1], value[0]+(value[1]*0.5)]);
    });
    sheet.getRange(1,1,data.length,4).setValues(data);

    //降順でソート
    sheet.getRange(2,1,data.length,4).sort({column: 1, ascending: true});
  }
  function fullSheet() {
    var sheet = ss.getSheetByName("full");
    sheet.clear();

    var data = [];
    data.push(['なまえ', '発言数', 'リアクション', 'Personal score']);

    for (var num = 1, ma = 9; num < ma; num++) {
      team_name = "team"+String(num)

      Object.entries(all_post_user[team_name]).forEach(function([key, value]){
        data.push([slack_members[key], value[0], value[1], value[0]+(value[1]*0.5)]);
      });
    };
    
    sheet.getRange(1,1,data.length,4).setValues(data);

    //降順でソート
    sheet.getRange(2,1,data.length,4).sort({column: 1, ascending: true});
  }
};

var SlackChannelHistoryLogger = (function () {
    function SlackChannelHistoryLogger() {
        this.memberNames = {};
    }
    SlackChannelHistoryLogger.prototype.requestSlackAPI = function (path, params) {
        if (params === void 0) { params = {}; }
        var url = "https://slack.com/api/" + path + "?";
        var qparams = [];
        for (var k in params) {
            qparams.push(encodeURIComponent(k) + "=" + encodeURIComponent(params[k]));
        }
        url += qparams.join('&');
        try{
          var headers = {
            'Authorization': 'Bearer '+ API_TOKEN
          };
          var options = {
            'headers': headers
          };
          var resp = UrlFetchApp.fetch(url, options);
          var data = JSON.parse(resp.getContentText());
          if (data.error) {
            throw "GET " + path + ": " + data.error;
          }
          return data;
        }catch(e){
          Logger.log(e)
          return "err";
          }
    };
    SlackChannelHistoryLogger.prototype.run = function () {
        var _this = this;
        var usersResp = this.requestSlackAPI('users.list');

        for (const member of usersResp.members) {
          // //削除済、botユーザー、Slackbotを除く
          if (!member.deleted && !member.is_bot && member.id !== "USLACKBOT") {
            slack_members[member.id] = member.name;
          }
        }

        var channelsResp = this.requestSlackAPI('conversations.list');
        var team_name = "team"
        var cch = channelsResp.channels
        for (var _i = 1, a = 9; _i < a; _i++) {
            var ch = cch.filter(function(item, index){
                  if (item.name == team_name+String(_i)) return true;
            });
            this.importChannelHistoryDelta(ch[0], team_name+String(_i), target_year);
        }
    };  

    SlackChannelHistoryLogger.prototype.importChannelHistoryDelta = function (ch, team_name, target_year) {
        var messages = this.loadMessagesBulk(ch, {}, target_year);
        var post_user = {};

        var options = {
          "method" : "get",
          "contentType": "application/x-www-form-urlencoded",
          "payload" : {
            "token": API_TOKEN,
            "channel":ch.id
          }
        };
        var m_url = "https://slack.com/api/conversations.members";
        var m_response = UrlFetchApp.fetch(m_url, options);
        var members = JSON.parse(m_response).members;

        members.forEach(function (member){
          post_user[member] = [0,0];
        });

      if(messages != "err"){
        messages.forEach(function (msg) {
          var date = new Date(+msg.ts * 1000);
          var reactions = msg.reactions ? msg.reactions : "";
          var postUser = msg.user ? msg.user : "";
          var m_year = Utilities.formatDate(date, timezone, 'yyyy-MM-dd');

          if(postUser !== "" && m_year >= target_year){
            post_user[postUser][0]++;
          }

          if(reactions !== "" && m_year >= target_year){
            reactions.forEach(function (reaction) {
              var name = reaction.name;
              if (stamps[name]) {
                stamps[name] = stamps[name] + reaction.count;
              } else {
                stamps[name] = reaction.count;
              }

                reaction.users.forEach(function (send_user) {
                  post_user[send_user][1]++;
                });
            });
          }
        });

        delete post_user.U044CAXMG8K;
        delete post_user.U045CS712Q2;
        all_post_user[team_name] = post_user;
      }
    };

    SlackChannelHistoryLogger.prototype.loadMessagesBulk = function (ch, options, target_year) {
        var _this = this;
        if (options === void 0) { options = {}; }
        var messages = [];
        options['limit'] = HISTORY_COUNT_PER_PAGE;
        options['channel'] = ch.id;
        var loadSince = function (cursor) {
          if (cursor) {
              options['cursor'] = cursor;
          }
          var resp = _this.requestSlackAPI('conversations.history', options);
          if(resp != "err"){
            messages = resp.messages.concat(messages);
          }
          return resp;
        };
        var resp = loadSince();
        var page = 1;
        // テスト用、取得数絞ってやる
        // while (resp.has_more && page <= MAX_HISTORY_PAGINATION) {
        // ページがあって、メッセージが対象年なら次ページを取得
        // HACK: 最初のメッセージで判定してる関係で無駄に1回多く叩いてるけど
        while (resp.has_more && target_year <= Utilities.formatDate(new Date(+resp.messages[resp.messages.length-1].ts * 1000), timezone, 'yyyy-MM-dd')) {          
            resp = loadSince(resp.response_metadata.next_cursor);
            page++;
        }
        return messages;
    };
    return SlackChannelHistoryLogger;
})();
