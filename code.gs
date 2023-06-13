function updateGraph() {
  var ss = SpreadsheetApp.openById("SpreadSheet ID");
  var sheet = ss.getSheetByName('Form Sheet Name'); // フォームの回答が記録されているシート名を指定
  var data = sheet.getDataRange().getValues();

  // 新しく追加：メールアドレスリストの取得
  var emailSheet = ss.getSheetByName(' EmailAddress Sheet Name'); // メールアドレスが記録されているシート名を指定
  var emailData = emailSheet.getRange('A:A').getValues(); // メールアドレスが記録されている列を指定
  var validEmails = emailData.map(function(row) { return row[0]; });

  //エントリーNo.を入力
  var entryNo =  1;
  
  //グラフを事前に削除
  var charts = ss.getSheetByName('Graph Sheet Name').getCharts();// グラフを表示するシート名を指定
  for (var j = 0; j < charts.length; j++) {
    sheet.removeChart(charts[j]);
  }

  var processedData = {};
  var allEntriesData = {};
  for(var i = data.length - 1; i >= 1; i--) {
    var row = data[i];
    var email = row[1];

    // 新しく追加：メールアドレスがリストに存在するか確認
    if (validEmails.indexOf(email) < 0) continue;

    var entry = row[2];
    var creativity = row[3];
    var interest = row[4];
    var effort = row[5];
    var clarity = row[6];
    var design = row[7];

    if(!allEntriesData.hasOwnProperty(email)) {
      allEntriesData[email] = {};
    }
    allEntriesData[email][entry] = [creativity, interest, effort, clarity, design];

    if(entry == entryNo && (!processedData.hasOwnProperty(email) || !processedData[email].hasOwnProperty(entry))) {
      if(!processedData.hasOwnProperty(email)) {
        processedData[email] = {};
      }
      processedData[email][entry] = [creativity, interest, effort, clarity, design];
    }
  }

  var categories = ['創造性', '面白さ', '工夫', '分かりやすさ', 'デザイン性']; //評価項目
  var graphData = [['カテゴリ', 'エントリーNo.' + entryNo + `の平均`, '全エントリーの平均']]; //表示するデータ群
  for(var i = 0; i < categories.length; i++) {
    var totalScoreForEntry = 0;
    var totalCountForEntry = 0;
    for(var email in processedData) {
      for(var entry in processedData[email]) {
        totalScoreForEntry += processedData[email][entry][i];
        totalCountForEntry++;
      }
    }

    var totalScoreForAllEntries = 0;
    var totalCountForAllEntries = 0;
    for(var email in allEntriesData) {
      for(var entry in allEntriesData[email]) {
        totalScoreForAllEntries += allEntriesData[email][entry][i];
        totalCountForAllEntries++;
      }
    }

    if(totalCountForEntry > 0 && totalCountForAllEntries > 0) {
      totalScoreForEntry /= totalCountForEntry;
      totalScoreForAllEntries /= totalCountForAllEntries;
      graphData.push([categories[i], totalScoreForEntry, totalScoreForAllEntries]);
    }
  }

  var graphSheet = ss.getSheetByName('Graph Sheet Name');
  if(!graphSheet) {
    graphSheet = ss.insertSheet('Graph Sheet Name');
  }
  graphSheet.clear();
  graphSheet.getRange(1, 1, graphData.length, 3).setValues(graphData);

  var chart = graphSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(graphSheet.getRange(2, 1, graphData.length, 3))
    .setPosition(8, 1, 0, 0)
    .setOption('title', '回答結果　エントリーNo.' + entryNo)
    .setOption('hAxis.title', 'カテゴリ')
    .setOption('vAxis.title', 'スコア')
    .setOption('series', {
      0: {
        labelInLegend: 'エントリーNo.' + entryNo + "の平均",
      },
      1: {
        labelInLegend: '全エントリーの平均',
      },
    })
    .build();

  graphSheet.insertChart(chart);


  //Scoresシートに結果を追加
  var scoresSheet = ss.getSheetByName('Scores Sheet Name');// 回答結果の平均を記録したいシート名を指定
  if(!scoresSheet) {
    scoresSheet = ss.insertSheet('Scores Sheet Name');
  }

  // Determine the row in which to write the scores. If the entry number already exists in column 1, overwrite it.
  // Otherwise, use the next empty row.
  var row = 2; // Start from row 2
  var found = false;
  while(scoresSheet.getRange(row, 1).getValue() !== "") {
    if(scoresSheet.getRange(row, 1).getValue() === entryNo) {
      found = true;
      break;
    }
    row++;
  }
  if(!found) {
    scoresSheet.getRange(row, 1).setValue(entryNo);
  }

  var scoresData = [];
  var totalScoreSum = 0;  // Sum of all totalScoreForEntry values for calculating the average

  for(var i = 0; i < categories.length; i++) {
    var totalScoreForEntry = 0;
    var totalCountForEntry = 0;
    for(var email in processedData) {
      for(var entry in processedData[email]) {
        totalScoreForEntry += processedData[email][entry][i];
        totalCountForEntry++;
      }
    }

    if(totalCountForEntry > 0) {
      totalScoreForEntry /= totalCountForEntry;
      scoresData.push(totalScoreForEntry);
      totalScoreSum += totalScoreForEntry;  // Add the totalScoreForEntry to the sum
    }
  }

  var averageTotalScore = totalScoreSum / scoresData.length;  // Calculate the average totalScoreForEntry
  scoresData.push(averageTotalScore);  // Add the average to the scoresData array

  for(var i = 0; i < scoresData.length; i++) {
    scoresSheet.getRange(row, i+2).setValue(scoresData[i]);  // i+2 to skip the category column
  }
}