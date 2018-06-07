//Global Declarations:
var ss = SpreadsheetApp.getActiveSpreadsheet();
var userRankingsPullSheet = ss.getSheetByName('User_Rankings_Pull');
var scoreAggregationSheet = ss.getSheetByName('Score_Aggregation');
var rankingTableSheet = ss.getSheetByName('Ranking_Table');
var companySheet = ss.getSheetByName('Company_Sheet');
var previousRankingTable = ss.getSheetByName('Previous_Ranking_Table');
var fundValueFomattingCSV = ss.getSheetByName('Fund Value FormattingCSV');
var tickerRow = 5

//Creates ID Row and Sums market Sentiment
function makeUserRankingsCleanSheet() {

  storePreviousDaysRanks();

  cleanUserInputData();

  scoreMovement();

  var aggregatedUserScore = setCurrentScore();

  rankString();

  compareScores();

  fundValue();

  mattPowerVote();

  makeUserROI(aggregatedUserScore);

  makeFundValue();

  fundValuesForUpload();

  userInformationForUpload();



  //Functions

  function storePreviousDaysRanks() {
    var dataRange = rankingTableSheet.getDataRange();
    var lastColumn = rankingTableSheet.getLastColumn();

    previousRankingTable.insertColumns(1, lastColumn + 2);
    dataRange.copyTo(previousRankingTable.getRange("A1"), SpreadsheetApp.CopyPasteType.PASTE_VALUES);

    dataRange.clear();
  }

  function cleanUserInputData() {

    scoreAggregationSheet.getDataRange().clear();
    var userRankingsPullLastRow = userRankingsPullSheet.getLastRow();
    var userRankingsPullLastColumn = userRankingsPullSheet.getLastColumn();

    var scoreAggregationSheetLastColumn = scoreAggregationSheet.getLastColumn();

    scoreAggregationSheet.getRange(1 , scoreAggregationSheetLastColumn + 1).setValue('User_ID');
    scoreAggregationSheet.getRange(1 , scoreAggregationSheetLastColumn + 2).setValue('Cumulative_Votes');

    for (var i = 2; i < (userRankingsPullLastRow + 1); i ++) {
      var tempUserScore = userRankingsPullSheet.getRange(i, 1).getValue();
      var tempUserData = userRankingsPullSheet.getRange(i ,3, 1, (userRankingsPullLastColumn - 3)).getValues();
      var sum = 0;
      for (var t = 0; t < tempUserData[0].length; t++) {
        sum += parseFloat(tempUserData[0][t]);
      }
      if (sum != 0) {
        scoreAggregationSheet.getRange(i, scoreAggregationSheetLastColumn + 1).setValue(tempUserScore);
        scoreAggregationSheet.getRange(i, scoreAggregationSheetLastColumn + 2).setValue(sum);
        continue;
      } else {
        scoreAggregationSheet.getRange(i, scoreAggregationSheetLastColumn + 1).setValue(tempUserScore);
        scoreAggregationSheet.getRange(i, scoreAggregationSheetLastColumn + 2).setValue(0);
        continue;
      }
    }
  }

  function scoreMovement() {

    var scoreAggregationSheetLastColumn = scoreAggregationSheet.getLastColumn();
    scoreAggregationSheet.getRange(1 , scoreAggregationSheetLastColumn + 1).setValue('Score_Movement');
    var cumulativeVotes = getCumulativeVotes();
    Logger.log(cumulativeVotes);

    //Functions
    function getCumulativeVotes() {
      var scoreAggregationSheetLastRow = scoreAggregationSheet.getLastRow();
      var marketMovement = companySheet.getRange(tickerRow + 4,1).getValue()
      var cumulativeVotes = scoreAggregationSheet.getRange(2,2,scoreAggregationSheetLastRow - 1, 1).getValues();

      for (var i = 0; i < cumulativeVotes.length; i++) {
        var scoreMovement = parseInt(cumulativeVotes[i][0]) * marketMovement;
        scoreAggregationSheet.getRange(i + 2, scoreAggregationSheetLastColumn + 1).setValue(scoreMovement);
      }
    }
  }

  function setCurrentScore() {

    var scoreAggregationSheetLastColumn = scoreAggregationSheet.getLastColumn();

    scoreAggregationSheet.getRange(1, scoreAggregationSheetLastColumn + 1).setValue('Cumulative_Score');

    var userRankingPullSheetLastRow = userRankingsPullSheet.getLastRow();
    var currentUserIDObject = userRankingsPullSheet.getRange(2, 1 , userRankingPullSheetLastRow - 1, 1).getValues();
    var currentUserIDArray = [];
    for (var i = 0; i < currentUserIDObject.length; i ++) {
      currentUserIDArray.push(currentUserIDObject[i][0]);
    }

    var previousRankingSheetLastRow = previousRankingTable.getLastRow();
    var previousUserIDObject = previousRankingTable.getRange(3, 1, previousRankingSheetLastRow - 2, 1).getValues();
    var previousUserIDArray = [];
    for (var i = 0; i < previousUserIDObject.length; i++) {
      previousUserIDArray.push(previousUserIDObject[i][0]);
    }
    var aggregatedUserScore = 0;
    for (var i = 0; i < currentUserIDArray.length; i++){
      for (var j = 0; j < previousUserIDArray.length; j++) {
         if (currentUserIDArray[i] == previousUserIDArray[j]) {
           Logger.log(currentUserIDArray[i] + ' ' + previousUserIDArray[j]);
           var userScore = previousRankingTable.getRange(j + 3, 6).getValue() + parseFloat(scoreAggregationSheet.getRange(i + 2, 3).getValue());
           aggregatedUserScore += userScore;
           scoreAggregationSheet.getRange(i + 2, scoreAggregationSheetLastColumn + 1).setValue(userScore);
           break;
         } else if (j == previousUserIDArray.length - 1){
           scoreAggregationSheet.getRange(i + 2, scoreAggregationSheetLastColumn + 1).setValue('0');
         } else {
           Logger.log(j);
           continue;
         }
      }
    }
    return aggregatedUserScore;
    Logger.log(aggregatedUserScore);
  }

  function rankString() {

    var scoreAggregationSheetLastColumn = scoreAggregationSheet.getLastColumn();

    scoreAggregationSheet.getRange(1, scoreAggregationSheetLastColumn + 1).setValue('Current_Rank');
    var lastRow = scoreAggregationSheet.getLastRow();

    for (var i = 1; i < lastRow; i ++) {
      var sheetPosition = i + 1;
      var name = ('=RANK(D'+ sheetPosition + ', D2:D' + lastRow + ')');
      scoreAggregationSheet.getRange(sheetPosition, scoreAggregationSheetLastColumn + 1).setValue(name);
    }
  }

  function compareScores() {

    var scoreAggregationSheetLastColumn = scoreAggregationSheet.getLastColumn();

    var scoreDifference = [];
    //Targeting current Rankings
    var currentScoresLastRow = scoreAggregationSheet.getLastRow();
    var currentScores = scoreAggregationSheet.getRange(2, 1 , (currentScoresLastRow - 1), 5).getValues();
    Logger.log(currentScores);
    //TargetingPreviousRankings
    var previousScoresLastRow = previousRankingTable.getLastRow();
    var previousScores = previousRankingTable.getRange(3, 1, previousScoresLastRow - 2, 3).getValues();
    Logger.log(previousScores);

    for (var i = 0; i < currentScores.length; i++) {
      Logger.log(currentScores[i]);
      for (var j = 0; j < previousScores.length; j++) {
        if (currentScores[i][0] == previousScores[j][0]) {
          var temp = parseInt(previousScores[j][2]) - parseInt(currentScores[i][4]);
          scoreDifference.push(temp);
          break;
        }
        if (j == (previousScores.length - 1)) {
          scoreDifference.push('0');
        }
      }
    }
    var binaryScoreDifference = makeBinary(scoreDifference);
    scoreAggregationSheet.getRange(1, scoreAggregationSheetLastColumn + 1).setValue('Ranking_Change;');

    for (var i = 0; i < binaryScoreDifference.length; i++) {
      scoreAggregationSheet.getRange(i + 2, scoreAggregationSheetLastColumn + 1).setValue(binaryScoreDifference[i]);
    }
  }

  function makeBinary(array) {

    var newArray = [];
    for (i = 0; i < array.length; i ++) {
      if (array[i] >= 1) {
        newArray.push('1');
        continue;
      }
      if (array[i] <= -1){
        newArray.push('-1');
        continue;
      }
      else {
        newArray.push('0');
        continue;
      }
    }
    return newArray;
  }

  function fundValue() {

    var scoreAggregationSheetLastColumn = scoreAggregationSheet.getLastColumn();
    var currentFundScore = fundValueFomattingCSV.getRange(1,2).getValue();

    scoreAggregationSheet.getRange(1, scoreAggregationSheetLastColumn + 1).setValue('Fund Value');

    var lastRow = scoreAggregationSheet.getLastRow();
    var currentRankObject = scoreAggregationSheet.getRange(2, scoreAggregationSheetLastColumn - 1, lastRow - 1, 1).getValues();
    Logger.log(currentRankObject);
    var currentRankArray = [];
    var totalScore = 0;
    for (var i = 0; i < currentRankObject.length; i++) {
      currentRankArray.push(1 / currentRankObject[i][0]);
      totalScore += parseFloat(1 / currentRankObject[i][0]);
    }

    Logger.log(currentRankArray);
    Logger.log(totalScore);
    for (var i = 0; i < currentRankArray.length; i++) {
      var temp = (currentRankArray[i] / totalScore) * currentFundScore;
      scoreAggregationSheet.getRange(i + 2, scoreAggregationSheetLastColumn + 1).setValue(temp);
    }
  }

  function mattPowerVote() {
      var scoreAggregationSheetLastColumn = scoreAggregationSheet.getLastColumn();
      var scoreAggregationSheetLastRow = scoreAggregationSheet.getLastRow();

      scoreAggregationSheet.getRange(1, scoreAggregationSheetLastColumn +  1).setValue('matt_power_vote');

      for (var i = 2; i < scoreAggregationSheetLastRow + 1; i++) {
        var tempValue = scoreAggregationSheet.getRange(i, scoreAggregationSheetLastColumn - 3).getValue();
        Logger.log(tempValue)
        var valueToBeSet = parseFloat(tempValue) * 100;

        Logger.log(valueToBeSet);
        scoreAggregationSheet.getRange(i, scoreAggregationSheetLastColumn + 1).setValue(parseInt(valueToBeSet));
      }
  }

  function makeUserROI(aggregatedUserScore){
    var scoreAggregationSheetLastColumn = scoreAggregationSheet.getLastColumn();
    var scoreAggregationSheetLastRow = scoreAggregationSheet.getLastRow();

    //var aggregatedUserScore = 30.74637 //for unit testing only

    scoreAggregationSheet.getRange(1, scoreAggregationSheetLastColumn +  1).setValue('user_ROI');

    // TODO: This function is very similar to the one above so should be refactored in accordance with DRY princples

    var userContributionScores = [];

    for (var i = 2; i < scoreAggregationSheetLastRow + 1; i++) {
      var tempValue = scoreAggregationSheet.getRange(i, scoreAggregationSheetLastColumn - 4).getValue();

      var contributionPerUser = parseFloat(tempValue) / aggregatedUserScore;
      userContributionScores.push(contributionPerUser);
    }
    var average = getAverage(userContributionScores);
    var stDev = getStandardDeviation(userContributionScores);

    for (var i = 0; i < userContributionScores.length; i++) {
      var formatString = '=NORMDIST(' + userContributionScores[i] + ',' + average + ',' + stDev + ',TRUE)' + '* 0.079';
      scoreAggregationSheet.getRange(i + 2, scoreAggregationSheetLastColumn + 1).setValue(formatString);
    }
  }

  function getAverage(array){
    var sum = array.reduce(function(sum, value){
      return sum + value;
    }, 0);

    var avg = sum / array.length;
    return avg;
  }

  function getStandardDeviation(values){
    var avg = getAverage(values);

    var squareDiffs = values.map(function(value){
      var diff = value - avg;
      var sqrDiff = diff * diff;
      return sqrDiff;
    });

    var avgSquareDiff = getAverage(squareDiffs);

    var stdDev = Math.sqrt(avgSquareDiff);
    return stdDev;
  }

  function makeFundValue() {

    var scoreAggregationSheetLastRow = scoreAggregationSheet.getLastRow();
    var scoreAggregationSheetLastColumn = scoreAggregationSheet.getLastColumn();
    var currentFundScore = fundValueFomattingCSV.getRange(1,2).getValue();

    var cumulativeVotes = scoreAggregationSheet.getRange(2, scoreAggregationSheetLastColumn - 7, scoreAggregationSheetLastRow - 1, 1).getValues();
    Logger.log(cumulativeVotes);

    var fundValue = scoreAggregationSheet.getRange(2, scoreAggregationSheetLastColumn - 2, scoreAggregationSheetLastRow - 1, 1).getValues();
    Logger.log(fundValue);

    var cumulativeMarketMovement = companySheet.getRange(tickerRow + 4, 1).getValue();
    Logger.log(cumulativeMarketMovement);

    var unallocatedFunds = 0;
    var negativeAllocation = 0;
    var negativeBase = 0;
    var positiveAllocation = 0;
    var positiveBase = 0;

    for (var i = 0; i < cumulativeVotes.length; i ++) {
      if (cumulativeVotes[i][0] == '0') {
        Logger.log(cumulativeVotes[i][0]);
        Logger.log(fundValue[i][0]);
        unallocatedFunds += fundValue[i][0];
        continue;
      } else if (cumulativeVotes[i][0] > 0) {
        Logger.log(cumulativeVotes[i][0]);
        Logger.log(fundValue[i][0]);
        positiveAllocation += (fundValue[i][0] * cumulativeVotes[i][0]);
        positiveBase += fundValue[i][0];
        continue;
      } else if (cumulativeVotes[i][0] < 0) {
        Logger.log(cumulativeVotes[i][0]);
        Logger.log(fundValue[i][0]);
        negativeAllocation += (fundValue[i][0] * cumulativeVotes[i][0]);
        negativeBase += fundValue[i][0];
        continue;
      } else {
        Logger.log(cumulativeVotes[i][0]);
        continue;
      }
    }
    Logger.log(unallocatedFunds);
    Logger.log(negativeAllocation);
    Logger.log(positiveAllocation);

    var lastColumn = companySheet.getLastColumn();

    companySheet.getRange(tickerRow + 10, 1, 4, lastColumn).clear({contentsOnly: true});

    companySheet.getRange(tickerRow + 6, 2).setValue(currentFundScore);
    companySheet.getRange(tickerRow + 10, 1).setValue(new Date());
    companySheet.getRange(tickerRow + 11, 1).setValue('Positive_Tilt');
    companySheet.getRange(tickerRow + 11, 2).setValue(positiveAllocation);
    companySheet.getRange(tickerRow + 11, 3).setValue(positiveBase);
    companySheet.getRange(tickerRow + 12, 1).setValue('Unallocated_Funds');
    companySheet.getRange(tickerRow + 12, 2).setValue(unallocatedFunds);
    companySheet.getRange(tickerRow + 13, 1).setValue('Negative_Tilt');
    companySheet.getRange(tickerRow + 13, 2).setValue(negativeAllocation);
    companySheet.getRange(tickerRow + 13, 3).setValue(negativeBase);
  }

  function fundValuesForUpload() {
    fundValueFomattingCSV.insertRows(1, 3)

    var fundValueName = 'Fund_Value';
    var fundValueFigure = companySheet.getRange(tickerRow + 7, 2).getValue();

    fundValueFomattingCSV.getRange(1,1).setValue(fundValueName);
    fundValueFomattingCSV.getRange(1,2).setValue(fundValueFigure);

    var fundValueChangeName = 'Fund_Change';
    var fundValueChangeFigure = companySheet.getRange(tickerRow + 8, 2).getValue();
    Logger.log(fundValueChangeFigure);
    var fundValueChangeFigureTwoDecimalPlaces = (fundValueChangeFigure * 100).toFixed(2); //The output figure will be expressed as a percentage to two percentage points, to interact with Matt's system

    fundValueFomattingCSV.getRange(2,1).setValue(fundValueChangeName);
    fundValueFomattingCSV.getRange(2,2).setValue(fundValueChangeFigureTwoDecimalPlaces);

    var marketChangeName = 'Market Change';
    var marketChangeValue = companySheet.getRange(tickerRow + 4, 1).getValue();
    var marketChangeValueTwoDecimalPlaces = (marketChangeValue * 100).toFixed(2)

    fundValueFomattingCSV.getRange(2,4).setValue(marketChangeName);
    fundValueFomattingCSV.getRange(2,5).setValue(marketChangeValueTwoDecimalPlaces);
  }

  function userInformationForUpload() {
    var userIDDescriptor = 'user_id';
    var userNameDescriptor = 'full_name';
    var userRankDescriptor = 'rank';
    var powerVoteDescriptor = 'power_vote';
    var rankStausDescriptor = 'rank_Status';
    var cumulativeScoreDescriptor = 'cumulative_Score';

    // TODO: Preferred Format to Integrate with Frontend
    // var sectorFocusDescriptor = 'sector_focus';
    // var roiWeekDescriptor = 'roi_1w';
    // var roiMonthDescriptor = 'roi_1m';
    // var roiSixmonthDescriptor = 'roi_6m';
    // var roiAllDescriptor = 'roi_all';
    // var bullBearDescriptor = 'bull_bear';

    rankingTableSheet.getRange(1,1).setValue(new Date());
    rankingTableSheet.getRange(2,1).setValue(userIDDescriptor);
    rankingTableSheet.getRange(2,2).setValue(userNameDescriptor);
    rankingTableSheet.getRange(2,3).setValue(userRankDescriptor);
    rankingTableSheet.getRange(2,4).setValue(powerVoteDescriptor);
    rankingTableSheet.getRange(2,5).setValue(rankStausDescriptor);
    rankingTableSheet.getRange(2,6).setValue(cumulativeScoreDescriptor);

    // TODO: Preferred Format to integrate with Frontend
    // rankingTableSheet.getRange(2,7).setValue(sectorFocusDescriptor);
    // rankingTableSheet.getRange(2,8).setValue(roiWeekDescriptor);
    // rankingTableSheet.getRange(2,9).setValue(roiMonthDescriptor);
    // rankingTableSheet.getRange(2,10).setValue(roiSixmonthDescriptor);
    // rankingTableSheet.getRange(2,11).setValue(roiAllDescriptor);
    // rankingTableSheet.getRange(2,11).setValue(bullBearDescriptor);

    var dataRange = scoreAggregationSheet.getDataRange().getValues();
    Logger.log(dataRange);

    var userRankingsPullLastRow = userRankingsPullSheet.getLastRow();
    var userFullName = userRankingsPullSheet.getRange(2,2, userRankingsPullLastRow - 1, 1).getValues();
    Logger.log(userFullName);


    for (var i = 1; i < dataRange.length; i++) {
        rankingTableSheet.getRange(i + 2, 1).setValue(dataRange[i][0]);
        rankingTableSheet.getRange(i + 2, 2).setValue(userFullName[i]);
        rankingTableSheet.getRange(i + 2, 3).setValue(dataRange[i][4]);
        rankingTableSheet.getRange(i + 2, 4).setValue(dataRange[i][7]);
        rankingTableSheet.getRange(i + 2, 5).setValue(dataRange[i][5]);
        rankingTableSheet.getRange(i + 2, 6).setValue(dataRange[i][3]);

    }
  }

}









//Updates the first three values in Company Sheet regarding Stock Current Values and changes
function marketMovements() {
  var companyTickers = getTickers();

  setPriceValuesInSheet(companyTickers);

  setPriceChangePercent(companyTickers);

  function getTickers() {
    var lastColumn = companySheet.getLastColumn();
    var companyTickers = companySheet.getRange(tickerRow, 2 , 1, lastColumn - 1).getValues();
    return companyTickers;
  }

  function setPriceValuesInSheet(companyTickers) {
    for (var i = 0; i < companyTickers[0].length; i++){
      var baseURL = "https://api.iextrading.com/1.0/stock/" + companyTickers[0][i] + "/time-series";
      var response = JSON.parse(UrlFetchApp.fetch(baseURL));
      var todayClosePrice = (response[response.length - 1].close);
      var yesterdayClosePrice = (response[response.length - 2].close);
      var columPosition = i + 2;
      companySheet.getRange(tickerRow - 3, columPosition).setValue(todayClosePrice);
      companySheet.getRange(tickerRow - 2, columPosition).setValue(yesterdayClosePrice);
    }
  }

  function setPriceChangePercent(companyTickers) {
    for (var i = 0; i < companyTickers[0].length; i++) {
      var baseURL = "https://api.iextrading.com/1.0/stock/" + companyTickers[0][i] + "/quote";
      var response = JSON.parse(UrlFetchApp.fetch(baseURL));
      var priceChangePercent = response.changePercent;
      var columPosition = i + 2;
      companySheet.getRange(tickerRow - 1, columPosition).setValue(priceChangePercent);
    }
  }
}

function marketValueAPI() {
  var companyTickers = getTickers();

  setMarketCapValues(companyTickers);

  function getTickers() {
    var lastColumn = companySheet.getLastColumn();
    var companyTickers = companySheet.getRange(tickerRow, 2 , 1, lastColumn - 1).getValues();
    return companyTickers;
  }

  function setMarketCapValues(companyTickers) {
    for (var i = 0; i < companyTickers[0].length; i++) {
      var baseURL = "https://api.iextrading.com/1.0/stock/" + companyTickers[0][i] + "/stats";
      var response = JSON.parse(UrlFetchApp.fetch(baseURL));
      var priceMarketCap = response.marketcap;
      var columPosition = i + 2;
      companySheet.getRange(tickerRow + 1, columPosition).setValue(priceMarketCap);
    }
  }
}


//Order so Far - ExternalFundValue
//CleanUser
//MarketValueAPI only needs to be run on first instance



//Output need to be:
//Spreadsheet with fund figures -
//User Voting Power, User Control of Fund

//Votes:
//New Users come in bottom
