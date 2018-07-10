
var aCode = 65;
var menuRow = "2";
var calcStartRow = 2;
var calcStartCol = 3;

function onOpen() {

    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu('Calc Wcs Competition Rank');
    menu.addItem('Prelim Format', 'readyPrelim').addToUi();
    menu.addItem('Calc Prelim', 'calcPrelim').addToUi();
    menu.addItem('Final Format', 'readyFinal').addToUi();
    menu.addItem('Calc Filnal', 'calcFinal').addToUi();
}

function initJudge() {

    var judgeCount = parseInt(Browser.inputBox("input judge count"));
    if(judgeCount == "") {
        judgeCount = 4;
    }
    PropertiesService.getScriptProperties().setProperty('judgeCount', judgeCount);
    return judgeCount;
}

function readyPrelim() {

    var judgeCount = initJudge();

    var sheet = SpreadsheetApp.getActiveSheet();
    var index = printBasicMenu(sheet, judgeCount);
    PropertiesService.getScriptProperties().setProperty('totalColumn', String.fromCharCode(index));

    sheet.getRange(String.fromCharCode(index++) + menuRow).setHorizontalAlignment("center").setValue("Total");
    sheet.getRange(String.fromCharCode(index++) + menuRow).setHorizontalAlignment("center").setValue("Chief");

    Browser.msgBox("Add Point and Next Calc Prelim");
}

function readyFinal() {

    var judgeCount = initJudge();

    var sheet = SpreadsheetApp.getActiveSheet();
    var index = printBasicMenu(sheet, judgeCount);

    sheet.getRange(String.fromCharCode(index++) + menuRow).setHorizontalAlignment("center").setValue("Chief");
    PropertiesService.getScriptProperties().setProperty('rankStart', index);

    Browser.msgBox("Add Point and Next Calc Final");
}

function printBasicMenu(sheet, judgeCount) {

    var index = aCode;
    sheet.getRange(String.fromCharCode(index++) + menuRow).setHorizontalAlignment("center").setValue("Rank");
    sheet.getRange(String.fromCharCode(index++) + menuRow).setHorizontalAlignment("center").setValue("#");
    sheet.getRange(String.fromCharCode(index++) + menuRow).setHorizontalAlignment("center").setValue("name");

    for(var i = aCode; i < aCode + judgeCount; i++) {

        sheet.getRange(String.fromCharCode(index++) + menuRow).setHorizontalAlignment("center").setValue(String.fromCharCode(i));
    }

    return index;
}

// 예선 계산식
// 점수가 낮을 수록 순위가 높다.
// 점수가 같을 경우 cheif 점수로 순위 결정 한다.
// cheif 의 점수는 넣을수도 넣지 않을 수도 있다.
// 저지 마다 점수 배정(ex) 0.8) 이 있을 수도 있다지만, 우린 알수가 없다.
function calcPrelim() {

    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getDataRange();
    var values = range.getValues();
    var judgeCount = parseInt(PropertiesService.getScriptProperties().getProperty('judgeCount'));
    var totalColumn = PropertiesService.getScriptProperties().getProperty('totalColumn');

    sheet.hideRows(1, 2);

    calcPrelimTotal(sheet, values, judgeCount,totalColumn);
    range.sort(calcStartCol + judgeCount + 1);
    unhide(sheet);
    orderPrelim(sheet, values);
    sheet.setActiveRange(sheet.getRange("A1"));

    Browser.msgBox("Calc Prelim End. Check it!!");
}

function calcPrelimTotal(sheet, values, judgeCount, totalColumn) {

    for (var row = calcStartRow; row < values.length; row++) {

        var sum = 0;
        for( var col = calcStartCol; col < calcStartCol + judgeCount; col++) {
            sum += values[row][col]
        }
        var rowCode = totalColumn + (row+1);
        sheet.getRange(rowCode).setValue(sum);
    }
}

function orderPrelim(sheet, values) {

    for (var i = 1; i < values.length-1; i++) {

        sheet.getRange("A"+(i+2)).setValue(i);
    }
}

function unhide(sheet) {

    var hideRange = sheet.getRange("A1");
    sheet.unhideRow(hideRange);
    hideRange = sheet.getRange("A2");
    sheet.unhideRow(hideRange);
}

// 결선 계산식
// 점수가 낮은 순위를 받은 것을 더하여 과반(5-3, 7-4 등)을 먼저 획득하면 높은 등수를 차지한다.
// 같은 등수에서 같이 끝날 경우 더하기에 적은 쪽이 승리한다.
// 같은 등수에서 같은 점수로 끝날 경우, 점수를 다 더해서 순위를 결정한다.
// 다 같으면 다음 점수가 낮은 쪽이 승리한다.
// 이것 까지 다 같으면 더 많은 저지가 낮은 등수를 준 쪽이 승리한다.
// 저지 마다 점수 배정(ex) 0.8) 이 있을 수도 있다지만, 우린 알수가 없다.
function calcFinal() {

    var sheet = SpreadsheetApp.getActiveSheet();
    var range = sheet.getDataRange();
    var values = range.getValues();
    var judgeCount = parseInt(PropertiesService.getScriptProperties().getProperty('judgeCount'));
    var rankStart = parseInt(PropertiesService.getScriptProperties().getProperty('rankStart'));

    var rowCount = values.length - calcStartRow;

    printPointArray(sheet, rowCount, rankStart);
    printRankCount(sheet, values, judgeCount, rowCount, rankStart);
}

function printPointArray(sheet, rowCount, rankStart) {

    var index = 1;
    var col = rankStart;
    while( index <= rowCount) {

        sheet.getRange(String.fromCharCode(col++) + menuRow).setHorizontalAlignment("center").setValue("1-" + index++);
    }
}

function printRankCount(sheet, values, judgeCount, rowCount, rankStart) {

    for (var row = calcStartRow; row < rowCount + calcStartRow; row++) {

        var index = 1;
        var col = rankStart;
        while( index <= rowCount) {

//      Browser.msgBox("index=" + index + " row=" + row);

            var count = getRankCount(sheet, values, judgeCount, row, index);
            if(count > 0) {
                sheet.getRange(String.fromCharCode(col) + (row+1)).setHorizontalAlignment("center").setValue(count);
            }
            col++;
            index++;
        }
    }
}

function getRankCount(sheet, values, judgeCount, row, index) {

    var count = 0;
    for( var col = calcStartCol; col < calcStartCol + judgeCount; col++) {

//    Browser.msgBox("col=" + col + " row=" + row + ", values[row][col]=" + values[row][col]);

        if(values[row][col] >= index) {
            count++;
        }
    }
//  Browser.msgBox("count=" + count);
    return count;
}































