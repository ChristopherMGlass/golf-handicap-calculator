import * as XLSX from 'xlsx';
export function uploadSpreadsheet(file) {
    var reader = new FileReader();
    var fileNameParts = file.name.split('.');
    var fileExtension = fileNameParts[fileNameParts.length - 1];
    switch (fileExtension.toLowerCase()) {
        case "csv":
            reader.onload = parseCSV;
            break;
        case "xlsx":
        case "xls":
            reader.onload = parseSpreadSheet;
            break;
        default:
            throw "Unhandled File Type";
    }
}
function parseCSV(event) {
    return;
}
function parseSpreadSheet(event) {
    var innerReader = event.target;
    var result = innerReader.result;
    var wrkBk;
    try {
        wrkBk = XLSX.read(result);
    }
    catch (e) {
        console.error("improper file tpye");
        throw ("FileType Exception");
    }
    wrkBk.SheetNames.forEach(function (name) {
        readSheet(wrkBk.Sheets[name]);
    });
}
function nextCol(column) {
    var length = column.length;
    var charCode = column.charCodeAt(length - 1);
    if (charCode == 90) {
        return column.slice(0, -1) + "AA";
    }
    else {
        return column.slice(0, -1) + String.fromCharCode(charCode++);
    }
}
/**
 * Gets the previous Excel style column
 * @param column :string - current column letter
 * @returns previous column letter or null if not found
 */
function prevCol(column) {
    var length = column.length;
    var charCode = column.charCodeAt(length - 1);
    if (charCode == 65) {
        if (length == 1) {
            return null;
        }
        return column.slice(0, -2) + 'Z';
    }
    else {
        return column.slice(0, -1) + String.fromCharCode(charCode--);
    }
}
//todo: expand to use vector
function readScoreFromSheet(sheet, address) {
    var scoreCount = 0;
    var scores = [];
    var orign = sheet[address.c + address.r];
    var leftCell = orign; //left looking cell
    var bottomCell = orign; //downward looking cell
    // look left for scores
    while (leftCell && (leftCell.t == 'n' || leftCell.t == 's') && scoreCount < 18) {
        scoreCount++;
        scores[scoreCount] = +leftCell.v; //ensure number if put in as string
        scores[0] += scores[scoreCount];
        leftCell = sheet[nextCol(address.c) + address.r];
    }
    if (scoreCount != 18 || !leftCell) {
        //check down for scores
        scoreCount = 0;
        while (bottomCell && (bottomCell.t == 'n' || bottomCell.t == 's') && scoreCount < 18) {
            scoreCount++;
            scores[scoreCount] = +bottomCell.v; //ensure number if put in as string
            scores[0] += scores[scoreCount];
            bottomCell = sheet[address.c + (address.r + 1)];
        }
    }
    if (scoreCount == 18) {
        return scores;
    }
    else {
        return null;
    }
}
function findAnchors(range, sheet) {
    var start = range.split(':')[0];
    var end = range.split(':')[1];
    var endRow = +end.split("")[1];
    var endCol = end.split("")[0];
    var col = start.split("")[0];
    var row = +start.split("")[1];
    var scoreSectionAnchors = {};
    while (col.charCodeAt(0) < endCol.charCodeAt(0)) {
        while (row < endRow) {
            var cell = sheet[col + row];
            if (cell.t == 's') { //if value type is string
                var legend = HOLE_LEGEND_LABEL_REGEX.test(cell.v);
                if (legend) {
                    var cursor = sheet[col + (row + 1)];
                    var cursorRow = row + 1;
                    var cursorCol = col;
                    if (cursor.t == 's' && HOLE_LABEL_REGEX.test(cursor.v)) {
                        while (cursor.t == 's' && !FIRST_HOLE_REGEX.test(cursor.v)) {
                            //check down
                            cursorRow++;
                            cursor = sheet[cursorCol + cursorRow];
                        }
                        if (cursor.t !== 's') {
                            cursor = sheet[col + (row + 1)];
                            while (cursor.t == 's' && !FIRST_HOLE_REGEX.test(cursor.v)) {
                                //check left
                                cursorCol = prevCol(cursorCol);
                                if (!cursorCol) {
                                    break;
                                }
                                cursor = sheet[cursorCol + (row + 1)];
                            }
                        }
                        else {
                            if (scoreSectionAnchors[cursorCol + cursorRow]) {
                                //anchor already exists - This is a stub for future behavior - for now it does nothing if a duplicant is found
                                //TODO- handle existing case
                            }
                            else {
                                //Assumes columns represent games starting to the right of th anchor
                                scoreSectionAnchors[cursorCol + cursorRow] = {
                                    next: { c: 1, r: 0 },
                                    label: cursor.v
                                };
                            }
                        }
                    }
                }
            }
            else if (cell.t == 'n') {
                //todo look for numerical patterns that look like golf scores
            }
            row++;
        }
        col = nextCol(col);
    }
    return scoreSectionAnchors;
}
function readSheet(sheet) {
    var range = sheet["!ref"];
    var anchors = findAnchors(range, sheet);
    var scores = {};
    Object.keys(anchors).forEach(function (key) {
        var cursor = key;
        while (sheet[cursor].t == 's' || sheet.cursor.t == 'n') {
            scores[anchors[key].label] = readScoreFromSheet(sheet, { c: key.slice(0, -1), r: +key.slice(-1) });
            var c = anchors[key].next.c;
            var r = anchors[key].next.r;
            var row = +key.slice(-1);
            var column = key.slice(0, -1);
            for (var i = 0; i < c; i++) {
                column = nextCol(column);
            }
            row += r;
            cursor = column + row;
        }
    });
    return scores;
}
/* If legend found:
    Look down
        if label found
             check for sequencial left
            else check for sequencial down
        else
            check for row value coresponding to hole
*/
var HOLE_LEGEND_LABEL_REGEX = RegExp("[h,H]ole( [n,N]umber)?");
var HOLE_LABEL_REGEX = RegExp("([h,H]ole )?#?(([1-9]{1}\b)|(1[0-8]\b))");
var FIRST_HOLE_REGEX = RegExp("([h,H]ole )?#?1\b");
//# sourceMappingURL=spreadsheetUpload.js.map