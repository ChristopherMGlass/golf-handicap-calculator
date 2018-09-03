import * as XLSX from 'xlsx'
import { Sheet, CellObject, WorkSheet, WorkBook, CellAddress } from 'xlsx'

export function uploadSpreadsheet(file: File) {
    let reader: FileReader = new FileReader()
    let fileNameParts: string[] = file.name.split('.')
    let fileExtension: string = fileNameParts[fileNameParts.length - 1]
    switch (fileExtension.toLowerCase()) {
        case "csv":
            reader.onload=parseCSV
            break;
        case "xlsx":
        case "xls":
            reader.onload=parseSpreadSheet
            break;
        default:
            throw "Unhandled File Type"

    }


}
function parseCSV(event: ProgressEvent) {
    return
}
function parseSpreadSheet(event: ProgressEvent) {
    let innerReader: FileReader = <FileReader>event.target
    let result:ArrayBuffer|string = innerReader.result
    let wrkBk: WorkBook
    try {
        wrkBk = XLSX.read(result)
    } catch (e) {
        console.error("improper file tpye")
        throw ("FileType Exception")
    }

    wrkBk.SheetNames.forEach((name: string) => {
        readSheet(wrkBk.Sheets[name])
    });
}

interface AlphaNumericAddress {
    c: string,
    r: number
}

function nextCol(column: string) {
    let length = column.length
    let charCode = column.charCodeAt(length - 1)
    if (charCode == 90) {
        return column.slice(0, -1) + "AA"
    } else {
        return column.slice(0, -1) + String.fromCharCode(charCode++)
    }
}
/**
 * Gets the previous Excel style column
 * @param column :string - current column letter
 * @returns previous column letter or null if not found
 */
function prevCol(column: string): string | null {
    let length = column.length
    let charCode = column.charCodeAt(length - 1)
    if (charCode == 65) {
        if (length == 1) {
            return null
        }
        return column.slice(0, -2) + 'Z'
    } else {
        return column.slice(0, -1) + String.fromCharCode(charCode--)
    }

}

//todo: expand to use vector
function readScoreFromSheet(sheet: Sheet, address: AlphaNumericAddress): number[] | null {
    let scoreCount: number = 0
    let scores: number[] = []
    let orign: CellObject = sheet[address.c + address.r]
    let leftCell: CellObject = orign; //left looking cell
    let bottomCell: CellObject = orign; //downward looking cell

    // look left for scores
    while (leftCell && (leftCell.t == 'n' || leftCell.t == 's') && scoreCount < 18) {
        scoreCount++
        scores[scoreCount] = +leftCell.v //ensure number if put in as string
        scores[0] += scores[scoreCount]
        leftCell = sheet[nextCol(address.c) + address.r]
    }
    if (scoreCount != 18 || !leftCell) {
        //check down for scores
        scoreCount = 0
        while (bottomCell && (bottomCell.t == 'n' || bottomCell.t == 's') && scoreCount < 18) {
            scoreCount++
            scores[scoreCount] = +bottomCell.v //ensure number if put in as string
            scores[0] += scores[scoreCount]
            bottomCell = sheet[address.c + (address.r + 1)]
        }
    }
    if (scoreCount == 18) {
        return scores
    } else {
        return null
    }

}
interface Anchors {
    [cellLabel: string]: {
        next: CellAddress, //vector indicating where to find next game
        label: string
    }
}
function findAnchors(range:string,sheet:Sheet):Anchors{
  
    let start: string = range.split(':')[0]
    let end: string = range.split(':')[1]
    let endRow: number = +end.split("")[1]
    let endCol: string = end.split("")[0]
    let col: string = start.split("")[0]
    let row: number = +start.split("")[1]

    let scoreSectionAnchors: Anchors = {}

    while (col.charCodeAt(0) < endCol.charCodeAt(0)) {
        while (row < endRow) {
            let cell: CellObject = sheet[col + row]
            if (cell.t == 's') { //if value type is string
                let legend = HOLE_LEGEND_LABEL_REGEX.test(<string>cell.v)
                if (legend) {
                    let cursor: CellObject = sheet[col + (row + 1)]
                    let cursorRow: number = row + 1
                    let cursorCol: string = col
                    if (cursor.t == 's' && HOLE_LABEL_REGEX.test(<string>cursor.v)) {
                        while (cursor.t == 's' && !FIRST_HOLE_REGEX.test(<string>cursor.v)) {
                            //check down
                            cursorRow++
                            cursor = sheet[cursorCol + cursorRow]
                        }
                        if (cursor.t !== 's') {
                            cursor = sheet[col + (row + 1)]
                            while (cursor.t == 's' && !FIRST_HOLE_REGEX.test(<string>cursor.v)) {
                                //check left
                                cursorCol = prevCol(cursorCol)
                                if (!cursorCol) {
                                    break;
                                }
                                cursor = sheet[cursorCol + (row + 1)]
                            }
                        } else {
                            if (scoreSectionAnchors[cursorCol + cursorRow]) {
                                //anchor already exists - This is a stub for future behavior - for now it does nothing if a duplicant is found
                                //TODO- handle existing case
                            } else {
                                //Assumes columns represent games starting to the right of th anchor
                                scoreSectionAnchors[cursorCol + cursorRow] = {
                                    next: { c: 1, r: 0 },
                                    label: <string>cursor.v
                                }
                            }
                        }
                    }
                }
            } else if (cell.t == 'n') {
                //todo look for numerical patterns that look like golf scores
            }
            row++
        }
        col=nextCol(col)
    }
    return scoreSectionAnchors
}

function readSheet(sheet: WorkSheet) {
    let range: string = sheet["!ref"]
    let anchors:Anchors=findAnchors(range,sheet)
    let scores:{[label:string]:number[]}={}
    Object.keys(anchors).forEach((key:string)=>{
        let cursor=key
        while(sheet[cursor].t=='s' || sheet.cursor.t=='n'){
            scores[anchors[key].label]=readScoreFromSheet(sheet,{c:key.slice(0,-1),r:+key.slice(-1)})
            let c:number= anchors[key].next.c
            let r:number= anchors[key].next.r
            let row:number=+key.slice(-1)
            let column=key.slice(0,-1)
            for(let i=0;i<c;i++){
               column=nextCol(column)
            }
            row+=r
            cursor=column+row
        }
    })
    return scores
}
/* If legend found: 
    Look down
        if label found
             check for sequencial left
            else check for sequencial down
        else
            check for row value coresponding to hole
*/
const HOLE_LEGEND_LABEL_REGEX: RegExp = RegExp("[h,H]ole( [n,N]umber)?")
const HOLE_LABEL_REGEX: RegExp = RegExp("([h,H]ole )?#?(([1-9]{1}\b)|(1[0-8]\b))")
const FIRST_HOLE_REGEX: RegExp = RegExp("([h,H]ole )?#?1\b")