import { Sheet, WorkSheet, CellAddress, CellObject } from "xlsx/types";
import { Anchors, AlphaNumericAddress, nextCol, colGreaterThan } from "./sheetUtils";

const DATE_LABEL_REGEX: RegExp = RegExp("^[D,d]ate$")
const COURSE_LABEL_REGEX: RegExp = RegExp("^([C,c]ourse)( [N,n]ame)?$")
const ESC_SCORE_LABEL: RegExp = RegExp("^([E,e][S,s][C,c] )?[S,s]core")
const RATING_LABEL: RegExp = RegExp("^([C,c]ourse )?[R,r]ating")
const SlOPE_LABEL: RegExp = RegExp("^([C,c]ourse )?[S,s]lope")
const TEE_LABEL: RegExp = RegExp("^[T,t]ees?( [M,m]arkers?)?")
export const LABELS = {
    date: DATE_LABEL_REGEX,
    courseLabel: COURSE_LABEL_REGEX,
    score: ESC_SCORE_LABEL,
    slope: SlOPE_LABEL,
    courseRating: RATING_LABEL,
    tees: TEE_LABEL
}
const REQUIRED_LABELS = ["score", "courseRating", "slope"]

function findAnchors(range: string, sheet: Sheet): Anchors {
    let start: string = range.split(':')[0]
    let end: string = range.split(':')[1]
    let endRow: number = +end.replace(/\D/g, '')
    let endCol: string = end.replace(/\d/g, '')
    let col: string = start.replace(/\d/g, '')
    let row: number = +start.replace(/\D/g, '')

    let handicapAnchors: Anchors = {}
    console.log("finding anchors")

    while (colGreaterThan(col, endCol)) {
        while (row < endRow) {
            let cell: CellObject = sheet[col + row]
            if (cell && cell.t == 's' && checkForField(<string>cell.v, { c: col, r: row })) {
                let right = checkForField(<string>sheet[nextCol(col) + row].v, { c: nextCol(col), r: row })
                let down = checkForField(<string>sheet[col + (row + 1)].v, { c: col, r: row + 1 })
                //assume only one anchor per sheet
                if (right) {
                    console.log("found right")
                    handicapAnchors[col + row] = { next: { c: 1, r: 0 } }
                    return handicapAnchors
                } else if (down) {
                    console.log("found down")
                    handicapAnchors[col + row] = { next: { c: 0, r: 1 } }
                    return handicapAnchors
                }
                console.log("not found")
            }
            row++
        }
        col = nextCol(col)
    }
    return handicapAnchors
}

export function readSheet(sheet: WorkSheet): handicapRound[] {
    let range: string = sheet["!ref"]

    let anchors: Anchors = findAnchors(range, sheet)
    let games: handicapRound[] = []
    console.log(anchors)
    Object.keys(anchors).forEach((key: string) => { //todo uses old structure - refactor
        let cursor: string = key
        let row: number = +key.replace(/\D/g, '')
        let col: string = key.replace(/\d/g, '')
        let next = anchors[key].next
        let labels = associateValues(sheet, cursor, next)
        if (!validateLabels(labels)) {
            throw "insufficient data found in sheet"
        }

        //Todo: do while?
        while (sheet[cursor] && (sheet[cursor].t == 's' || sheet[cursor].t == 'n')) {
            if (next.r) {
                //label are on the left side
                col = nextCol(col)
                cursor = col + row

                //TODO /CODE SMELL this can probable be refactored into while condition somehow
                if (!sheet[cursor]) {
                    break;
                }
                let game = {
                }
                Object.keys(labels).forEach(key => {
                    game[key] = sheet[col + labels[key]].w
                });
                // console.log("game found by column:",game)
                //labels has been validated to have required fields already
                games.push(<handicapRound>game)
            } else if (next.c) {
                //label are along the top
                row++
                cursor = col + row
                let game = {
                }
                //TODO /CODE SMELL this can probable be refactored into while condition somehow
                if (!sheet[cursor]) {
                    break;
                }
                Object.keys(labels).forEach(key => {
                    game[key] = sheet[labels[key] + row].w
                });
                // console.log("game found:",game)
                //labels has been validated to have required fields already
                games.push(<handicapRound>game)

            } else {
                throw "sheet structure invalid"
            }
            // console.log(sheet)
        }
    })
    return games
}

function checkForField(cellValue: string, cellAddress: AlphaNumericAddress): boolean {
    let allLabels = {
        date: DATE_LABEL_REGEX,
        course: COURSE_LABEL_REGEX,
        score: ESC_SCORE_LABEL,
        rating: RATING_LABEL,
        slope: SlOPE_LABEL,
        tee: TEE_LABEL
    }
    for (let idx in Object.keys(allLabels)) {
        let key = Object.keys(allLabels)[idx]
        if (allLabels[key].test(cellValue)) {
            console.log("feild", key, "found")
            return true
        }
    }
    return false
}
function validateLabels(labels) {
    REQUIRED_LABELS.forEach((key) => {
        if (!labels.hasOwnProperty(key)) {
            return false
        }
    })
    return true
}
/**
 * find the coresponding rows/columns for necessary data
 * @param sheet 
 * @param addr 
 * @param next 
 */
export function associateValues(sheet: Sheet, addr: string, next: CellAddress) {

    let cursor = sheet[addr]
    let col: string = addr.replace(/\d/g, '')
    let row: number = +addr.replace(/\D/g, '')
    if (next.c) {
        let columns = {
        }
        let cursorCol = col
        while (cursor && cursor.t == 's') {
            Object.keys(LABELS).forEach((key) => {
                if (LABELS[key].test(cursor.v)) {
                    columns[key] = cursorCol
                }
            })
            cursorCol = nextCol(cursorCol)
            cursor = sheet[cursorCol + row]
        }
        return columns
    } else if (next.r) {
        let rows = {
        }
        let cursorRow = row
        while (cursor && cursor.t == 's') {
            Object.keys(LABELS).forEach((key) => {
                if (LABELS[key].test(cursor.v)) {
                    rows[key] = cursorRow
                }
            })
        }
        cursorRow++
        cursor = sheet[col + cursorRow]
        return rows
    }
}

const HANDICAP_SCALER = 113
interface Course {
    name: string,
    rating: string,
    slope: string
}
export interface handicapRound {
    date?: string,
    course?: Course,
    tess?: string,
    rawScore?: number,
    scorecore: number,
    courseRating: number,
    slope: number
    scoreDiff?: number
}

function calcHandicap(rounds: handicapRound[]) {
    let numberScores: number = rounds.length;
    let adjustedGrossScore
    let courseRating
    let slopeRating
    let handcapDifferntial = (adjustedGrossScore - courseRating) * HANDICAP_SCALER / slopeRating
    if (numberScores < 20) {

    }
}