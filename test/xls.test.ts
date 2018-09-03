import * as fs from 'fs'
import { uploadSpreadsheet } from '../spreadsheetUpload';
import 'mocha'
import 'typescript'

const TEST_BOOK_1_LOCATION = "./Book1.xlsx"
describe("xlsx sheet reading", () => {
    let buffer: Buffer = fs.readFileSync(TEST_BOOK_1_LOCATION)
    let arrayBuf = buffer.buffer
    let file: File = new File([arrayBuf], "Book1.xlsx")
    it("should read from test book1", () => {
        console.log(uploadSpreadsheet(file))
    })
})