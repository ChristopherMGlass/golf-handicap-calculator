import * as fs from 'fs';
import { uploadSpreadsheet } from '../spreadsheetUpload';
import { File } from 'file-api';
var TEST_BOOK_1_LOCATION = "./Book1.xlsx";
describe("xlsx sheet reading", function () {
    var buffer = fs.readFileSync(TEST_BOOK_1_LOCATION);
    var arrayBuf = buffer.buffer;
    var file = new File({ buffer: [arrayBuf], name: "Book1.xlsx" });
    it("should read from test book1", function () {
        console.log(uploadSpreadsheet(file));
    });
});
//# sourceMappingURL=xls.test.js.map