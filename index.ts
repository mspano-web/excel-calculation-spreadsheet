import { cursorTo } from 'readline';
import * as XLSX from 'xlsx';

interface Movement {
    CUSTOMER: string;
    'REGISTRATION-DATE': string;
    AMOUNT: number;
}

interface SummaryRow {
    CUSTOMER: string;
    'YEAR-MONTH': string;
    AMOUNT: number;
}

// -------------------------------------------------------

function generateSummaryInfo(movements: Movement[]): SummaryRow[] {
    const summaryMap: Map<string, number> = new Map();

    movements.forEach(movement => {
        const date = new Date(movement['REGISTRATION-DATE']);
        const isoDate = date.toISOString(); // Convert date to an ISO-8601 string
        const yearMonth = isoDate.slice(0, 7); // Get the year and month in YYYY-MM format

        const key = `${movement.CUSTOMER}-${yearMonth}`;
        // If summaryMap.get(key) returns a value other than undefined, that value is taken as the result of the expression. 
        // If summaryMap.get(key) returns undefined, the expression evaluates to 0.
        summaryMap.set(key, (summaryMap.get(key) || 0) + movement.AMOUNT);
    });

    console.log("generateSummary - summaryMap; ", summaryMap)

    const summaryRows: SummaryRow[] = [];

    summaryMap.forEach((amount, key) => {
        // divides as follows: customer: A00001 rest: [ '2023', '12']
        const [customer, ...rest] = key.split('-'); 
        const yearMonth = rest.join('-');
        summaryRows.push({ CUSTOMER: customer, 'YEAR-MONTH': yearMonth, AMOUNT: amount });
    });

    return summaryRows;
}

// -------------------------------------------------------

function generateSummaryExcel(summaryRows: SummaryRow[]) {
    const finalFilename = 'summary.xlsx';

    // Create a new (empty) workbook in Excel. In the xlsx library, 
    //  a workbook is an object that can contain one or more worksheets
    const wb = XLSX.utils.book_new();

    // Additional Information
    const data = [
        ['SUMMARY'],
        ['FECHA:', getDateString()]
    ];

    // Add the summary data to the matrix
    summaryRows.forEach(row => {
        data.push([row.CUSTOMER, row['YEAR-MONTH'], row.AMOUNT.toString()]);
    });

    // Converts an array of arrays into an Excel spreadsheet. 
    // The abbreviation "aoa" stands for "array of arrays."
    const ws = XLSX.utils.aoa_to_sheet(data);

    // A spreadsheet is added to the Excel workbook.
    //      wb: It is the Excel workbook object to which you want to add the spreadsheet.
    //      ws: It is the spreadsheet that you want to add to the workbook.
    //      'summary': This is the name that will be given to the spreadsheet in the workbook.
    XLSX.utils.book_append_sheet(wb, ws, 'Summary');

    // Takes the workbook object (wb) and writes it to an Excel file with the name specified in finalFilename
    XLSX.writeFile(wb, finalFilename);

    console.log(`El archivo "${finalFilename}" se ha generado correctamente.`);
}

// -------------------------------------------------------

function getDateString(): string {
    const date = new Date();
    return `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}-${date.getDate().toString().padStart(2, '0')}`;
}

// ---------------------------------------------

function excelDateToISODate(excelDate: number): string {
    const date = new Date((excelDate - 25569 + 1) * 86400 * 1000); 
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    return `${year}-${month}-${day}`; 
}

// ---------------------------------------------

function main() {
    // open and read the spreadsheet
    // workbook contains all the information in the Excel file, including the spreadsheets.
    const workbook = XLSX.readFile('movements.xlsx');
    // The workbook SheetNames property is an array that contains the names of all the worksheets in the Excel file.
    // SheetNames[0] accesses the first element of the SheetNames array, that is, the name of the first worksheet in the workbook.
    const sheetName = workbook.SheetNames[0];
    // workbook.Sheets[sheet Name], does not give access to the worksheet corresponding to the name stored in the sheetName variable.
    // Gives direct access to the data in that particular spreadsheet so that we can manipulate or read it as needed in our code.
    const worksheet = workbook.Sheets[sheetName];
    // The variable worksheet contains the data for a specific worksheet in the Excel file. 
    // This data is represented in the form of an object, where the keys are the addresses of the cells and the values ​​are the contents of those cells.

    const movements: Movement[] = [];

    for (let cellAddress in worksheet) {
        // We make sure that we are only iterating over the actual cells of the worksheet and 
        //      not other properties that may be present in the worksheet object.

        if (!worksheet.hasOwnProperty(cellAddress)) continue;
        // It skips processing a special property on the worksheet object called !ref, 
        //      which indicates the range of cells used in the worksheet, for example: "A2:E12".

        if (cellAddress === '!ref') continue;

        // Returns the content of the cell corresponding to the cellAddress address
        // The cell variable contains the representation of a specific cell in the spreadsheet
        const cell = worksheet[cellAddress];

        // We use a regular expression to split the cell address (cellAddress) into its column part and its row part.
        // The match expression returns an array with the results of the regular expression match. 
        // By adding ! In the end, we are telling the TypeScript compiler that we trust the regular expression
        //       to match and not to return null or undefined. This is made possible using ! 
        //      because we know that the regex will always return a result in this case since cellAddress 
        //      will always be a valid cell address.
        const [column, row] = cellAddress.match(/[a-zA-Z]+|[^a-zA-Z]+/g)!;
    
        let value: string | number = cell.v;

        // The date comes in numeric ('n') type (cell.t) and is formatted as "YYYY-MM-DD"
        if (cell.t === 'n' && cell.w && cell.w.match(/^\d{4}-\d{2}-\d{2}$/)) {
            value = excelDateToISODate(value as number);
        }
        
        if (row === '1') continue; // Ignorar la fila de encabezado

        // The output record is built
        if (column === 'A') movements.push({ CUSTOMER: value as string, 'REGISTRATION-DATE': '', AMOUNT: 0 });
        if (column === 'B') movements[movements.length - 1]['REGISTRATION-DATE'] = value as string;
        if (column === 'C') movements[movements.length - 1].AMOUNT = value as number;
    
    }
    
    console.log("main - movements: (", movements, ")");
    const summaryRows = generateSummaryInfo(movements);
    generateSummaryExcel(summaryRows);
}

// ---------------------------------------------


main();

