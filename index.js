"use strict";
const csv = require("csvtojson");
const { Workbook } = require("exceljs");
const { stateCode, headers } = require("./config.json");
const path = require("path");
const { SingleBar,Presets } = require("cli-progress");
/**
 * @param {Workbook} sheet
 * WorkBook -> {@link Workbook}
 * A workbook is a file that contains one or more worksheets to help you organize data. 
 * You can create a new workbook from a blank workbook or a template.
 * WorkSheet -> 
 * A worksheet is a part of WorkBook
 */

const setHeader = (sheet) => {
    const mergedCol = ({ C1, C2, value }) => {
        sheet.mergeCells(`${C1}:${C2}`);
        sheet.getCell(C1).value = value;
        sheet.getCell(C1).alignment = { horizontal: "center" };
    }
    headers.forEach(value => { mergedCol(value); })
    let assciiCode = 65,
        visit = { 0: false, 1: false };
    do {
        const chr = String.fromCharCode(assciiCode);
        const cell = sheet.getCell(`${chr}:2`);
        cell.style = { alignment: { horizontal: "center" } };
        if (chr == "A") {
            cell.value = "State";
        }
        else if (chr == "B") {
            cell.value = "District";
        }
        else {
            if (!visit[0]) {
                cell.value = "Male";
                visit[0] = true;
            }
            else if (!visit[1]) {
                cell.value = "Female";
                visit[1] = true;
            }
            else {
                cell.value = "Total";
                visit = { 0: false, 1: false }
            }
        }
    } while ((++assciiCode) <= 81);
};
/**
 * 
 * @param {String} filepath 
 * @returns {*}
 * convert CSV to JSON 
 * validate the result
 * arrange it using key-value pair
 */
const extractAndConvert = async (filepath) => {
    try {
        //convert csv to json
        const json_result = await csv({ headers: ["dist", "avg"], trim: true, nullObject: true }).fromFile(filepath)
        const fileresult = {};
        //validate json result and arrange it in the from of key value pair
        for (const json of json_result) {

            if (json["avg"] && json["avg"] !== "NA" && json["avg"] !== " ") {
                fileresult[json["dist"]] = json["avg"].split(" ")[0];
            }
        }
        return fileresult;
    } catch (error) {
        throw new Error(`Operation failure ${filepath} ${error.message}`);
    }
};
/**
 * 
 * @param {SingleBar} bar 
 * @param {Number} limit 
 * main function
 * limit is represented as how many row a worksheet can contain
 */
async function main(bar, limit = 300) {
    const workbook = new Workbook();
    workbook.created = new Date();
    workbook.creator = "Debadutta Panda"
    workbook.title = "Education"
    workbook.description = "avg performance of class5 and class8 student across all state";
    //create worksheet under a workbook
    const createWorkSheet = (sheetno) => {
        const workSheet = workbook.addWorksheet(`Sheet-${sheetno}`);
        setHeader(workSheet);
        return workSheet;
    }
    let workSheet = createWorkSheet(1),
        count = 3,
        sheetno = 1,
        insertRecord = 0;
    bar.start(617, 0);
    for (const code in stateCode) {
        const alldata = [
            await extractAndConvert(`./Sheets/class-5-language/${code}-Male-L.csv`),
            await extractAndConvert(`./Sheets/class-5-language/${code}-Female-L.csv`),
            await extractAndConvert(`./Sheets/class-5-language/${code}-Total-L.csv`),
            await extractAndConvert(`./Sheets/class-5-math/${code}-Male-M.csv`),
            await extractAndConvert(`./Sheets/class-5-math/${code}-Female-M.csv`),
            await extractAndConvert(`./Sheets/class-5-math/${code}-Total-M.csv`),
            await extractAndConvert(`./Sheets/class-5-EVS/${code}-Male-E.csv`),
            await extractAndConvert(`./Sheets/class-5-EVS/${code}-Female-E.csv`),
            await extractAndConvert(`./Sheets/class-5-EVS/${code}-Total-E.csv`),
            await extractAndConvert(`./Sheets/class-8-language/${code}-Male-L.csv`),
            await extractAndConvert(`./Sheets/class-8-language/${code}-Female-L.csv`),
            await extractAndConvert(`./Sheets/class-8-language/${code}-Total-L.csv`),
            await extractAndConvert(`./Sheets/class-8-math/${code}-Male-M.csv`),
            await extractAndConvert(`./Sheets/class-8-math/${code}-Female-M.csv`),
            await extractAndConvert(`./Sheets/class-8-math/${code}-Total-M.csv`),
        ]
        await Promise.all(alldata).then((value) => {
            for (const key in value[0]) {
                const row = [stateCode[code], key];
                for (let i = 0; i < alldata.length; i++) {
                    row.push(value[i][key]);
                }
                if (count <= limit) {
                    workSheet.insertRow(count, row, "i");
                    count++;
                }
                else {
                    //if the given limit is exceeded then it will create another worksheet and continue insertion
                    count = 3;
                    workSheet = createWorkSheet(++sheetno);
                    workSheet.insertRow(count, row, "i");
                    count++;
                }
                bar.update(++insertRecord);
            }
        })
    }
    await workbook.xlsx.writeFile("./Result/result.xlsx").then(value => {
        bar.stop();
        console.log("\n**********************************\n\n\tSuccessfully Done\t\t\n\n**********************************\n");
        console.log(`Please visit ${path.resolve("./Result/result.xlsx")}`)
    }).catch(err => {
        console.log("\n\nYou may view result.xlsx using other apps please closed it before you run this program\n");
        process.exit(0);
    })
}
const bar = new SingleBar({}, Presets.shades_classic);
main(bar);

