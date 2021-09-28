const xlsxj = require("xlsx-to-json");
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet("output");
var finalData = [];

// Convert Excel To Json
const convertExcelToJson = () => {
    return new Promise((resolve, reject) => {
        xlsxj({
            input: "input.xlsx",
            output: "output.json"
        }, function (err, result) {
            if (err) {
                console.error(err);
                reject(err)
            } else {

                outputData = result;

                // Count the occurance of genus

                let rez = []
                result.forEach(function (item) {
                    rez[item.genus] ? rez[item.genus]++ : rez[item.genus] = 1;
                });

                /*
                * Loop through the data to  find the occurance
                * Find the first digit of the occurance count
                * If occurance count is 93 then takecount will be the first letter i.e 9
                * If occurance is single digit like 4 then take count will be the 1
                * Take count is the final count for the number of data to be needs to be written into the output file
                */
                for (var key in rez) {
                    if (rez.hasOwnProperty(key)) {

                        let takeCount = 0;
                        if (rez[key] < 10) {

                            takeCount = 1;

                        } else if (rez[key] > 10 && rez[key] < 99) {
                            takeCount = parseInt(('' + rez[key]).slice(0, 1))

                        } else if (rez[key] > 99 && rez[key] < 999) {
                            takeCount = parseInt(('' + rez[key]).slice(0, 2))

                        }

                        let filterData = result.filter(function (ele) {

                            return ele.genus == key
                        })
                        // console.log("Filter data count --", filterData.length + "   take count ---  " + takeCount + " -----key-----> " + key);
                        for (let i = 0; i < takeCount; i++) {
                            finalData.push(filterData[i])
                        }
                    }
                }
                data = finalData.filter(function (element) {
                    return element !== undefined;
                });

                const headingColumnNames = Object.keys(data[0]);

                //Write Column Title in Excel file
                let headingColumnIndex = 1;

                headingColumnNames.forEach(heading => {

                    ws.cell(1, headingColumnIndex++)
                        .string(heading)


                })

                //Write Data in Excel file
                let rowIndex = 2;

                data.forEach(record => {

                    let columnIndex = 1;

                    Object.keys(record).forEach(columnName => {

                        ws.cell(rowIndex, columnIndex++)
                            .string(record[columnName])
                    });
                    rowIndex++;
                });
                wb.write("output.xlsx");

                // console.log(data)
                resolve(rez)
            }
        });
    })

}

convertExcelToJson()