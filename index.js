const xlsxj = require("xlsx-to-json");
const xl = require('excel4node');
const wb = new xl.Workbook();
const ws = wb.addWorksheet("output");
var finalData = [];
var frequencyData = [];

// Convert Excel To Json
const convertInputExcelToJson = () => {
    return new Promise(async (resolve, reject) => {
        xlsxj({
            input: "input.xlsx",
            output: "output.json"
        }, function (err, result) {
            if (err) {
                console.error(err);
                reject(err)
            } else {
                console.log("Running the script ")
                outputData = result;
                let rez = []
                result.forEach(function (item) {
                    rez[item.genus] ? rez[item.genus]++ : rez[item.genus] = 1;
                });

                xlsxj({
                    input: "frq_repr.xlsx",
                    output: "outputFrq.json"
                }, function (err1, result1) {
                    if (err1) {
                        console.log(err1)
                        reject(err1)
                    } else {
                        /*
* Loop through the data to  find the occurance
* Find the first digit of the occurance count
* If occurance count is 93 then takecount will be the first letter i.e 9
* If occurance is single digit like 4 then take count will be the 1
* Take count is the final count for the number of data to be needs to be written into the output file
*/


                        frequencyData = result1;
                        for (var key in rez) {
                            if (rez.hasOwnProperty(key)) {
                                let takeCount = 0;
                                // console.log("=++++")
                                var filtered = frequencyData.filter(function (el) {

                                    return el.genus === key;
                                });

                                let filterData = result.filter(function (ele) {

                                    return ele.genus == key
                                })
                                takeCount = (filtered[0]['selected'])
                                for (let i = 0; i < takeCount; i++) {
                                    finalData.push(filterData[i])
                                }



                            }

                        }

                        var data = finalData;
                        const headingColumnNames = Object.keys(data[0]);
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
                        console.log("Writing data to excel ...")
                        wb.write("output.xlsx");
                        console.log("Check output.xlsx file for output")

                        resolve()

                    }

                })




            }
        })
    })
}

convertInputExcelToJson()