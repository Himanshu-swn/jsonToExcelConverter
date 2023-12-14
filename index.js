var xlsx = require("xlsx");
var jsonObject = require("./input.json");
var workBook = xlsx.utils.book_new();       // creating workbook

var sheetCount = 0;
function appendSheet(jsonObj, sheet, name) {        // function to append sheet into workbook
    sheetCount += 1;                                //  using book_append_sheet() function of 'xlsx' package
    var data = [];                                  // taking 3 arguments:- object to be converted, sheet, name of sheet
    data.push(jsonObj);
    xlsx.utils.sheet_add_json(sheet, data, {
        origin: -1,
    });
    xlsx.utils.book_append_sheet(workBook, sheet, `${name}${sheetCount}`);
}


function arrayValue(arr, name) {                        // recursive function which will handle the case when
                                                        // value of any key is an array

    const sheet = xlsx.utils.aoa_to_sheet([[]]);        // creating empty sheet

    var flag = 0;                                       // 2 cases arises here: if items inside arrays are
    if (typeof arr[0] === "object") {                   // objects or not
        flag = 1;
    }

    for (let i = 0; i < arr.length; i++) {
        const element = arr[i];
        if (typeof element === "object") {                      
            for (const key in element) {
                if (element[key] !== null) {            
                    if (Array.isArray(element[key])) {
                        element[key] = `sheet::${arrayValue(element[key], key)}`;   // using recursion to handle 
                    } else if (typeof element[key] === "object") {                  // inner objects or arrays
                        element[key] = `sheet::${objectValue(element[key], key)}`;  // and changing arrays and objects 
                    }                                                               // names to the corresponding sheets
                }                                                                   // created from recursion
            }
        }
    }

    if (flag) {                                                 // array will be inserted differently in sheet
        sheetCount += 1;                                        // depending on if it contains objects or not
        xlsx.utils.sheet_add_json(sheet, arr, {
            origin: -1,
        });
        xlsx.utils.book_append_sheet(workBook, sheet, `${name}${sheetCount}`);
    }
    else {
        appendSheet(arr, sheet, name);
    }

    return `${name}${sheetCount}`;                            // returing the name of sheet so that we can link sheet
}                                                             // to previous sheet from where it emerged




function objectValue(jsonObj, name) {                   // recrsive function to handle case when an object is the value
    const sheet = xlsx.utils.aoa_to_sheet([[]]);        // of a key in json file

    for (const key in jsonObj) {
        if (jsonObj[key] !== null) {
            if (Array.isArray(jsonObj[key])) {
                jsonObj[key] = `sheet::${arrayValue(jsonObj[key], key)}`;      // similar to above case
            } else if (typeof jsonObj[key] === "object") {                     // using recursive to create sheets and
                jsonObj[key] = `sheet::${objectValue(jsonObj[key], key)}`;     // replacing them with their 
            }                                                                   // corresponding sheet name
        }
    }

    appendSheet(jsonObj, sheet, name);  
    return `${name}${sheetCount}`;
}


objectValue(jsonObject, "start");       // calling the function to start the recursion
                                        // as initally json file will contain an object so calling objectValue function


xlsx.writeFile(workBook, "output.xlsx");            // creating excel file
