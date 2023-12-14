var xlsx = require("xlsx");
var jsonObject = require("./input.json");
var workBook = xlsx.utils.book_new();

var sheetCount = 0;
function appendSheet(jsonObj, sheet, name) {
    sheetCount += 1;
    var data = [];
    data.push(jsonObj);
    xlsx.utils.sheet_add_json(sheet, data, {
        origin: -1,
    });
    xlsx.utils.book_append_sheet(workBook, sheet, `${name}${sheetCount}`);
}


function arrayValue(arr, name) {

    const sheet = xlsx.utils.aoa_to_sheet([[]]);

    var flag = 0;
    if (typeof arr[0] === "object") {
        flag = 1;
    }

    for (let i = 0; i < arr.length; i++) {
        const element = arr[i];
        if (typeof element === "object") {
            for (const key in element) {
                if (element[key] !== null) {
                    if (Array.isArray(element[key])) {
                        element[key] = `sheet::${arrayValue(element[key], key)}`;
                    } else if (typeof element[key] === "object") {
                        element[key] = `sheet::${objectValue(element[key], key)}`;
                    }
                }
            }
        }
    }

    if (flag) {
        sheetCount += 1;
        xlsx.utils.sheet_add_json(sheet, arr, {
            origin: -1,
        });
        xlsx.utils.book_append_sheet(workBook, sheet, `${name}${sheetCount}`);
    }
    else {
        appendSheet(arr, sheet, name);
    }

    return `${name}${sheetCount}`;
}




function objectValue(jsonObj, name) {
    const sheet = xlsx.utils.aoa_to_sheet([[]]);

    for (const key in jsonObj) {
        if (jsonObj[key] !== null) {
            if (Array.isArray(jsonObj[key])) {
                jsonObj[key] = `sheet::${arrayValue(jsonObj[key], key)}`;
            } else if (typeof jsonObj[key] === "object") {
                jsonObj[key] = `sheet::${objectValue(jsonObj[key], key)}`;
            }
        }
    }

    appendSheet(jsonObj, sheet, name);
    return `${name}${sheetCount}`;
}


objectValue(jsonObject, "start");



xlsx.writeFile(workBook, "output.xlsx");
