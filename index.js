import fs from 'fs-extra';
import {google} from "googleapis";

// To edit available scopes navigate to https://admin.google.com/u/1/ac/owl/domainwidedelegation?hl=en_US
const auth = new google.auth.GoogleAuth({
  keyFile: "project-id-0974793541229075327-841878290eca.json",
  scopes: ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
});

const wcag = fs.readJSONSync('./wcag.json');
const client = "BrandExtract";
const version = "2.2";
const level = "AAA";
const pages = [siteLink("https://www.brandextract.com", "Home"), siteLink("https://www.brandextract.com/About/", "About"), siteLink("https://www.brandextract.com/Insights/", "Insights"), siteLink("https://www.brandextract.com/Contact/", "Contact"), siteLink("https://www.brandextract.com/About/Jobs/", "Jobs")];


const sheets = google.sheets({version: 'v4', auth});
const drive = google.drive({version: 'v3', auth});

const resource = {
    properties: {
        title: `${client} - WCAG ${version} ${level} Audit`,
    },
};

let styles = {
    header: {
        backgroundColor: [28, 69, 135],
        textColor: [255, 255, 255],
        fontSize: 11,
        bold: true,
        verticalAlignment: "MIDDLE"
    },
    subheader: {
        backgroundColor: [17, 85, 204],
        textColor: [255, 255, 255],
        fontSize: 11,
        bold: true,
    },
    guideline: {
        backgroundColor: [60, 120, 216],
        textColor: [255, 255, 255],
        fontSize: 11,
        verticalAlignment: "TOP"
    },
    criterionOdd: {
        backgroundColor: [201, 218, 248],
        verticalAlignment: "TOP"
    },
    criterionEven: {
        verticalAlignment: "TOP"
    }
}

let columnWidths = [194, 446, 72].concat(new Array(pages.length).fill(289));

// Create the spreadsheet
try {
    const spreadsheet = await sheets.spreadsheets.create({
        resource,
        fields: ['spreadsheetId', 'sheets']
    });
    
    // Retrieve the existing parents to remove
    const file = await drive.files.get({
      fileId: spreadsheet.data.spreadsheetId,
      fields: 'parents',
    });

    // Move the file to the new folder
    const previousParents = file.data.parents.join(',');
    
    const files = await drive.files.update({
        includeItemsFromAllDrives: true,
        supportsAllDrives: true,
        fileId: spreadsheet.data.spreadsheetId,
        addParents: "1EevvAmFabUUw-lrqFSAtWc5CKBvkE9m9",
        removeParents: previousParents,
        fields: 'id, parents',
    });

    // Loop through the principles
    await asyncForEach(wcag.principles, async (principle, index) => {
        
        if (principle.versions.includes(version)) {
            
            principle.sheetTitle = `${principle.num}. ${principle.handle}`;
            
            // Create a sheet for the principle
            let thisSheet = await sheets.spreadsheets.batchUpdate({
                spreadsheetId: spreadsheet.data.spreadsheetId,
                resource: { requests: { addSheet: { properties: { title: principle.sheetTitle } } } }
            });

            let sheetId = thisSheet.data.replies[0].addSheet.properties.sheetId;

            let rowStyles = ["header", "subheader"];
            
            let sheetData = [
                [ `${principle.handle} - ${principle.title}` ],
                [ "Success Criteria", "Description", "WCAG Level"].concat(pages)
            ];
            
            // Loop through the principles guidelines
            await asyncForEach(principle.guidelines, async (guideline) => {
                if (guideline.versions.includes(version)) {
                    
                    sheetData.push([ wcagLink(guideline), `${guideline.title}` ]);
                    rowStyles.push("guideline");
                    
                    // Loop throught this guideline's success criteria
                    await asyncForEach(guideline.successcriteria, async(successcriterion, index) => {
                        if (successcriterion.versions.includes(version) && level.includes(successcriterion.level)) {
                            
                            let body = successcriterion.title;
                            
                            if (successcriterion.details) {
                                successcriterion.details.forEach(detail => {
                                    detail.items.forEach(item => {
                                        body += `\n\n${item.handle} - ${item.text}`;
                                    });
                                });
                            }
                            
                            rowStyles.push(`criterion${index % 2 ? "Even": "Odd"}`);
                            sheetData.push([ wcagLink(successcriterion), body, `${successcriterion.versions[0]}\n${successcriterion.level}` ]);
                            
                        }
                    });
                }
            });

            // Style the rows
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId: spreadsheet.data.spreadsheetId,
                resource: {
                    "requests": rowStyles.map((style, index) => {
                        return formatRow({ sheetId, ...styles[style], row: index+1 });
                    })
                }
            });
            
            // Insert the sheet data
            await sheets.spreadsheets.values.batchUpdate({
                spreadsheetId: spreadsheet.data.spreadsheetId,
                resource: { 
                    data: [{ range: `'${principle.sheetTitle}'!A1`, values: sheetData}], 
                    valueInputOption: "USER_ENTERED" 
                }
            });

            // Add the dropdowns
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId: spreadsheet.data.spreadsheetId,
                resource: {
                    "requests": rowStyles.map((style, index) => {
                        if (style !== "criterionOdd" && style !== "criterionEven") return null;
                        return populateDropDowns({ sheetId, row: index+1 });
                    }).filter(row => { return row !== null })
                }
            });

            // Add failed color coding
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId: spreadsheet.data.spreadsheetId,
                resource: {
                    "requests": rowStyles.map((style, index) => {
                        if (style !== "criterionOdd" && style !== "criterionEven") return null;
                        return colorCodeDropDowns({ sheetId, row: index+1, value: "FAILED", textColor: [177, 2, 2], backgroundColor: [254, 201, 195] });
                    }).filter(row => { return row !== null })
                }
            });

            // Add passed color coding
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId: spreadsheet.data.spreadsheetId,
                resource: {
                    "requests": rowStyles.map((style, index) => {
                        if (style !== "criterionOdd" && style !== "criterionEven") return null;
                        return colorCodeDropDowns({ sheetId, row: index+1, value: "PASSED", textColor: [17, 115, 75  ], backgroundColor: [206, 234, 183] });
                    }).filter(row => { return row !== null })
                }
            });

            // Add cannot tell color coding
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId: spreadsheet.data.spreadsheetId,
                resource: {
                    "requests": rowStyles.map((style, index) => {
                        if (style !== "criterionOdd" && style !== "criterionEven") return null;
                        return colorCodeDropDowns({ sheetId, row: index+1, value: "CANNOT TELL", textColor: [71, 56, 33], backgroundColor: [255, 255, 155] });
                    }).filter(row => { return row !== null })
                }
            });

            // Add not present color coding
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId: spreadsheet.data.spreadsheetId,
                resource: {
                    "requests": rowStyles.map((style, index) => {
                        if (style !== "criterionOdd" && style !== "criterionEven") return null;
                        return colorCodeDropDowns({ sheetId, row: index+1, value: "NOT PRESENT", textColor: [17, 115, 75  ], backgroundColor: [206, 234, 183] });
                    }).filter(row => { return row !== null })
                }
            });


            // Set the column widths
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId: spreadsheet.data.spreadsheetId,
                resource: {
                    "requests": columnWidths.map((width, index) => {
                        return setColumnWidths({ sheetId, column: index+1, width });
                    })
                }
            });

            await freezeHeaderRowsandColumns(spreadsheet, sheetId);
            await mergeCells(spreadsheet, sheetId);        
            
        };

    });

    await removeDefaultSheet(spreadsheet);


    
    
} catch (err) {
    // TODO(developer) - Handle error
    throw err;
}


async function removeDefaultSheet(spreadsheet) {
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheet.data.spreadsheetId,
        resource: {
            "requests": [{
                "deleteSheet": {
                    "sheetId": 0
                }
            }]
        }
    });
};

async function freezeHeaderRowsandColumns(spreadsheet, sheetId) {
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheet.data.spreadsheetId,
        resource: {
            "requests": [{
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheetId,
                        "gridProperties": {
                            "frozenRowCount": 2,
                            "frozenColumnCount": 2
                        }
                    },
                    "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
                }
            }]
        }
    });
}

async function mergeCells(spreadsheet, sheetId) {
    await sheets.spreadsheets.batchUpdate({
        spreadsheetId: spreadsheet.data.spreadsheetId,
        resource: {
            "requests": [
                {
                  "mergeCells": {
                    "range": {
                      "sheetId": sheetId,
                      "startRowIndex": 0,
                      "endRowIndex": 1,
                      "startColumnIndex": 0,
                      "endColumnIndex": 2
                    },
                    "mergeType": "MERGE_ALL"
                  }
                }
            ]
        }
    });
}

function formatRow({ sheetId, row, backgroundColor = [255, 255, 255], textColor = [0, 0, 0], fontSize = 10, bold = false, verticalAlignment = "BOTTOM", wrapStrategy = "WRAP" }) {
    return {
        "repeatCell": {
            "range": {
                "sheetId": sheetId,
                "startRowIndex": row-1,
                "endRowIndex": row
            },
            "cell": {
                "userEnteredFormat": {
                    "backgroundColor": {
                        "red": 1.0/255.0 * [backgroundColor[0]],
                        "green": 1.0/255.0 * [backgroundColor[1]],
                        "blue": 1.0/255.0 * [backgroundColor[2]]
                    },
                    "verticalAlignment": verticalAlignment,
                    "textFormat": {
                        "foregroundColor": {
                            "red": 1.0/255 * [textColor[0]],
                            "green": 1.0/255 * [textColor[1]],
                            "blue": 1.0/255 * [textColor[2]],
                        },
                        "fontSize": fontSize,
                        "bold": bold
                    },
                    "wrapStrategy": wrapStrategy
                },
            },
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,wrapStrategy)"
        }
    };
};

function colorCodeDropDowns({ sheetId, row, value, textColor, backgroundColor }) {
    return {
        "addConditionalFormatRule": {
            "rule": {
              "ranges": [
                {
                  "sheetId": sheetId,
                  "startRowIndex": row - 1,
                  "endRowIndex": row,
                  "startColumnIndex": 3,
                  "endColumnIndex": 3 + pages.length,
                }
              ],
              "booleanRule": {
                "condition": {
                  "type": "TEXT_CONTAINS",
                  "values": [
                    {
                      "userEnteredValue": value
                    }
                  ]
                },
                "format": {
                  "textFormat": {
                      "foregroundColor": {
                          "red": 1.0/255.0 * textColor[0],
                          "green": 1.0/255.0 * textColor[1],
                          "blue": 1.0/255.0 * textColor[2]
                      },
                  },
                  "backgroundColor": {
                      "red": 1.0/255.0 * backgroundColor[0],
                      "green": 1.0/255.0 * backgroundColor[1],
                      "blue": 1.0/255.0 * backgroundColor[2]
                  },
                }
              }
            },
            "index": 0
        }
    }
}
    
function populateDropDowns({ sheetId, row }) {
    return {
        "setDataValidation": {
            "range": {
                "sheetId": sheetId,
                "startRowIndex": row - 1,
                "endRowIndex": row,
                "startColumnIndex": 3,
                "endColumnIndex": 3 + pages.length
            },
            "rule": {
                "condition": {
                    "type": 'ONE_OF_LIST',
                    "values": [
                        {
                            "userEnteredValue": 'PASSED',
                        },
                        {
                            "userEnteredValue": 'FAILED',
                        },
                        {
                            "userEnteredValue": 'CANNOT TELL',
                        },
                        {
                            "userEnteredValue": 'NOT PRESENT',
                        },
                        {
                            "userEnteredValue": 'NOT CHECKED',
                        },
                    ],
                },
                "showCustomUi": true,
                "strict": true
            }
        }
    }
}

function setColumnWidths( { sheetId, column, endColumn = column, width } ) {
    return {
        "updateDimensionProperties": {
            "range": {
                "sheetId": sheetId,
                "dimension": "COLUMNS",
                "startIndex": column - 1,
                "endIndex": endColumn
            },
            "properties": {
                "pixelSize": width
            },
            "fields": "pixelSize"
        }
    };
}

// Async for each function
async function asyncForEach(array, callback) {
    for (let index = 0; index < array.length; index++) {
        await callback(array[index], index, array)
    }
}

function wcagLink(target) {
    return `=HYPERLINK("https://www.w3.org/TR/WCAG22/#${formatAnchor(target.handle)}", "${target.num} ${target.handle}")`;
}

function siteLink(target, name) {
    return `=HYPERLINK("${target}", "${name}")`;
}

function formatAnchor(url) {
    return url.replaceAll("(", "", ).replaceAll(")", "", ).replaceAll(" ", "-").toLowerCase();
}