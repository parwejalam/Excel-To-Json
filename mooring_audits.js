const excelToJson = require('convert-excel-to-json');
const fs = require('fs');
const path = require('path');


const excelDir = 'exceles/vessels exceles';

// Define the sheet name working with
const vesselSheetName = 'SIRE 2.0 AND RISQ 3';

const excelFiles = fs.readdirSync(excelDir).filter(file => file.endsWith('.xlsx'));

// an empty string to store all SQL queries
let allVesselSqls = '';

excelFiles.forEach(file => {
    const filePath = path.join(excelDir, file);
    
    // Convert Excel file to JSON
    const allData = excelToJson({
        sourceFile: filePath,
        header: {
            rows: 1, // Skip the first row (likely a header row)
        },
        sheets: [
            {
                name: vesselSheetName,
                range: 'A2:G27',
                columnToKey: {
                    A: 'S_NO',
                    B: 'ITEM',
                    C: 'RISQ_3',
                    D: 'MEG4',
                    E: 'SIRE_2_0',
                    F: 'TMSA_KPI',
                    G: 'OBJECTIVE_EVIDENCE',
                }
            }
        ]
    });

    // console.log(allData)
    
    let vesselParticlars = allData[vesselSheetName];

    // Generate SQL queries for this file's data
    let vesselSqls = '';

    for (let index = 0; index < vesselParticlars.length; index++) {
        const item = vesselParticlars[index];

        if (item.S_NO) {
            vesselSqls +=
                `INSERT INTO public."VesselParticlars"(
                "S_NO", "ITEM", "RISQ_3", "MEG4", "SIRE_2_0", "TMSA_KPI", "OBJECTIVE_EVIDENCE"
                ) OVERRIDING SYSTEM VALUE VALUES (
                  '${item.S_NO}', '${item.ITEM}', '${item.RISQ_3}', '${item.MEG4}', '${item.SIRE_2_0}', '${item.TMSA_KPI}', '${item.OBJECTIVE_EVIDENCE}'
                );\r\n`;
        }
    }

    // Append the queries from this file to the overall SQL string
    allVesselSqls += vesselSqls;
});
// Write all SQL queries to a single SQL file
fs.writeFile('sqls/vessel.sql', allVesselSqls, function (err) {
    if (err) throw err;
    console.log('All Vessel Particlars Saved!');
});
