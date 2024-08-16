const excelToJson = require('convert-excel-to-json');
const fs = require('fs');
const path = require('path');


const excelDir = 'exceles/';

// Define the sheet name working with
const vesselSheetName = 'Vessel Particlars ';

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
                range: 'A2:T2',
                columnToKey: {
                    B: 'VesselName',
                    C: 'CallSign',
                    D: 'Flag',
                    E: 'PortOfRegistry',
                    F: 'OfficialNumber',
                    G: 'ImoNumber',
                    H: 'MMSI',
                    I: 'ClassSociety',
                    J: 'VesselType',
                    K: 'DateKeelLaid',
                    L: 'Delivery',
                    M: 'ClassNotation',
                    N: 'PIClub',
                    O: 'Owners',
                    P: 'Operators',
                    Q: 'LOA',
                    R: 'LBP',
                    S: 'BreadthMoulded',
                    T: 'DepthModulded'
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

        if (item.VesselName) {
            vesselSqls +=
                `INSERT INTO public."VesselParticlars"(
                "VesselName", "CallSign", "Flag", "PortOfRegistry", "OfficialNumber", "ImoNumber", "MMSI", "ClassSociety", "VesselType", "DateKeelLaid", "Delivery", "ClassNotation", "PIClub", "Owners", "Operators", "LOA", "LBP", "BreadthMoulded", "DepthModulded"
                ) OVERRIDING SYSTEM VALUE VALUES (
                  '${item.VesselName}', '${item.CallSign}', '${item.Flag}', '${item.PortOfRegistry}', '${item.OfficialNumber}', '${item.ImoNumber}', '${item.MMSI}', '${item.ClassSociety}', '${item.VesselType}', '${item.DateKeelLaid}', '${item.Delivery}', '${item.ClassNotation}', '${item.PIClub}', '${item.Owners}', '${item.Operators}', '${item.LOA}', '${item.LBP}', '${item.BreadthMoulded}', '${item.DepthModulded}'
                );\r\n`;
        }
    }

    // Append the queries from this file to the overall SQL string
    allVesselSqls += vesselSqls;
});
// Write all SQL queries to a single SQL file
fs.writeFile('sqls/multiVessel.sql', allVesselSqls, function (err) {
    if (err) throw err;
    console.log('All Vessel Particlars Saved!');
});
