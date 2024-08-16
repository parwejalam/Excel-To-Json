const excelToJson = require('convert-excel-to-json');
const fs = require('fs');

const vesselSheetName = 'Vessel Particlars ';

const allData = excelToJson({
    sourceFile: 'exceles/ARICA BRIDGE.xlsx',
    header: {
        rows:1 ,// Skip the first row (likely a header row)
    },
   
    sheets:[
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
                N: 'PIClub',  // Correct the column name to match your data structure
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
console.log(allData);
/******* Prepare data *******/
let vesselParticlars = allData[vesselSheetName];

/***  Write Providers data ***********************/
let vesselSqls = '';
for (let index = 0; index < vesselParticlars.length; index++) {
    const item = vesselParticlars[index];

    if (item.VesselName) {
        vesselSqls +=
            `INSERT INTO public."VesselParticlars"(
            "VesselName", "CallSign", "Flag", "PortOfRegistry", "OfficialNumber", "ImoNumber", "MMSI", "ClassSociety", "VesselType", "DateKeelLaid", "Delivery", "ClassNotation", "PIClub", "Owners", "Operators", "LOA", "LBP", "BreadthMoulded", "DepthModulded"
            ) OVERRIDING SYSTEM VALUE VALUES (
              '${item.VesselName}', '${item.CallSign}', '${item.Flag}', '${item.PortOfRegistry}', '${item.OfficialNumber}', '${item.ImoNumber}', '${item.MMSI}', '${item.ClassSociety}', '${item.VesselType}', '${item.DateKeelLaid}', '${item.Delivery}', '${item.ClassNotation}', '${item.PIClub}', '${item.Owners}', '${item.Operators}', '${item.LOA}', '${item.LBP}', '${item.BreadthMoulded}', '${item.DepthModulded}'
            );\r\n`
    }
}
// Write to vessel.sql
fs.writeFile('sqls/singleVessel.sql', vesselSqls, function (err) {
    if (err) throw err;
    console.log('Vessel Particlars Saved!');
});
