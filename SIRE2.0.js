const excelToJson = require("convert-excel-to-json");
const fs = require("fs");
const path = require("path");
const { v4: uuidv4 } = require("uuid"); 

// Correct directory or use the specific path of the uploaded file
const excelDir = "exceles/questionExceles/";
const SheetName = "Basic attributes"; // Ensure this matches the actual sheet name

// Get all Excel files in the directory
const excelFiles = fs
    .readdirSync(excelDir)
    .filter((file) => file.endsWith(".xlsx"));

// String to store all SQL queries
let allQuestionSqls = "";
let allSectionSqls = "";
let allChapterSqls = "";

// Helper function to convert boolean-like values
function toBoolean(value) {
    if (value === undefined) {
        return "NULL";
    }
    return value && value.toString().toLowerCase() === "true" ? "TRUE" : "FALSE";
}

// Process each Excel file
excelFiles.forEach((file) => {
    const filePath = path.join(excelDir, file);

    // Convert Excel file to JSON
    const allData = excelToJson({
        sourceFile: filePath,
        header: {
            rows: 3, // Skip the first three rows
        },
        sheets: [
            {
                name: SheetName,
                columnToKey: {
                    A: "ccode",
                    B: "chaptercode",
                    C: "sectioncode",
                    D: "questioncode",
                    E: "description",
                    F: "sectiontitle",
                    G: "questiontitle",
                    H: "questiontype",
                    I: "chemical",
                    J: "lng",
                    K: "lpg",
                    L: "oil",
                    M: "conditional",
                    N: "hardwareresponsetype",
                    O: "humanresponsetype",
                    P: "processresponsetype",
                    Q: "roviqlist",
                    R: "captain",
                    S: "chiefofficer",
                    T: "secondofficer",
                    U: "thirdofficer",
                    V: "chiefengineer",
                    W: "secondengineer",
                    X: "thirdengineer",
                    Y: "fourthengineer",
                    Z: "electrician",
                    AA: "chiefcook",
                    AB: "deckrating",
                    AC: "enginerating",
                    AD: "bridge",
                    AE: "accommodation",
                    AF: "ccr",
                    AG: "ecr",
                    AH: "pumproom",
                    AI: "maindeck",
                    AJ: "mooring",
                    AK: "documentation",
                    AL: "lsa",
                },
            },
        ],
    });

    // Access the specific sheet data
    let allexcellData = allData[SheetName];
    let QuestionsqlStatements = "";
    let sectionsqlStatements = "";
    let chaptersqlStatements = "";

    //   Auto generated Values
    // chapterId =  uuidv4();
    // sectionId =  uuidv4();
    // questionId =  uuidv4();


    //   Deffault values
    title="TO DO"
    version = "2.0";
    createdby = "11111111-1111-1111-1111-111111111111";
    modifiedby = "11111111-1111-1111-1111-111111111111";

    // Generate Question SQL for each row of data
    allexcellData.forEach((item) => {
        // Check if 'description' exists
        if (item.description) {
            const questionId = uuidv4();
            QuestionsqlStatements += `INSERT INTO questions(questionId, ccode, reatedby, version, questioncode, description, questionorder, title, questiontype, chemical, lng, lpg, oil, conditional, hardwareresponsetype, humanresponsetype, processresponsetype, roviqlist, captain, chiefofficer, secondofficer, thirdofficer, chiefengineer, secondengineer, thirdengineer, fourthengineer, electrician, chiefcook, deckrating, enginerating, bridge, accommodation, ccr, ecr, pumproom, maindeck, mooring, documentation, lsa
      ) VALUES ('${questionId}', '${item.ccode}', ${createdby}, ${version}, ${parseInt(item.questioncode)}, '${item.description.replace(/’/g, "''")}', ${parseInt(item.questioncode)}, '${item.questiontitle.replace(/’/g, "''")}', '${item.questiontype.replace(/’/g, "''")}', ${toBoolean(item.chemical)}, ${toBoolean(item.lng)}, ${toBoolean(item.lpg)}, ${toBoolean(item.oil)}, ${toBoolean(item.conditional)}, '${item.hardwareresponsetype.replace(/'/g, "''")}', '${item.humanresponsetype.replace(/'/g, "''")}', '${item.processresponsetype.replace(/'/g, "''")}', '${item.roviqlist.replace(/'/g, "''")}', ${toBoolean(item.captain)}, ${toBoolean(item.chiefofficer)}, ${toBoolean(item.secondofficer)}, ${toBoolean(item.thirdofficer)}, ${toBoolean(item.chiefengineer)}, ${toBoolean(item.secondengineer)}, ${toBoolean(item.thirdengineer)}, ${toBoolean(item.fourthengineer)}, ${toBoolean(item.electrician)}, ${toBoolean(item.chiefcook)}, ${toBoolean(item.deckrating)}, ${toBoolean(item.enginerating)}, ${toBoolean(item.bridge)}, ${toBoolean(item.accommodation)}, ${toBoolean(item.ccr)}, ${toBoolean(item.ecr)}, ${toBoolean(item.pumproom)}, ${toBoolean(item.maindeck)}, ${toBoolean(item.mooring)}, ${toBoolean(item.documentation)}, ${toBoolean(item.lsa)}
      );\r\n`;
        }
    });

    // Generate Section SQL for each row of data
    allexcellData.forEach((item) => {
        // Check if 'sectioncode' exists
        if (item.sectioncode) {
            const sectionId = uuidv4();
            sectionsqlStatements += `INSERT INTO sections(
       sectionId, title, sectionorder, modifiedby, createdby, sectioncode, version
      ) VALUES ('${sectionId}', '${item.sectiontitle.replace(/’/g, "''")}', ${parseInt(item.sectioncode)}, '${modifiedby}', '${createdby}', ${parseInt(item.sectioncode)}, ${version}
      );\r\n`;
        }
    });

    // Generate chapter SQL for each row of data
    allexcellData.forEach((item) => {
        // Check if 'chaptercode' exists
        if (item.chaptercode) {
            const chapterId = uuidv4();
          chaptersqlStatements += `INSERT INTO chapters(
           chapterId, title, modifiedby, createdby, chaptercode, version
          ) VALUES (
            '${chapterId}', '${title.replace(/’/g, "''")}','${modifiedby}', '${createdby}', ${parseInt(item.chaptercode)}, ${version}
          );\r\n`;
        }
      });



    // Append the generated SQL to the main string
    allQuestionSqls += QuestionsqlStatements;
    allSectionSqls += sectionsqlStatements;
    allChapterSqls += chaptersqlStatements;
});

// Output the Question SQL to a file
fs.writeFileSync("sqls/questionTable.sql", allQuestionSqls, (err) => {
    if (err) {
        console.error("Error writing to SQL file:", err);
        throw err;
    }
    console.log("All Question SQLs have been saved successfully!");
});

// Output the Section SQL to a file
fs.writeFileSync("sqls/sectionsTable.sql", allSectionSqls, (err) => {
    if (err) {
      console.error("Error writing to SQL file:", err);
      throw err;
    }
    console.log("All Section SQLs have been saved successfully!");
  });

  // Output the Chapter SQL to a file
  fs.writeFileSync("sqls/chaptersTable.sql", allChapterSqls, (err) => {
    if (err) {
      console.error("Error writing to SQL file:", err);
      throw err;
    }
    console.log("All Chapters SQLs have been saved successfully!");
  });