const excelToJson = require("convert-excel-to-json");
const fs = require("fs");
const path = require("path");

// Correct directory or use the specific path of the uploaded file
const excelDir = "exceles/questionExceles/";
const excelSheetName = "Basic attributes"; // Ensure this matches the actual sheet name

// Get all Excel files in the directory
const excelFiles = fs
  .readdirSync(excelDir)
  .filter((file) => file.endsWith(".xlsx"));

// String to store all SQL queries//Change name acording to requirements
let allSqls = "";

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
        name: excelSheetName,
        columnToKey: {
          B: "chaptercode",
        },
      },
    ],
  });

  // Access the specific sheet data
  let excelData = allData[excelSheetName];
  let sqlStatements = "";

  // console.log(questionsData)
  // Default values
  title = "TODO";
  version = "2.0";
  createdby = "11111111-1111-1111-1111-111111111111";
  modifiedby = "11111111-1111-1111-1111-111111111111";
  chapterId = '';
  sectionId = '';
  questionId = '';

  // Generate SQL for each row of data
  excelData.forEach((item) => {
    // Check if 'chaptercode' exists
    if (item.chaptercode) {
      sqlStatements += `INSERT INTO chapters(
        title, modifiedby, createdby, chaptercode, version
      ) VALUES (
         '${title.replace(/’/g, "''")}','${modifiedby}', '${createdby}', ${parseInt(item.chaptercode)}, ${version}
      );\r\n`;
    }
  });

  // Append the generated SQL to the main string
  allSqls += sqlStatements;
});

// Output the SQL to a file
fs.writeFileSync("sqls/chaptersTable.sql", allSqls, (err) => {
  if (err) {
    console.error("Error writing to SQL file:", err);
    throw err;
  }
  console.log("All Chapters SQLs have been saved successfully!");
});
