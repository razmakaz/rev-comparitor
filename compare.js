const xlsx = require("xlsx");
const fs = require("fs");

const compare = () => {
  // Get both files from input folder. Could be named anything
  const inputDir = fs.readdirSync("./input");

  let earlierFile = null;
  let laterFile = null;

  let earlierTransactions = [];
  let laterTransactions = [];
  let headers = [];

  let missingTransactions = [];

  // List all files
  earlierFile = xlsx.readFile(`./input/previous.xlsx`);
  laterFile = xlsx.readFile(`./input/today.xlsx`);

  // Read the headers from earlier file
  headers = xlsx.utils.sheet_to_json(
    earlierFile.Sheets[earlierFile.SheetNames[0]],
    {
      header: 1, // Use 1 to get the raw data including headers as the first row
      range: 8, // Get the 9th row (0-based index, so 8)
    }
  )[0]; // Use row 9 as the headers

  for (let i = 0; i < headers.length; i++) {
    headers[i] = headers[i] || "";
  }

  console.log(headers);

  // Get all rows from the first sheet on or after row 10 with headers being on row 9
  earlierTransactions = xlsx.utils.sheet_to_json(
    earlierFile.Sheets[earlierFile.SheetNames[0]],
    {
      header: xlsx.utils.sheet_to_json(
        earlierFile.Sheets[earlierFile.SheetNames[0]],
        {
          header: 1, // Use 1 to get the raw data including headers as the first row
          range: 8, // Get the 9th row (0-based index, so 8)
        }
      )[0], // Use row 9 as the headers
      range: 9, // Data from row 10 onwards
    }
  );

  // Get all rows from the first sheet on or after row 10
  laterTransactions = xlsx.utils.sheet_to_json(
    laterFile.Sheets[laterFile.SheetNames[0]],
    {
      header: xlsx.utils.sheet_to_json(
        laterFile.Sheets[laterFile.SheetNames[0]],
        {
          header: 1, // Use 1 to get the raw data including headers as the first row
          range: 8, // Get the 9th row (0-based index, so 8)
        }
      )[0], // Use row 9 as the headers
      range: 9, // Data from row 10 onwards
    }
  );

  // Delete everything in the output folder
  const outputDir = fs.readdirSync("./output");
  outputDir.forEach((file) => {
    fs.unlinkSync(`./output/${file}`);
  });

  // Compare any missing rows with "Cheque#" and "Invoice No" from the earlier file to the later file
  earlierTransactions.forEach((earlierTransaction) => {
    const found = laterTransactions.find(
      (laterTransaction) =>
        laterTransaction["Invoice No"] === earlierTransaction["Invoice No"] &&
        laterTransaction["Cheque#"] === earlierTransaction["Cheque#"]
    );

    if (!found) {
      missingTransactions.push(earlierTransaction);
    }
  });

  // Write the missing transactions to a new file using the headers from the earlier file
  const newWB = xlsx.utils.book_new();
  const newWS = xlsx.utils.json_to_sheet(missingTransactions, {
    // Append 8 rows before the data
    origin: 8,
    header: headers,
  });
  xlsx.utils.book_append_sheet(newWB, newWS, "Missing Transactions");
  xlsx.writeFile(newWB, "./output/missing.xlsx");

  console.log("Missing Transactions: ", missingTransactions);
};

compare();
