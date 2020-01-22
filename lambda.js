const excel = require("exceljs");

const addDays = (dateObj, numDays) =>
  new Date(dateObj.setTime(dateObj.getTime() + numDays * 86400000));

const formatDate = date =>
  date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getFullYear();

// read from a file
const workbook = new excel.Workbook();

const cols = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U"
];

const nextColumn = letter => {
  const currentOffset = cols.indexOf(letter);
  if (currentOffset == cols.length - 1) {
    return "V";
  }
  return cols[currentOffset + 1];
};

workbook.xlsx
  .readFile("/Users/ryanmahoney/Documents/excel-lambda/test-in.xlsx")
  .then(() => {
    const newSheet = workbook.addWorksheet("Worksheet");

    const oldSheet = workbook.getWorksheet(1);

    let i = 0;
    for (;;) {
      i = i += 1;
      if (!oldSheet.getCell(`A${i}`).value) {
        break;
      }
      cols.forEach((column, index) => {
        const value = oldSheet.getCell(`${column}${i}`).value;
        const newCell = index > 7 ? nextColumn(column) : column;
        newSheet.getCell(`${newCell}${i}`).value = value;
        if (column === "H") {
          if (i === 1) {
            newSheet.getCell(`I1`).value = "Week ends on";
          } else {
            try {
              const date = new Date(
                value.toISOString().substring(0, 10) + " EDT"
              );
              const newDate = addDays(date, 6);
              newSheet.getCell(`I${i}`).value = formatDate(newDate);
            } catch (e) {}
          }
        }
      });
    }

    workbook.removeWorksheet(1);

    workbook.xlsx
      .writeFile("/Users/ryanmahoney/Documents/excel-lambda/test-out.xlsx")
      .then(() => {
        console.log("done");
      });
  });
