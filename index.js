const excel = require("exceljs");
const AWS = require("aws-sdk");
const s3 = new AWS.S3();
const fs = require("fs");
var nodemailer = require("nodemailer");
var ses = new AWS.SES();

const addDays = (dateObj, numDays) =>
  new Date(dateObj.setTime(dateObj.getTime() + numDays * 86400000));

const formatDate = date =>
  date.getMonth() + 1 + "/" + date.getDate() + "/" + date.getFullYear();

const today = new Date();
const dd = String(today.getDate()).padStart(2, "0");
const mm = String(today.getMonth() + 1).padStart(2, "0"); //January is 0!
const yyyy = today.getFullYear();

const todayFormatted = mm + "-" + dd + "-" + yyyy;

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

exports.handler = async (event, context, callback) => {
  const bucket = event.Records[0].s3.bucket.name;
  const key = event.Records[0].s3.object.key;
  const filename = decodeURIComponent(key.replace(/\+/g, " "));
  const message = `File is uploaded in - ${bucket} -> ${filename}`;
  console.log(message);
  const s3Object = await s3.getObject({ Bucket: bucket, Key: key }).promise();
  const localFile = "/tmp/file.xlsx";
  const newLocalFile = "/tmp/new.xlsx";

  try {
    fs.writeFileSync(localFile, s3Object.Body, "utf8");
  } catch (e) {
    console.log("error writing file");
    console.log(e);
  }

  const exists = fs.existsSync(localFile);
  console.log(`File exists: ${exists ? "Yes" : "No"}`);

  const workbook = new excel.Workbook();

  try {
    await workbook.xlsx.readFile(localFile).then(() => {
      const newSheet = workbook.addWorksheet("Worksheet");
      const oldSheet = workbook.getWorksheet(1);

      let i = 0;
      for (;;) {
        i = i += 1;
        // if there is no value in the first column, stop processing
        if (!oldSheet.getCell(`A${i}`).value) {
          break;
        }
        cols.forEach((column, index) => {
          // get current value of cell
          const value = oldSheet.getCell(`${column}${i}`).value;

          // calculate the cell we will write to so a space is added to account for the new column
          const newCell = index > 5 ? nextColumn(column) : column;
          newSheet.getCell(`${newCell}${i}`).value = value;

          // column F is the "week starts on" column with the start date
          if (column === "F") {
            if (i === 1) {
              // add a new header if this is the first iteration
              newSheet.getCell(`G1`).value = "Week ends on";
            } else {
              // add a new calculated date value
              try {
                const date = new Date(
                  value.toISOString().substring(0, 10) + " EDT"
                );
                const newDate = addDays(date, 6);
                newSheet.getCell(`G${i}`).value = formatDate(newDate);
              } catch (e) {
                console.log("ERROR PROCESSING DATE");
                console.log(value);
              }
            }
          }
        });
      }

      workbook.removeWorksheet(1);

      workbook.xlsx.writeFile(newLocalFile).then(() => {
        const mailOptions = {
          from: "alert@mbta.com",
          subject: "Daily MODIS MBTA Time Sheet Report",
          html: `<p>This reports lists the weekly time sheets for the prior 6 weeks for all MODIS resources working at MBTA CTD. It is scheduled to run daily. If there are any questions on the report, please contact Avery Stroman at astroman@mbta.com.</p>
          <p>The State column reflects the status of the time sheet at the time the report is produced.</p>
          <p>Processed - the time sheet was approved</p>
          <p>Submitted - the time sheet was submitted and is pending approval</p>
          <p>Pending - the time sheet has not been submitted for approval</p>`,
          to: "rmahoney@mbta.com",
          cc: "astroman@mbta.com",
          attachments: [
            {
              // utf-8 string as an attachment
              filename: `daily-timesheet-report-${todayFormatted}.xlsx`,
              path: newLocalFile
            }
          ]
        };

        // create Nodemailer SES transporter
        const transporter = nodemailer.createTransport({
          SES: ses
        });

        // send email
        transporter.sendMail(mailOptions, function(err, info) {
          if (err) {
            console.log("Error sending email");
            callback(err);
          } else {
            console.log("Email sent successfully");
            callback(null);
          }
        });
      });
    });
  } catch (e) {
    console.log(`Excel error: ${e}`);
  }

  callback(null);
};
