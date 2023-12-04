const express = require("express");
const fs = require("fs");
const path = require("path");
const htmlToDocx = require("html-docx-js");
const bodyParser = require("body-parser");
const cors = require("cors");

const app = express();
const port = 3015;

// Enable CORS for all routes
app.use(cors());

app.use(bodyParser.json({ limit: "500mb" }));
app.use(bodyParser.urlencoded({ extended: true, limit: "500mb" }));

app.post("/generate-docx", (req, res) => {
  // Retrieve HTML string from the request body
  let { htmlString } = req.body;

  htmlString = `
            <html>
            <head>
                <style>
                    /* Set wider page width */
                    body {
                        width: 1000px; /* Set your desired width */
                    }
                </style>
            </head>
            <body>
                ${htmlString}
            </body>
            </html>
        `;

  console.log("htmlString=>", htmlString);

  // Convert HTML to DOCX
  const docx = htmlToDocx.asBlob(htmlString, { orientation: "landscape" });

  // Temporary path to save the file
  const filePath = path.join(__dirname, "temp.docx");

  // Write file to the system
  fs.writeFile(filePath, docx, (err) => {
    if (err) {
      console.error(err);
      return res.status(500).send("Error generating document");
    }

    // Set headers
    res.setHeader("Content-Disposition", "attachment; filename=download.docx");
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );

    // Send the file
    fs.createReadStream(filePath).pipe(res);
  });
});

app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}/`);
});
