import express from "express";
import { dirname, join } from "path";
import { fileURLToPath } from "url";
import fs from "fs";
import XLSX from "xlsx";

const __dirname = dirname(fileURLToPath(import.meta.url));

const app = express();
const port = 5000;

app.set('view engine', 'ejs');
app.set('views', join(__dirname, 'views'));

app.use(express.urlencoded({ extended: true }));

app.use(express.static(join(__dirname, 'public')));

app.get("/", (req, res) => {
    res.render("index");
});

app.post("/submit", (req, res) => {
    const formData = {
        name: req.body.NAME,
        phone: req.body.PHONE,
        email: req.body.EMAIL,
        roll: req.body.ROLL,
        password: req.body.PASSWORD
    };

    const filePath = join(__dirname, "registration_data.xlsx");

    let workbook;
    if (fs.existsSync(filePath)) {
        workbook = XLSX.readFile(filePath);
    } else {
        workbook = XLSX.utils.book_new();
        workbook.SheetNames.push("Registrations");
        const worksheetData = [
            ["Name", "Phone", "Email", "Roll No.", "Password"]
        ];
        const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);
        workbook.Sheets["Registrations"] = worksheet;
    }

    const worksheet = workbook.Sheets["Registrations"];
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const newRow = [formData.name, formData.phone, formData.email, formData.roll, formData.password];
    sheetData.push(newRow);

    const updatedWorksheet = XLSX.utils.aoa_to_sheet(sheetData);
    workbook.Sheets["Registrations"] = updatedWorksheet;

    XLSX.writeFile(workbook, filePath);

    res.send("Congratulations! Your registration data has been saved.");
});

app.listen(port, () => {
    console.log(`Listening on port ${port}`);
});
