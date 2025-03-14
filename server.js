const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const fs = require("fs").promises;

const app = express();
app.use(cors());
app.use(express.json());

const filePath = "datastored.xlsx";
const winnersFilePath = "winners.xlsx";

// Function to load existing Excel data
async function loadExcelData() {
    try {
        if (await fs.access(filePath).then(() => true).catch(() => false)) {
            const workbook = xlsx.readFile(filePath);
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            return xlsx.utils.sheet_to_json(sheet);
        }
        return [];
    } catch (error) {
        console.error("Error loading Excel data:", error);
        return [];
    }
}

// Function to save data to Excel
async function saveExcelData(data) {
    const workbook = xlsx.utils.book_new();
    const sheet = xlsx.utils.json_to_sheet(data);
    xlsx.utils.book_append_sheet(workbook, sheet, "Results");
    await xlsx.writeFile(workbook, filePath);
}

// Endpoint to add or update student marks
app.post("/add-marks", async (req, res) => {
    const { rollNumber, section, morningMarks, eveningMarks } = req.body;
    let data = await loadExcelData();

    let existingStudent = data.find(student => student.RollNumber === rollNumber);

    if (existingStudent) {
        if ((morningMarks && existingStudent.EveningMarks) || (eveningMarks && existingStudent.MorningMarks)) {
            return res.status(400).json({ message: "Morning and Evening marks can't be entered at the same time." });
        }
        if (morningMarks !== undefined) existingStudent.MorningMarks = morningMarks;
        if (eveningMarks !== undefined) existingStudent.EveningMarks = eveningMarks;
    } else {
        data.push({
            RollNumber: rollNumber,
            Section: section,
            MorningMarks: morningMarks || "",
            EveningMarks: eveningMarks || ""
        });
    }

    try {
        await saveExcelData(data);
        res.json({ message: "Marks saved successfully!" });
    } catch (error) {
        console.error("Error saving marks:", error);
        res.status(500).json({ message: "Error saving marks." });
    }
});

// Endpoint to download datastored.xlsx
app.get("/download-results", async (req, res) => {
    const code = req.query.code;
    if (code !== "data") {
        return res.status(403).send("Invalid code. Access denied.");
    }
    if (!(await fs.access(filePath).then(() => true).catch(() => false))) {
        return res.status(404).send("No data stored found.");
    }
    res.download(filePath);
});

// Function to generate winners
async function generateWinners() {
    let data = await loadExcelData();
    let winners = [];

    let sections = {};
    data.forEach(student => {
        if (!sections[student.Section]) sections[student.Section] = [];
        sections[student.Section].push(student);
    });

    Object.keys(sections).forEach(section => {
        let students = sections[section];
        let morningWinners = students
            .filter(s => s.MorningMarks)
            .sort((a, b) => b.MorningMarks - a.MorningMarks)
            .slice(0, 2);
        let eveningWinners = students
            .filter(s => s.EveningMarks)
            .sort((a, b) => b.EveningMarks - a.EveningMarks)
            .slice(0, 1);
        winners.push(...morningWinners, ...eveningWinners);
    });

    const workbook = xlsx.utils.book_new();
    const sheet = xlsx.utils.json_to_sheet(winners);
    xlsx.utils.book_append_sheet(workbook, sheet, "Winners");
    await xlsx.writeFile(workbook, winnersFilePath);
}

// Endpoint to generate winners and download winners list
app.get("/download-winners", async (req, res) => {
    const code = req.query.code;
    if (code !== "winner") {
        return res.status(403).send("Invalid code. Access denied.");
    }
    try {
        if (!(await fs.access(filePath).then(() => true).catch(() => false))) {
            return res.status(404).send("No data stored found.");
        }
        await generateWinners();
        if (await fs.access(winnersFilePath).then(() => true).catch(() => false)) {
            res.download(winnersFilePath);
        } else {
            res.status(500).send("Error generating winners list.");
        }
    } catch (error) {
        console.error("Error in /download-winners:", error);
        res.status(500).send("Server error occurred.");
    }
});

// Start the server using Render's PORT environment variable
const PORT = process.env.PORT || 5000; // Use Render's port or default to 5000 locally
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});