//Dependencies
const express = require('express');
const app = express();
const excelJS = require('exceljs');
const XLSX = require('xlsx');
const readLine = require('readline');
const prompt = require("prompt-sync")({ sigint: true });

//Main directory
app.get('/', (req, res) => {
    res.send("This is the main directory.");
});

//Creates variable to store workbook and ensure file exists
const workbook = XLSX.readFile("JobApplications.xlsx");

//Convert the XLSX to JSON
let worksheets = {};
for (const sheetName of workbook.SheetNames) {
    worksheets[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
}

function appendToExcel() {
    //Modify the XLSX after getting user input for values
    let company_name = prompt("Please enter the company name: ");
    let job_title = prompt("Please enter the job title: ");
    let city = prompt("Please enter the city: ");
    let country = prompt("Please enter the country: ");
    let ft_pt = prompt("Please enter FT (full-time) or PT (part-time): ");
    let date = prompt("Please enter the date: ");
    let job_postingID = prompt("Please enter the job posting ID: ");


    worksheets.Sheet1.push({
        "Company Name": company_name,
        "Job Title": job_title,
        "City": city,
        "Country": country,
        "Full Time / Part Time": ft_pt,
        "Date": date,
        "Job Posting ID": job_postingID
    });
}

function updateExcel() {
    //Update the XLSX file
    XLSX.utils.sheet_add_json(workbook.Sheets["Sheet1"], worksheets.Sheet1)
    XLSX.writeFile(workbook, "JobApplications.xlsx");
}

function printExcelJSON() {
    //Show the data as JSON
    console.log("JSON data from excel file:", JSON.stringify(worksheets.Sheet1), "\n");
}

appendToExcel();
updateExcel();
printExcelJSON();

app.listen(3000, ()=> {
    console.log("Server is running on port 3000, http://localhost:3000/.");
});

/* 
Create a new XLSX file
const newBook = XLSX.utils.book_new();
const newSheet = XLSX.utils.json_to_sheet(worksheets.Sheet1);
XLSX.utils.book_append_sheet(newBook, newSheet, "Sheet1");
XLSX.writeFile(newBook,"new-book.xlsx");
*/