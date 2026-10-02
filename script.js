/* =========================================================
   KPI WEBSITE - EXCEL BASED SYSTEM
   ========================================================= */


/* =========================================================
   1. GLOBAL DATA
   ========================================================= */

let kpiData = [];
let headers = [];

let currentEmployee = null;


/* =========================================================
   2. WEBSITE CONFIGURATION
   ========================================================= */

let config = {

    // Which columns can be used for search
    searchColumns: [
        "Emp.ID",
        "Name"
    ],

    // Employee information
    employeeColumns: [
        "Emp.ID",
        "Name",
        "Department",
        "Category",
        "B.Unit"
    ],

    // KPI result columns
    resultColumns: [
        "Total KPI",
        "Weight",
        "Qualitative (Weight)",
        "Quantitative (Weight)",
        "Performed Score",
        "Qualitative (Score)",
        "Quantitative (Score)",
        "Qualitative total",
        "Quantitative Total",
        "Score/100",
        "Overall Rank",
        "Dept.Rank"
    ],

    // Main final score
    finalScoreColumn: "Score/100",

    // Excel sheet name
    sheetName: "Individual Score",

    // Excel header row
    headerRow: 1
};


/* =========================================================
   3. LOAD SAVED CONFIGURATION
   ========================================================= */

function loadConfiguration() {

    const savedConfig =
        localStorage.getItem("kpiWebsiteConfig");

    if (savedConfig) {

        try {

            config = JSON.parse(savedConfig);

        } catch (error) {

            console.log("Configuration load error.");

        }

    }

}


/* =========================================================
   4. SAVE CONFIGURATION
   ========================================================= */

function saveConfiguration() {

    localStorage.setItem(
        "kpiWebsiteConfig",
        JSON.stringify(config)
    );

}


/* =========================================================
   5. EXCEL FILE UPLOAD
   ========================================================= */

function handleExcelUpload() {

    const fileInput = document.getElementById("excelFile");
    const message = document.getElementById("uploadMessage");

    if (!fileInput) {
        alert("Excel file input not found.");
        return;
    }

    const file = fileInput.files[0];

    if (!file) {
        message.innerText = "❌ Please select an Excel file first.";
        return;
    }

    const allowedExtensions = [".xlsx", ".xls", ".csv"];

    const fileName = file.name.toLowerCase();

    const validFile = allowedExtensions.some(ext =>
        fileName.endsWith(ext)
    );

    if (!validFile) {
        message.innerText =
            "❌ Please upload an Excel file (.xlsx, .xls or .csv).";
        return;
    }

    message.innerText = "⏳ Reading Excel file...";

    const reader = new FileReader();

    reader.onload = function (event) {

        try {

            const data = new Uint8Array(
                event.target.result
            );

            const workbook = XLSX.read(data, {
                type: "array"
            });

            console.log(
                "Available Sheets:",
                workbook.SheetNames
            );


            /* ---------------------------------
               FIND SHEET
            --------------------------------- */

            let sheetName = "Individual Score";

            if (!workbook.SheetNames.includes(sheetName)) {

                sheetName = workbook.SheetNames[0];

            }


            const worksheet =
                workbook.Sheets[sheetName];


            if (!worksheet) {

                throw new Error(
                    "No worksheet found."
                );

            }


            /* ---------------------------------
               READ EXCEL
            --------------------------------- */

            const rows =
                XLSX.utils.sheet_to_json(
                    worksheet,
                    {
                        header: 1,
                        defval: "",
                        raw: false
                    }
                );


            if (!rows || rows.length === 0) {

                throw new Error(
                    "Excel file is empty."
                );

            }


            /* ---------------------------------
               HEADER ROW
               Your file header is Row 2
            --------------------------------- */

            const headerRowIndex = 2;

            headers =
                rows[headerRowIndex]
                .map(value =>
                    String(value).trim()
                );


            console.log(
                "Excel Headers:",
                headers
            );


            /* ---------------------------------
               CHECK HEADER
            --------------------------------- */

            if (
                !headers.includes("Emp.ID") ||
                !headers.includes("Name")
            ) {

                throw new Error(
                    "Required columns Emp.ID and Name were not found."
                );

            }


            /* ---------------------------------
               CREATE DATA
            --------------------------------- */

            kpiData = [];


            for (
                let i = headerRowIndex + 1;
                i < rows.length;
                i++
            ) {

                const row = rows[i];

                if (!row) {
                    continue;
                }


                const isEmpty =
                    row.every(
                        value =>
                            value === "" ||
                            value === null ||
                            value === undefined
                    );


                if (isEmpty) {
                    continue;
                }


                const employee = {};


                headers.forEach(
                    (header, index) => {

                        employee[header] =
                            row[index] ?? "";

                    }
                );


                /*
                 * Skip rows without Employee ID
                 */

                if (
                    employee["Emp.ID"] === "" ||
                    employee["Emp.ID"] === undefined
                ) {

                    return;

                }


                kpiData.push(employee);

            }


            /* ---------------------------------
               SAVE DATA
            --------------------------------- */

            localStorage.setItem(
                "kpiData",
                JSON.stringify(kpiData)
            );


            localStorage.setItem(
                "kpiHeaders",
                JSON.stringify(headers)
            );


            /* ---------------------------------
               UPDATE DASHBOARD
            --------------------------------- */

            updateDashboard();


            /* ---------------------------------
               SUCCESS MESSAGE
            --------------------------------- */

            message.innerHTML = `

                <span style="color:green; font-weight:bold;">
                    ✅ Excel uploaded successfully!
                </span>

                <br>

                File:
                ${file.name}

                <br>

                Sheet:
                ${sheetName}

                <br>

                Records:
                ${kpiData.length}

            `;


            console.log(
                "KPI Data:",
                kpiData
            );


            /*
             * Refresh Editor Column Settings
             */

            if (
                typeof createEditorConfiguration ===
                "function"
            ) {

                createEditorConfiguration();

            }


        } catch (error) {

            console.error(
                "Excel Upload Error:",
                error
            );


            message.innerHTML = `

                <span style="color:red; font-weight:bold;">
                    ❌ Excel upload failed!
                </span>

                <br>

                ${error.message}

            `;

        }

    };


    reader.onerror = function () {

        message.innerText =
            "❌ Could not read the Excel file.";

    };


    reader.readAsArrayBuffer(file);

}
