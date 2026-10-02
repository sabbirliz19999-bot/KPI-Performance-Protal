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

function uploadExcelFile(file) {

    if (!file) {

        alert("Please select an Excel file.");

        return;

    }


    const reader = new FileReader();


    reader.onload = function(event) {

        try {

            const data =
                new Uint8Array(event.target.result);


            const workbook =
                XLSX.read(data, {
                    type: "array"
                });


            /*
             * Get configured sheet
             */

            let sheetName = config.sheetName;


            if (!workbook.SheetNames.includes(sheetName)) {

                sheetName =
                    workbook.SheetNames[0];

            }


            const worksheet =
                workbook.Sheets[sheetName];


            /*
             * Convert Excel to array
             */

            const rows =
                XLSX.utils.sheet_to_json(
                    worksheet,
                    {
                        header: 1,
                        defval: ""
                    }
                );


            /*
             * Get header row
             */

            headers =
                rows[config.headerRow]
                .map(h => String(h).trim());


            /*
             * Convert remaining rows
             */

            kpiData = [];


            for (
                let i = config.headerRow + 1;
                i < rows.length;
                i++
            ) {

                const row = rows[i];


                if (
                    !row ||
                    row.every(
                        value =>
                            value === "" ||
                            value === null ||
                            value === undefined
                    )
                ) {

                    continue;

                }


                let employee = {};


                headers.forEach(
                    (header, index) => {

                        employee[header] =
                            row[index] ?? "";

                    }
                );


                kpiData.push(employee);

            }


            /*
             * Save data
             */

            localStorage.setItem(
                "kpiData",
                JSON.stringify(kpiData)
            );


            localStorage.setItem(
                "kpiHeaders",
                JSON.stringify(headers)
            );


            alert(
                "Excel uploaded successfully.\n\n" +
                "Records: " +
                kpiData.length
            );


            updateDashboard();


        } catch (error) {

            console.error(error);

            alert(
                "Excel file could not be processed."
            );

        }

    };


    reader.readAsArrayBuffer(file);

}


/* =========================================================
   6. LOAD EXISTING DATA
   ========================================================= */

function loadSavedData() {

    const savedData =
        localStorage.getItem("kpiData");

    const savedHeaders =
        localStorage.getItem("kpiHeaders");


    if (savedData) {

        try {

            kpiData =
                JSON.parse(savedData);

        } catch (error) {

            kpiData = [];

        }

    }


    if (savedHeaders) {

        try {

            headers =
                JSON.parse(savedHeaders);

        } catch (error) {

            headers = [];

        }

    }

}


/* =========================================================
   7. SEARCH EMPLOYEE
   ========================================================= */

function searchEmployee() {

    const inputElement =
        document.getElementById("searchInput");


    if (!inputElement) {

        return;

    }


    const searchText =
        inputElement.value
            .trim()
            .toLowerCase();


    const resultSection =
        document.getElementById("resultSection");


    if (!searchText) {

        showSearchMessage();

        return;

    }


    /*
     * Search only configured columns
     */

    const results =
        kpiData.filter(employee => {

            return config.searchColumns.some(
                column => {

                    if (
                        employee[column] === undefined ||
                        employee[column] === null
                    ) {

                        return false;

                    }


                    return String(
                        employee[column]
                    )
                    .toLowerCase()
                    .includes(searchText);

                }
            );

        });


    if (results.length === 0) {

        resultSection.innerHTML = `

            <div class="empty-result">

                <div class="search-icon">
                    ❌
                </div>

                <h3>
                    Employee Not Found
                </h3>

                <p>
                    No matching employee was found.
                </p>

            </div>

        `;

        return;

    }


    /*
     * If multiple employees found
     */

    if (results.length > 1) {

        showMultipleResults(results);

        return;

    }


    /*
     * Show exact result
     */

    currentEmployee =
        results[0];

    showEmployeeResult(
        currentEmployee
    );

}


/* =========================================================
   8. SHOW MULTIPLE RESULTS
   ========================================================= */

function showMultipleResults(results) {

    const resultSection =
        document.getElementById(
            "resultSection"
        );


    let html = `

        <div class="employee-card">

            <h2>
                Search Results
            </h2>

            <p>
                ${results.length}
                employees found.
            </p>

            <table class="kpi-table">

                <thead>

                    <tr>

    `;


    config.searchColumns.forEach(
        column => {

            html += `
                <th>${column}</th>
            `;

        }
    );


    html += `

                        <th>Action</th>

                    </tr>

                </thead>

                <tbody>

    `;


    results.forEach(
        (employee, index) => {

            html += `<tr>`;


            config.searchColumns.forEach(
                column => {

                    html += `

                        <td>
                            ${employee[column] ?? ""}
                        </td>

                    `;

                }
            );


            html += `

                <td>

                    <button
                        onclick="selectEmployee(${index})"
                        class="download-btn"
                    >
                        View Result
                    </button>

                </td>

            </tr>

            `;

        }
    );


    html += `

                </tbody>

            </table>

        </div>

    `;


    resultSection.innerHTML =
        html;


    /*
     * Temporary result storage
     */

    window.searchResults =
        results;

}


/* =========================================================
   9. SELECT EMPLOYEE
   ========================================================= */

function selectEmployee(index) {

    const employee =
        window.searchResults[index];


    currentEmployee =
        employee;


    showEmployeeResult(
        employee
    );

}


/* =========================================================
   10. SHOW EMPLOYEE RESULT
   ========================================================= */

function showEmployeeResult(employee) {

    const resultSection =
        document.getElementById(
            "resultSection"
        );


    /*
     * Employee information
     */

    let employeeInfo = "";


    config.employeeColumns.forEach(
        column => {

            employeeInfo += `

                <p>

                    <strong>
                        ${column}:
                    </strong>

                    ${employee[column] ?? ""}

                </p>

            `;

        }
    );


    /*
     * KPI result table
     */

    let resultRows = "";


    config.resultColumns.forEach(
        column => {

            resultRows += `

                <tr>

                    <td>
                        ${column}
                    </td>

                    <td>
                        ${employee[column] ?? ""}
                    </td>

                </tr>

            `;

        }
    );


    /*
     * Final score
     */

    const finalScore =
        employee[
            config.finalScoreColumn
        ] ?? "";


    resultSection.innerHTML = `

        <div
            class="employee-card"
            id="employeeResult"
        >

            <div class="employee-header">

                <div class="employee-info">

                    ${employeeInfo}

                </div>


                <div class="score-box">

                    <span>
                        FINAL KPI SCORE
                    </span>

                    <strong>
                        ${formatScore(finalScore)}
                    </strong>

                </div>

            </div>


            <table class="kpi-table">

                <thead>

                    <tr>

                        <th>
                            KPI / Result
                        </th>

                        <th>
                            Value
                        </th>

                    </tr>

                </thead>


                <tbody>

                    ${resultRows}

                </tbody>

            </table>


            <button
                class="download-btn"
                onclick="downloadPDF()"
            >
                📄 Download PDF
            </button>

        </div>

    `;

}


/* =========================================================
   11. FORMAT SCORE
   ========================================================= */

function formatScore(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return "-";

    }


    const number =
        Number(value);


    if (isNaN(number)) {

        return value;

    }


    return number.toFixed(2) + "%";

}


/* =========================================================
   12. DOWNLOAD PDF
   ========================================================= */

function downloadPDF() {

    if (!currentEmployee) {

        alert(
            "Please search an employee first."
        );

        return;

    }


    const {
        jsPDF
    } = window.jspdf;


    const pdf =
        new jsPDF();


    /*
     * Company title
     */

    pdf.setFontSize(18);

    pdf.text(
        "KPI PERFORMANCE REPORT",
        105,
        20,
        {
            align: "center"
        }
    );


    pdf.setFontSize(10);


    let y = 35;


    /*
     * Employee information
     */

    config.employeeColumns.forEach(
        column => {

            pdf.text(
                `${column}: ${
                    currentEmployee[column] ?? ""
                }`,
                20,
                y
            );

            y += 7;

        }
    );


    y += 5;


    /*
     * Final score
     */

    pdf.setFontSize(14);

    pdf.text(
        `Final KPI Score: ${
            formatScore(
                currentEmployee[
                    config.finalScoreColumn
                ]
            )
        }`,
        20,
        y
    );


    y += 12;


    /*
     * Result
     */

    pdf.setFontSize(10);


    config.resultColumns.forEach(
        column => {

            /*
             * Page break
             */

            if (y > 275) {

                pdf.addPage();

                y = 20;

            }


            pdf.text(
                column + ":",
                20,
                y
            );


            pdf.text(
                String(
                    currentEmployee[column] ?? ""
                ),
                85,
                y
            );


            y += 8;

        }
    );


    /*
     * Footer
     */

    pdf.setFontSize(9);

    pdf.text(
        "KPI Performance Management System",
        105,
        290,
        {
            align: "center"
        }
    );


    /*
     * File name
     */

    const employeeId =
        currentEmployee["Emp.ID"] ||
        "Employee";


    pdf.save(
        employeeId +
        "_KPI_Result.pdf"
    );

}


/* =========================================================
   13. EMPTY SEARCH MESSAGE
   ========================================================= */

function showSearchMessage() {

    const resultSection =
        document.getElementById(
            "resultSection"
        );


    resultSection.innerHTML = `

        <div class="empty-result">

            <div class="search-icon">
                🔍
            </div>

            <h3>
                Search an Employee
            </h3>

            <p>
                Enter Employee ID or Name
                to view KPI performance.
            </p>

        </div>

    `;

}


/* =========================================================
   14. EDITOR - COLUMN CONFIGURATION
   ========================================================= */

function createEditorConfiguration() {

    const container =
        document.getElementById(
            "columnConfiguration"
        );


    if (!container) {

        return;

    }


    if (!headers.length) {

        container.innerHTML = `

            <p>
                Please upload an Excel file first.
            </p>

        `;

        return;

    }


    let html = `

        <h3>
            Configure Website Columns
        </h3>

        <p>
            Select which Excel columns
            should appear on the website.
        </p>

    `;


    headers.forEach(
        header => {

            const employeeChecked =
                config.employeeColumns
                    .includes(header);

            const resultChecked =
                config.resultColumns
                    .includes(header);

            const searchChecked =
                config.searchColumns
                    .includes(header);


            html += `

                <div class="column-setting">

                    <strong>
                        ${header}
                    </strong>


                    <label>

                        <input
                            type="checkbox"
                            ${
                                searchChecked
                                ? "checked"
                                : ""
                            }
                            onchange="
                                toggleColumn(
                                    'searchColumns',
                                    '${escapeQuotes(header)}',
                                    this.checked
                                )
                            "
                        >

                        Search

                    </label>


                    <label>

                        <input
                            type="checkbox"
                            ${
                                employeeChecked
                                ? "checked"
                                : ""
                            }
                            onchange="
                                toggleColumn(
                                    'employeeColumns',
                                    '${escapeQuotes(header)}',
                                    this.checked
                                )
                            "
                        >

                        Employee Info

                    </label>


                    <label>

                        <input
                            type="checkbox"
                            ${
                                resultChecked
                                ? "checked"
                                : ""
                            }
                            onchange="
                                toggleColumn(
                                    'resultColumns',
                                    '${escapeQuotes(header)}',
                                    this.checked
                                )
                            "
                        >

                        Result

                    </label>

                </div>

            `;

        }
    );


    html += `

        <button
            onclick="saveConfiguration()"
            class="upload-btn"
        >
            Save Configuration
        </button>

    `;


    container.innerHTML =
        html;

}


/* =========================================================
   15. TOGGLE COLUMN
   ========================================================= */

function toggleColumn(
    type,
    column,
    enabled
) {

    if (!config[type]) {

        config[type] = [];

    }


    if (enabled) {

        if (
            !config[type].includes(column)
        ) {

            config[type].push(column);

        }

    }

    else {

        config[type] =
            config[type].filter(
                item =>
                    item !== column
            );

    }


    saveConfiguration();

}


/* =========================================================
   16. ESCAPE QUOTES
   ========================================================= */

function escapeQuotes(text) {

    return String(text)
        .replace(/'/g, "\\'");

}


/* =========================================================
   17. DASHBOARD
   ========================================================= */

function updateDashboard() {

    const totalEmployees =
        document.getElementById(
            "totalEmployees"
        );


    if (totalEmployees) {

        totalEmployees.innerText =
            kpiData.length;

    }


    const totalKPI =
        document.getElementById(
            "totalKPI"
        );


    if (totalKPI) {

        totalKPI.innerText =
            kpiData.length;

    }

}


/* =========================================================
   18. EDITOR LOGIN
   ========================================================= */

function showLogin() {

    const modal =
        document.getElementById(
            "loginModal"
        );


    if (modal) {

        modal.style.display =
            "flex";

    }

}


/* =========================================================
   19. EDITOR LOGIN
   ========================================================= */

function editorLogin() {

    const username =
        document.getElementById(
            "username"
        ).value.trim();


    const password =
        document.getElementById(
            "password"
        ).value.trim();


    /*
     * DEMO ONLY
     *
     * Production version should use
     * server-side authentication.
     */

    if (
        username === "admin" &&
        password === "123456"
    ) {

        document.getElementById(
            "loginModal"
        ).style.display =
            "none";


        document.getElementById(
            "editorPanel"
        ).style.display =
            "block";


        createEditorConfiguration();

    }

    else {

        const message =
            document.getElementById(
                "loginMessage"
            );


        if (message) {

            message.innerText =
                "Invalid username or password.";

        }

    }

}


/* =========================================================
   20. LOGOUT
   ========================================================= */

function logout() {

    const panel =
        document.getElementById(
            "editorPanel"
        );


    if (panel) {

        panel.style.display =
            "none";

    }

}


/* =========================================================
   21. PAGE LOAD
   ========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    function() {

        loadConfiguration();

        loadSavedData();

        updateDashboard();

    }
);
