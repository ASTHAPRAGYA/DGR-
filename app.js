/* =========================================================
   SOLAR DGR ANALYTICS
   app.js

   ONLY USES:

   Dashboard
   Annual_KPI
   Daily_KPI
   PA
   Curtailment records

   ========================================================= */


/* =========================================================
   GLOBAL VARIABLES
========================================================= */

let workbook = null;

let sheetData = {};

let charts = {};

let uploadedFile = null;


/* =========================================================
   REQUIRED WORKSHEETS
========================================================= */

const REQUIRED_SHEETS = [
    "Dashboard",
    "Annual_KPI",
    "Daily_KPI",
    "PA",
    "Curtailment records"
];


/* =========================================================
   DOM ELEMENTS
========================================================= */

const fileInput =
    document.getElementById("dgrFile");

const dropZone =
    document.getElementById("dropZone");

const fileInfo =
    document.getElementById("fileInfo");

const fileName =
    document.getElementById("fileName");

const fileSheets =
    document.getElementById("fileSheets");

const removeFile =
    document.getElementById("removeFile");

const workbookStatus =
    document.getElementById("workbookStatus");

const sheetBadges =
    document.getElementById("sheetBadges");

const statusText =
    document.getElementById("statusText");

const sidebarFileName =
    document.getElementById("sidebarFileName");

const emptyState =
    document.getElementById("emptyState");


/* =========================================================
   INITIAL STATE
========================================================= */

hideAnalytics();


// Navigation still remains visible.
setupNavigation();


// File upload.
setupFileUpload();


/* =========================================================
   NAVIGATION
========================================================= */

function setupNavigation() {

    const navItems =
        document.querySelectorAll(".nav-item");


    navItems.forEach(button => {

        button.addEventListener(
            "click",
            () => {

                navItems.forEach(item => {

                    item.classList.remove(
                        "active"
                    );

                });


                button.classList.add(
                    "active"
                );


                const targetId =
                    button.dataset.target;


                const target =
                    document.getElementById(
                        targetId
                    );


                if (target) {

                    target.scrollIntoView({
                        behavior: "smooth",
                        block: "start"
                    });

                }

            }
        );

    });

}


/* =========================================================
   FILE UPLOAD
========================================================= */

function setupFileUpload() {


    fileInput.addEventListener(
        "change",
        event => {

            const file =
                event.target.files[0];

            if (file) {

                processFile(file);

            }

        }
    );


    dropZone.addEventListener(
        "click",
        () => {

            fileInput.click();

        }
    );


    dropZone.addEventListener(
        "dragover",
        event => {

            event.preventDefault();

            dropZone.classList.add(
                "dragging"
            );

        }
    );


    dropZone.addEventListener(
        "dragleave",
        () => {

            dropZone.classList.remove(
                "dragging"
            );

        }
    );


    dropZone.addEventListener(
        "drop",
        event => {

            event.preventDefault();

            dropZone.classList.remove(
                "dragging"
            );


            const file =
                event.dataTransfer.files[0];


            if (!file) {
                return;
            }


            const extension =
                file.name
                    .split(".")
                    .pop()
                    .toLowerCase();


            if (
                extension !== "xlsx" &&
                extension !== "xls" &&
                extension !== "csv"
            ) {

                alert(
                    "Please upload an Excel or CSV file."
                );

                return;

            }


            processFile(file);

        }
    );


    removeFile.addEventListener(
        "click",
        resetDashboard
    );

}


/* =========================================================
   PROCESS FILE
========================================================= */

function processFile(file) {

    uploadedFile = file;


    statusText.textContent =
        "Reading DGR workbook...";


    const reader =
        new FileReader();


    reader.onload =
        function(event) {


            try {

                const data =
                    new Uint8Array(
                        event.target.result
                    );


                workbook =
                    XLSX.read(
                        data,
                        {
                            type: "array",
                            cellDates: true
                        }
                    );


                readRequiredSheets();

                updateFileInformation();

                buildDashboard();


            }

            catch (error) {

                console.error(error);


                alert(
                    "Unable to read this DGR. Please check that it is a valid Excel workbook."
                );


                statusText.textContent =
                    "Error reading DGR.";

            }

        };


    reader.readAsArrayBuffer(file);

}


/* =========================================================
   READ ONLY REQUIRED SHEETS
========================================================= */

function readRequiredSheets() {

    sheetData = {};


    REQUIRED_SHEETS.forEach(
        sheetName => {

            const actualSheetName =
                findSheetName(sheetName);


            if (!actualSheetName) {

                sheetData[sheetName] = null;

                return;

            }


            const worksheet =
                workbook.Sheets[
                    actualSheetName
                ];


            sheetData[sheetName] =
                XLSX.utils.sheet_to_json(
                    worksheet,
                    {
                        header: 1,
                        defval: null,
                        raw: true
                    }
                );

        }
    );

}


/* =========================================================
   FIND SHEET NAME
   Handles small naming differences.
========================================================= */

function findSheetName(requiredName) {

    if (
        workbook.SheetNames.includes(
            requiredName
        )
    ) {

        return requiredName;

    }


    const normalizedRequired =
        normalizeText(
            requiredName
        );


    return workbook.SheetNames.find(
        sheetName => {

            return normalizeText(
                sheetName
            ) === normalizedRequired;

        }
    ) || null;

}


/* =========================================================
   NORMALIZE TEXT
========================================================= */

function normalizeText(value) {

    return String(value ?? "")
        .trim()
        .toLowerCase()
        .replace(
            /[\s_-]+/g,
            ""
        );

}


/* =========================================================
   FILE INFORMATION
========================================================= */

function updateFileInformation() {

    fileName.textContent =
        uploadedFile.name;


    fileSheets.textContent =
        workbook.SheetNames.length +
        " worksheets detected";


    sidebarFileName.textContent =
        uploadedFile.name;


    fileInfo.classList.remove(
        "hidden"
    );


    workbookStatus.classList.remove(
        "hidden"
    );


    sheetBadges.innerHTML = "";


    REQUIRED_SHEETS.forEach(
        sheetName => {

            const badge =
                document.createElement(
                    "span"
                );


            badge.className =
                "sheet-badge";


            if (
                sheetData[sheetName]
            ) {

                badge.textContent =
                    sheetName;

            }

            else {

                badge.textContent =
                    sheetName +
                    " — not found";

                badge.classList.add(
                    "missing"
                );

            }


            sheetBadges.appendChild(
                badge
            );

        }
    );

}


/* =========================================================
   BUILD DASHBOARD
========================================================= */

function buildDashboard() {

    emptyState.classList.add(
        "hidden"
    );


    statusText.textContent =
        "DGR loaded successfully. Analytics generated from the selected worksheets.";


    showAnalytics();


    buildPAChart();

    buildPRChart();

    buildOperatingHoursChart();

    buildSystemLossChart();

    buildCurtailmentChart();

    buildEnergyChart();


    updateKPIs();

}


/* =========================================================
   SHOW / HIDE ANALYTICS
========================================================= */

function hideAnalytics() {

    document
        .querySelectorAll(
            ".dashboard-section"
        )
        .forEach(section => {

            section.style.display =
                "none";

        });

}


function showAnalytics() {

    document
        .querySelectorAll(
            ".dashboard-section"
        )
        .forEach(section => {

            section.style.display =
                "block";

        });

}


/* =========================================================
   GENERIC CHART DESTROY
========================================================= */

function destroyChart(name) {

    if (charts[name]) {

        charts[name].destroy();

        charts[name] = null;

    }

}


/* =========================================================
   CHART DEFAULTS
========================================================= */

Chart.defaults.font.family =
    "Inter, Arial, sans-serif";


Chart.defaults.font.size = 9;


Chart.defaults.color =
    "#728286";


/* =========================================================
   GET DAILY KPI ROWS
=========================================================

   Excel columns:

   I  = 9th column
   V  = 22nd column
   AD = 30th column

   JavaScript index:

   I  -> 8
   V  -> 21
   AD -> 29

========================================================= */

function getDailyKPIData() {

    const rows =
        sheetData["Daily_KPI"];


    if (
        !rows ||
        rows.length === 0
    ) {

        return [];

    }


    const result = [];


    for (
        let i = 0;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i];


        if (
            !row ||
            row.length === 0
        ) {

            continue;

        }


        const operatingHours =
            parseNumber(
                row[8]
            );


        const pr =
            parseNumber(
                row[21]
            );


        const systemLoss =
            parseNumber(
                row[29]
            );


        const date =
            findDateInRow(
                row,
                i
            );


        if (
            operatingHours === null &&
            pr === null &&
            systemLoss === null
        ) {

            continue;

        }


        result.push({

            index:
                result.length + 1,

            date,

            operatingHours,

            pr,

            systemLoss

        });

    }


    return result;

}


/* =========================================================
   FIND DATE IN ROW
========================================================= */

function findDateInRow(row, index) {

    for (
        let i = 0;
        i < Math.min(
            row.length,
            8
        );
        i++
    ) {

        const value =
            row[i];


        if (
            value instanceof Date &&
            !isNaN(value)
        ) {

            return value;

        }


        if (
            typeof value === "number" &&
            value > 20000 &&
            value < 60000
        ) {

            const date =
                XLSX.SSF.parse_date_code(
                    value
                );


            if (date) {

                return new Date(
                    date.y,
                    date.m - 1,
                    date.d
                );

            }

        }


        if (
            typeof value === "string" &&
            looksLikeDate(value)
        ) {

            const parsed =
                new Date(value);


            if (
                !isNaN(parsed)
            ) {

                return parsed;

            }

        }

    }


    return index + 1;

}


/* =========================================================
   DATE CHECK
========================================================= */

function looksLikeDate(value) {

    const text =
        String(value).trim();


    return (
        /^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/
            .test(text)
        ||
        /^\d{1,2}[\/\-]\w{3,9}[\/\-]\d{2,4}$/
            .test(text)
    );

}


/* =========================================================
   PARSE NUMBER
========================================================= */

function parseNumber(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return null;

    }


    if (
        typeof value === "number" &&
        Number.isFinite(value)
    ) {

        return value;

    }


    let text =
        String(value)
            .trim()
            .replace(
                /,/g,
                ""
            );


    if (!text) {

        return null;

    }


    const percent =
        text.includes("%");


    text =
        text.replace(
            /[^0-9.\-+]/g,
            ""
        );


    if (!text) {

        return null;

    }


    const number =
        Number(text);


    if (
        !Number.isFinite(number)
    ) {

        return null;

    }


    return number;

}


/* =========================================================
   FORMAT DATE
========================================================= */

function formatDate(value) {

    if (
        value instanceof Date &&
        !isNaN(value)
    ) {

        return value.toLocaleDateString(
            "en-IN",
            {
                day: "2-digit",
                month: "short"
            }
        );

    }


    return String(value);

}


/* =========================================================
   BUILD PR SCATTER CHART
   Daily_KPI Column V
========================================================= */

function buildPRChart() {

    const data =
        getDailyKPIData();


    const valid =
        data.filter(
            item =>
                item.pr !== null
        );


    destroyChart("pr");


    const canvas =
        document.getElementById(
            "prChart"
        );


    if (!canvas) {
        return;
    }


    charts.pr =
        new Chart(
            canvas,
            {

                type: "scatter",

                data: {

                    datasets: [

                        {

                            label:
                                "Performance Ratio",

                            data:
                                valid.map(
                                    item => ({

                                        x:
                                            item.index,

                                        y:
                                            item.pr

                                    })
                                ),

                            pointRadius: 5,

                            pointHoverRadius: 8,

                            backgroundColor:
                                "#27A5AD",

                            borderColor:
                                "#27A5AD"

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,

                    interaction: {

                        mode: "nearest",

                        intersect: false

                    },

                    plugins: {

                        legend: {

                            display: false

                        },

                        tooltip: {

                            callbacks: {

                                title:
                                    tooltipPRTitle,

                                label:
                                    tooltipPRLabel

                            }

                        }

                    },


                    scales: {

                        x: {

                            title: {

                                display: true,

                                text:
                                    "Day of Month"

                            },

                            ticks: {

                                stepSize: 1

                            },

                            grid: {

                                color:
                                    "#edf2f3"

                            }

                        },


                        y: {

                            title: {

                                display: true,

                                text:
                                    "PR (%)"

                            },

                            grid: {

                                color:
                                    "#edf2f3"

                            }

                        }

                    }

                }

            }
        );


    buildDashboardPRChart(
        valid
    );

}


/* =========================================================
   PR TOOLTIP
========================================================= */

function tooltipPRTitle(items) {

    if (
        !items ||
        !items.length
    ) {

        return "";

    }


    const index =
        items[0].parsed.x - 1;


    const data =
        getDailyKPIData();


    const item =
        data[index];


    if (!item) {

        return "Day " +
            items[0].parsed.x;

    }


    return (
        "Day " +
        items[0].parsed.x +
        " — " +
        formatDate(item.date)
    );

}


function tooltipPRLabel(context) {

    return (
        "PR: " +
        Number(
            context.parsed.y
        ).toFixed(2) +
        "%"
    );

}


/* =========================================================
   DASHBOARD PR
========================================================= */

function buildDashboardPRChart(data) {

    destroyChart(
        "dashboardPR"
    );


    const canvas =
        document.getElementById(
            "dashboardPRChart"
        );


    if (!canvas) {
        return;
    }


    charts.dashboardPR =
        new Chart(
            canvas,
            {

                type: "scatter",

                data: {

                    datasets: [

                        {

                            label:
                                "PR",

                            data:
                                data.map(
                                    item => ({

                                        x:
                                            item.index,

                                        y:
                                            item.pr

                                    })
                                ),

                            pointRadius: 4,

                            pointHoverRadius: 7,

                            backgroundColor:
                                "#27A5AD"

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,

                    plugins: {

                        legend: {

                            display: false

                        }

                    },


                    scales: {

                        x: {

                            title: {

                                display: true,

                                text:
                                    "Day"

                            }

                        },

                        y: {

                            title: {

                                display: true,

                                text:
                                    "PR (%)"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   OPERATING HOURS
   Daily_KPI Column I
========================================================= */

function buildOperatingHoursChart() {

    const data =
        getDailyKPIData();


    const valid =
        data.filter(
            item =>
                item.operatingHours !== null
        );


    destroyChart("hours");


    const canvas =
        document.getElementById(
            "hoursChart"
        );


    if (!canvas) {
        return;
    }


    charts.hours =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    labels:
                        valid.map(
                            item =>
                                item.index
                        ),

                    datasets: [

                        {

                            label:
                                "Operating Hours",

                            data:
                                valid.map(
                                    item =>
                                        item.operatingHours
                                ),

                            borderColor:
                                "#27A5AD",

                            backgroundColor:
                                "rgba(39,165,173,0.10)",

                            fill: true,

                            tension: 0.3,

                            pointRadius: 3,

                            pointHoverRadius: 6

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,

                    interaction: {

                        mode: "index",

                        intersect: false

                    },

                    plugins: {

                        legend: {

                            display: false

                        }

                    },


                    scales: {

                        x: {

                            title: {

                                display: true,

                                text:
                                    "Day of Month"

                            }

                        },

                        y: {

                            title: {

                                display: true,

                                text:
                                    "Operating Hours"

                            },

                            beginAtZero: true

                        }

                    }

                }

            }
        );

}


/* =========================================================
   SYSTEM LOSS CHART
   Daily_KPI Column AD
========================================================= */

function buildSystemLossChart() {

    const data =
        getDailyKPIData();


    const valid =
        data.filter(
            item =>
                item.systemLoss !== null
        );


    destroyChart("loss");


    const canvas =
        document.getElementById(
            "lossChart"
        );


    if (!canvas) {
        return;
    }


    charts.loss =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    labels:
                        valid.map(
                            item =>
                                item.index
                        ),

                    datasets: [

                        {

                            label:
                                "System Losses (%)",

                            data:
                                valid.map(
                                    item =>
                                        item.systemLoss
                                ),

                            borderColor:
                                "#27A5AD",

                            backgroundColor:
                                "rgba(39,165,173,0.12)",

                            fill: true,

                            tension: 0.25,

                            pointRadius: 3,

                            pointHoverRadius: 6

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,

                    interaction: {

                        mode: "index",

                        intersect: false

                    },

                    plugins: {

                        legend: {

                            display: false

                        }

                    },


                    scales: {

                        x: {

                            title: {

                                display: true,

                                text:
                                    "Day of Month"

                            }

                        },

                        y: {

                            title: {

                                display: true,

                                text:
                                    "System Loss (%)"

                            }

                        }

                    }

                }

            }
        );


    buildDashboardLossChart(
        valid
    );

}


/* =========================================================
   DASHBOARD SYSTEM LOSS
========================================================= */

function buildDashboardLossChart(data) {

    destroyChart(
        "dashboardLoss"
    );


    const canvas =
        document.getElementById(
            "dashboardLossChart"
        );


    if (!canvas) {
        return;
    }


    charts.dashboardLoss =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    labels:
                        data.map(
                            item =>
                                item.index
                        ),

                    datasets: [

                        {

                            label:
                                "System Loss",

                            data:
                                data.map(
                                    item =>
                                        item.systemLoss
                                ),

                            borderColor:
                                "#27A5AD",

                            backgroundColor:
                                "rgba(39,165,173,0.1)",

                            fill: true,

                            tension: 0.25,

                            pointRadius: 2

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,

                    plugins: {

                        legend: {

                            display: false

                        }

                    },

                    scales: {

                        x: {

                            title: {

                                display: true,

                                text:
                                    "Day"

                            }

                        },

                        y: {

                            title: {

                                display: true,

                                text:
                                    "Loss (%)"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   CURTAILMENT DATA
========================================================= */

function getCurtailmentData() {

    const rows =
        sheetData[
            "Curtailment records"
        ];


    if (
        !rows ||
        rows.length === 0
    ) {

        return [];

    }


    /*
       We deliberately do not assume a
       specific column number here.

       The function looks for common
       curtailment-loss headers.
    */


    const headers =
        rows[0] || [];


    let dateIndex =
        findColumn(
            headers,
            [
                "date",
                "day",
                "timestamp",
                "time",
                "start time"
            ]
        );


    let lossIndex =
        findColumn(
            headers,
            [
                "curtailment loss",
                "curtailment losses",
                "loss",
                "losses",
                "curtailed energy",
                "energy loss",
                "mwh"
            ]
        );


    /*
       If no header match exists,
       find numeric columns automatically.
    */

    if (
        lossIndex === -1
    ) {

        lossIndex =
            findLikelyNumericColumn(
                rows,
                1
            );

    }


    const result = [];


    for (
        let i = 1;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i];


        if (!row) {
            continue;
        }


        const loss =
            lossIndex >= 0
                ? parseNumber(
                    row[lossIndex]
                )
                : null;


        if (
            loss === null
        ) {

            continue;

        }


        let date =
            dateIndex >= 0
                ? row[dateIndex]
                : i;


        date =
            convertExcelDate(
                date
            );


        result.push({

            date,

            loss

        });

    }


    return result;

}


/* =========================================================
   CURTAILMENT CHART
========================================================= */

function buildCurtailmentChart() {

    const data =
        getCurtailmentData();


    destroyChart(
        "curtailment"
    );


    const canvas =
        document.getElementById(
            "curtailmentChart"
        );


    if (!canvas) {
        return;
    }


    charts.curtailment =
        new Chart(
            canvas,
            {

                type: "line",

                data: {

                    labels:
                        data.map(
                            item =>
                                formatDate(
                                    item.date
                                )
                        ),

                    datasets: [

                        {

                            label:
                                "Curtailment Loss",

                            data:
                                data.map(
                                    item =>
                                        item.loss
                                ),

                            borderColor:
                                "#27A5AD",

                            backgroundColor:
                                "rgba(39,165,173,0.10)",

                            fill: true,

                            tension: 0.25,

                            pointRadius: 3,

                            pointHoverRadius: 7

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,

                    interaction: {

                        mode: "index",

                        intersect: false

                    },

                    plugins: {

                        legend: {

                            display: false

                        },

                        tooltip: {

                            callbacks: {

                                label:
                                    context =>
                                        "Curtailment Loss: " +
                                        Number(
                                            context.parsed.y
                                        ).toFixed(2)

                            }

                        }

                    },


                    scales: {

                        x: {

                            title: {

                                display: true,

                                text:
                                    "Date"

                            }

                        },

                        y: {

                            title: {

                                display: true,

                                text:
                                    "Curtailment Loss"

                            },

                            beginAtZero: true

                        }

                    }

                }

            }
        );


    const total =
        data.reduce(
            (
                sum,
                item
            ) =>
                sum + item.loss,
            0
        );


    document.getElementById(
        "curtailmentSummary"
    ).textContent =
        data.length +
        " records · Total loss: " +
        total.toFixed(2);

}


/* =========================================================
   ENERGY DATA
   Annual_KPI:

   Budgeted Energy = E10:E21
   Measured Energy = F10:F21

   Excel row 10 = JS index 9
   Excel row 21 = JS index 20
========================================================= */

function getEnergyData() {

    const rows =
        sheetData[
            "Annual_KPI"
        ];


    if (
        !rows ||
        rows.length < 21
    ) {

        return [];

    }


    const result = [];


    for (
        let excelRow = 10;
        excelRow <= 21;
        excelRow++
    ) {

        const row =
            rows[
                excelRow - 1
            ];


        if (!row) {
            continue;
        }


        const budget =
            parseNumber(
                row[4]
            );


        const measured =
            parseNumber(
                row[5]
            );


        /*
           Try to obtain month name.

           Usually it will be in one
           of the first four columns.
        */

        let month = null;


        for (
            let c = 0;
            c < 4;
            c++
        ) {

            const value =
                row[c];


            if (
                value !== null &&
                value !== undefined &&
                value !== ""
            ) {

                const text =
                    String(
                        value
                    ).trim();


                if (
                    text.length <= 15
                ) {

                    month =
                        text;

                    break;

                }

            }

        }


        if (!month) {

            month =
                "Month " +
                (
                    excelRow - 9
                );

        }


        result.push({

            month,

            budget,

            measured

        });

    }


    return result;

}


/* =========================================================
   ENERGY CHART
========================================================= */

function buildEnergyChart() {

    const data =
        getEnergyData();


    destroyChart(
        "energy"
    );


    const canvas =
        document.getElementById(
            "energyChart"
        );


    if (!canvas) {
        return;
    }


    charts.energy =
        new Chart(
            canvas,
            {

                type: "bar",

                data: {

                    labels:
                        data.map(
                            item =>
                                item.month
                        ),

                    datasets: [

                        {

                            label:
                                "Budgeted Energy",

                            data:
                                data.map(
                                    item =>
                                        item.budget
                                ),

                            backgroundColor:
                                "#17252A",

                            borderRadius: 4,

                            barPercentage:
                                0.65,

                            categoryPercentage:
                                0.7

                        },


                        {

                            label:
                                "Measured Energy",

                            data:
                                data.map(
                                    item =>
                                        item.measured
                                ),

                            backgroundColor:
                                "#27A5AD",

                            borderRadius: 4,

                            barPercentage:
                                0.65,

                            categoryPercentage:
                                0.7

                        }

                    ]

                },


                options: {

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,

                    interaction: {

                        mode: "index",

                        intersect: false

                    },

                    plugins: {

                        legend: {

                            position:
                                "top",

                            labels: {

                                boxWidth: 10,

                                font: {

                                    size: 9

                                }

                            }

                        }

                    },


                    scales: {

                        x: {

                            title: {

                                display: true,

                                text:
                                    "Month"

                            },

                            grid: {

                                display:
                                    false

                            }

                        },

                        y: {

                            title: {

                                display: true,

                                text:
                                    "Energy"

                            },

                            beginAtZero: true

                        }

                    }

                }

            }
        );


    updateEnergySummary(
        data
    );

}


/* =========================================================
   ENERGY SUMMARY
========================================================= */

function updateEnergySummary(
    data
) {

    const budget =
        data.reduce(
            (
                sum,
                item
            ) =>
                sum +
                (
                    item.budget ||
                    0
                ),
            0
        );


    const measured =
        data.reduce(
            (
                sum,
                item
            ) =>
                sum +
                (
                    item.measured ||
                    0
                ),
            0
        );


    const variance =
        measured - budget;


    document.getElementById(
        "totalBudget"
    ).textContent =
        budget.toFixed(2);


    document.getElementById(
        "totalMeasured"
    ).textContent =
        measured.toFixed(2);


    document.getElementById(
        "energyVariance"
    ).textContent =
        (
            variance >= 0
                ? "+"
                : ""
        ) +
        variance.toFixed(2);

}


/* =========================================================
   PA DATA

   The PA sheet can have different
   column arrangements.

   We identify Start / End / Duration
   by header names.
========================================================= */

function getPAData() {

    const rows =
        sheetData["PA"];


    if (
        !rows ||
        rows.length < 2
    ) {

        return [];

    }


    const headers =
        rows[0] || [];


    const startIndex =
        findColumn(
            headers,
            [
                "breakdown start",
                "start time",
                "start",
                "from",
                "outage start",
                "failure start"
            ]
        );


    const endIndex =
        findColumn(
            headers,
            [
                "breakdown end",
                "end time",
                "end",
                "to",
                "outage end",
                "failure end"
            ]
        );


    const durationIndex =
        findColumn(
            headers,
            [
                "duration",
                "breakdown duration",
                "downtime",
                "outage duration"
            ]
        );


    const equipmentIndex =
        findColumn(
            headers,
            [
                "equipment",
                "asset",
                "plant",
                "device",
                "element"
            ]
        );


    const result = [];


    for (
        let i = 1;
        i < rows.length;
        i++
    ) {

        const row =
            rows[i];


        if (!row) {
            continue;
        }


        let start =
            startIndex >= 0
                ? convertExcelDate(
                    row[startIndex]
                )
                : null;


        let end =
            endIndex >= 0
                ? convertExcelDate(
                    row[endIndex]
                )
                : null;


        let duration =
            durationIndex >= 0
                ? parseDuration(
                    row[durationIndex]
                )
                : null;


        /*
           If start and end are present,
           calculate duration automatically.
        */

        if (
            start instanceof Date &&
            end instanceof Date
        ) {

            duration =
                (
                    end.getTime() -
                    start.getTime()
                ) /
                3600000;

        }


        if (
            !start &&
            !end
        ) {

            continue;

        }


        result.push({

            equipment:
                equipmentIndex >= 0
                    ? String(
                        row[equipmentIndex] ??
                        "Breakdown"
                    )
                    : "Breakdown " +
                      i,

            start,

            end,

            duration

        });

    }


    return result;

}


/* =========================================================
   PA TIMELINE CHART

   Chart.js floating horizontal bars
========================================================= */

function buildPAChart() {

    const data =
        getPAData();


    destroyChart(
        "pa"
    );


    const canvas =
        document.getElementById(
            "paChart"
        );


    if (!canvas) {
        return;
    }


    /*
       Convert dates into timestamps.

       For entries without valid
       dates, use sequential values.
    */

    const valid =
        data.filter(
            item =>
                item.start instanceof Date
        );


    if (
        valid.length === 0
    ) {

        showPAEmptyMessage();

        return;

    }


    const minTime =
        Math.min(
            ...valid.map(
                item =>
                    item.start.getTime()
            )
        );


    const maxTime =
        Math.max(
            ...valid.map(
                item =>
                    (
                        item.end instanceof Date
                            ? item.end.getTime()
                            : item.start.getTime()
                    )
            )
        );


    /*
       Chart.js horizontal floating bars
       use x values as [start,end].
    */

    charts.pa =
        new Chart(
            canvas,
            {

                type: "bar",

                data: {

                    labels:
                        valid.map(
                            item =>
                                item.equipment
                        ),

                    datasets: [

                        {

                            label:
                                "Breakdown Duration",

                            data:
                                valid.map(
                                    item => {

                                        const start =
                                            item.start.getTime();


                                        const end =
                                            item.end instanceof Date
                                                ? item.end.getTime()
                                                : start;


                                        return [
                                            start,
                                            end
                                        ];

                                    }
                                ),

                            backgroundColor:
                                "#27A5AD",

                            borderRadius: 4,

                            barThickness: 22

                        }

                    ]

                },


                options: {

                    indexAxis: "y",

                    responsive: true,

                    maintainAspectRatio:
                        false,

                    animation: false,

                    scales: {

                        x: {

                            type: "linear",

                            min:
                                minTime,

                            max:
                                maxTime,

                            ticks: {

                                callback:
                                    value =>
                                        formatDateTime(
                                            new Date(
                                                value
                                            )
                                        )

                            },

                            title: {

                                display: true,

                                text:
                                    "Breakdown Timeline"

                            },

                            grid: {

                                color:
                                    "#edf2f3"

                            }

                        },

                        y: {

                            title: {

                                display: true,

                                text:
                                    "Equipment"

                            },

                            grid: {

                                display:
                                    false

                            }

                        }

                    },


                    plugins: {

                        legend: {

                            display: false

                        },

                        tooltip: {

                            callbacks: {

                                title:
                                    tooltipPATitle,

                                label:
                                    tooltipPALabel

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   PA TOOLTIP
========================================================= */

function tooltipPATitle(
    context
) {

    const index =
        context[0].dataIndex;


    const data =
        getPAData();


    const item =
        data[index];


    if (!item) {

        return "";

    }


    return item.equipment;

}


function tooltipPALabel(
    context
) {

    const index =
        context.dataIndex;


    const data =
        getPAData();


    const item =
        data[index];


    if (!item) {

        return "";

    }


    const lines = [];


    if (
        item.start instanceof Date
    ) {

        lines.push(
            "Start: " +
            formatDateTime(
                item.start
            )
        );

    }


    if (
        item.end instanceof Date
    ) {

        lines.push(
            "End: " +
            formatDateTime(
                item.end
            )
        );

    }


    if (
        item.duration !== null
    ) {

        lines.push(
            "Duration: " +
            formatDuration(
                item.duration
            )
        );

    }


    return lines;

}


/* =========================================================
   PA EMPTY MESSAGE
========================================================= */

function showPAEmptyMessage() {

    const canvas =
        document.getElementById(
            "paChart"
        );


    const ctx =
        canvas.getContext(
            "2d"
        );


    ctx.clearRect(
        0,
        0,
        canvas.width,
        canvas.height
    );


    ctx.font =
        "12px Inter";


    ctx.fillStyle =
        "#8b999c";


    ctx.textAlign =
        "center";


    ctx.fillText(
        "No valid breakdown start/end times found in PA worksheet.",
        canvas.width / 2,
        canvas.height / 2
    );

}


/* =========================================================
   FIND COLUMN
========================================================= */

function findColumn(
    headers,
    possibleNames
) {

    if (!headers) {
        return -1;
    }


    const normalizedHeaders =
        headers.map(
            header =>
                normalizeText(
                    header
                )
        );


    const normalizedNames =
        possibleNames.map(
            name =>
                normalizeText(
                    name
                )
        );


    /*
       Exact match first.
    */

    for (
        let i = 0;
        i < normalizedHeaders.length;
        i++
    ) {

        if (
            normalizedNames.includes(
                normalizedHeaders[i]
            )
        ) {

            return i;

        }

    }


    /*
       Partial match second.
    */

    for (
        let i = 0;
        i < normalizedHeaders.length;
        i++
    ) {

        for (
            const name of normalizedNames
        ) {

            if (
                normalizedHeaders[i]
                    .includes(name)
                ||
                name.includes(
                    normalizedHeaders[i]
                )
            ) {

                return i;

            }

        }

    }


    return -1;

}


/* =========================================================
   FIND LIKELY NUMERIC COLUMN
========================================================= */

function findLikelyNumericColumn(
    rows,
    startRow
) {

    if (!rows.length) {
        return -1;
    }


    const maxColumns =
        Math.max(
            ...rows.map(
                row =>
                    row
                        ? row.length
                        : 0
            )
        );


    let bestIndex = -1;

    let bestScore = 0;


    for (
        let c = 0;
        c < maxColumns;
        c++
    ) {

        let numericCount = 0;


        for (
            let r = startRow;
            r < rows.length;
            r++
        ) {

            if (
                parseNumber(
                    rows[r]?.[c]
                ) !== null
            ) {

                numericCount++;

            }

        }


        if (
            numericCount >
            bestScore
        ) {

            bestScore =
                numericCount;

            bestIndex =
                c;

        }

    }


    return bestIndex;

}


/* =========================================================
   CONVERT EXCEL DATE
========================================================= */

function convertExcelDate(
    value
) {

    if (
        value instanceof Date &&
        !isNaN(value)
    ) {

        return value;

    }


    if (
        typeof value === "number"
    ) {

        if (
            value > 20000 &&
            value < 60000
        ) {

            const parsed =
                XLSX.SSF.parse_date_code(
                    value
                );


            if (parsed) {

                return new Date(
                    parsed.y,
                    parsed.m - 1,
                    parsed.d,
                    parsed.H || 0,
                    parsed.M || 0,
                    parsed.S || 0
                );

            }

        }

    }


    if (
        typeof value === "string"
    ) {

        const parsed =
            new Date(value);


        if (
            !isNaN(parsed)
        ) {

            return parsed;

        }

    }


    return value;

}


/* =========================================================
   PARSE DURATION
========================================================= */

function parseDuration(
    value
) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return null;

    }


    if (
        typeof value === "number"
    ) {

        /*
           Excel time duration may be
           represented as a fraction
           of a day.
        */

        if (
            value > 0 &&
            value < 1
        ) {

            return value * 24;

        }


        return value;

    }


    const text =
        String(value)
            .trim();


    /*
       HH:MM
    */

    const match =
        text.match(
            /^(\d+):(\d{1,2})(?::(\d{1,2}))?$/
        );


    if (match) {

        const hours =
            Number(
                match[1]
            );


        const minutes =
            Number(
                match[2]
            );


        const seconds =
            Number(
                match[3] || 0
            );


        return (
            hours +
            minutes / 60 +
            seconds / 3600
        );

    }


    return parseNumber(
        text
    );

}


/* =========================================================
   FORMAT DATETIME
========================================================= */

function formatDateTime(
    value
) {

    if (
        !(value instanceof Date) ||
        isNaN(value)
    ) {

        return String(value);

    }


    return value.toLocaleString(
        "en-IN",
        {

            day: "2-digit",

            month: "short",

            hour: "2-digit",

            minute: "2-digit"

        }
    );

}


/* =========================================================
   FORMAT DURATION
========================================================= */

function formatDuration(
    hours
) {

    if (
        hours === null ||
        !Number.isFinite(hours)
    ) {

        return "—";

    }


    const totalMinutes =
        Math.round(
            hours * 60
        );


    const h =
        Math.floor(
            totalMinutes / 60
        );


    const m =
        totalMinutes % 60;


    return (
        h +
        "h " +
        m +
        "m"
    );

}


/* =========================================================
   UPDATE KPI CARDS
========================================================= */

function updateKPIs() {

    const daily =
        getDailyKPIData();


    /*
       Latest valid value.
    */

    const latestPR =
        getLatest(
            daily,
            "pr"
        );


    const latestHours =
        getLatest(
            daily,
            "operatingHours"
        );


    const latestLoss =
        getLatest(
            daily,
            "systemLoss"
        );


    document.getElementById(
        "dashboardPR"
    ).textContent =
        latestPR !== null
            ? latestPR.toFixed(2) + "%"
            : "—";


    document.getElementById(
        "dashboardHours"
    ).textContent =
        latestHours !== null
            ? latestHours.toFixed(2) + " h"
            : "—";


    document.getElementById(
        "dashboardLoss"
    ).textContent =
        latestLoss !== null
            ? latestLoss.toFixed(2) + "%"
            : "—";


    /*
       PA is calculated from PA breakdown
       duration rather than assuming a
       particular PA column.
    */

    const pa =
        getPAData();


    const totalBreakdownHours =
        pa.reduce(
            (
                sum,
                item
            ) =>
                sum +
                (
                    item.duration || 0
                ),
            0
        );


    /*
       The dashboard card shows
       breakdown-based availability
       only when enough information
       exists.

       Otherwise it remains "—".
    */

    const paValue =
        calculatePA(
            pa
        );


    document.getElementById(
        "dashboardPA"
    ).textContent =
        paValue !== null
            ? paValue.toFixed(2) + "%"
            : "—";

}


/* =========================================================
   GET LATEST
========================================================= */

function getLatest(
    data,
    property
) {

    for (
        let i = data.length - 1;
        i >= 0;
        i--
    ) {

        const value =
            data[i][property];


        if (
            value !== null &&
            Number.isFinite(value)
        ) {

            return value;

        }

    }


    return null;

}


/* =========================================================
   CALCULATE PA
=========================================================

   Only calculates if a meaningful
   date range can be established.

========================================================= */

function calculatePA(
    pa
) {

    const valid =
        pa.filter(
            item =>
                item.start instanceof Date &&
                item.end instanceof Date
        );


    if (
        valid.length === 0
    ) {

        return null;

    }


    const earliest =
        Math.min(
            ...valid.map(
                item =>
                    item.start.getTime()
            )
        );


    const latest =
        Math.max(
            ...valid.map(
                item =>
                    item.end.getTime()
            )
        );


    const totalHours =
        (
            latest -
            earliest
        ) /
        3600000;


    if (
        totalHours <= 0
    ) {

        return null;

    }


    const downtime =
        valid.reduce(
            (
                sum,
                item
            ) =>
                sum +
                (
                    item.duration ||
                    0
                ),
            0
        );


    return Math.max(
        0,
        Math.min(
            100,
            (
                (
                    totalHours -
                    downtime
                ) /
                totalHours
            ) *
            100
        )
    );

}


/* =========================================================
   RESET DASHBOARD
========================================================= */

function resetDashboard() {

    workbook = null;

    sheetData = {};

    uploadedFile = null;


    Object.keys(charts)
        .forEach(
            key =>
                destroyChart(
                    key
                )
        );


    fileInput.value = "";


    fileInfo.classList.add(
        "hidden"
    );


    workbookStatus.classList.add(
        "hidden"
    );


    emptyState.classList.remove(
        "hidden"
    );


    hideAnalytics();


    sidebarFileName.textContent =
        "No DGR uploaded";


    statusText.textContent =
        "Upload a DGR to generate the analytics.";

}


/* =========================================================
   END
========================================================= */
