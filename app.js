/* =========================================================
   SOLAR DGR ANALYTICS
   app.js
   Complete DGR reader + Chart.js renderer

   EXACT WORKBOOK MAPPING

   PA worksheet:
     W  = Issue / Fault
     Z  = Fault Start Time
     AC = Work Completion Time
     AG = Breakdown Time (minutes)
     AL = System Loss (MWh)

   Daily_KPI worksheet:
     B  = Date
     I  = Operating Hours
     S  = PA (%)
     V  = PR (%)
     AD = System Losses (%)

   Curtailment records worksheet:
     C  = Date
     H  = Start Time
     I  = End Time
     R  = Loss of Generation MWh

   Annual_KPI worksheet:
     E10:E21 = Budgeted Energy
     F10:F21 = Measured Energy

   ========================================================= */

"use strict";


/* =========================================================
   GLOBAL VARIABLES
========================================================= */

let workbook = null;

const charts = {};


/* =========================================================
   DOM HELPER
========================================================= */

function $(id) {
    return document.getElementById(id);
}


/* =========================================================
   INITIALISATION
========================================================= */

document.addEventListener("DOMContentLoaded", () => {

    setupNavigation();

    setupUpload();

    setupRemoveButton();

    hideAnalytics();

});


/* =========================================================
   NAVIGATION
========================================================= */

function setupNavigation() {

    const buttons =
        document.querySelectorAll(".nav-item");

    buttons.forEach(button => {

        button.addEventListener("click", () => {

            buttons.forEach(item => {
                item.classList.remove("active");
            });

            button.classList.add("active");

            const target =
                $(button.dataset.target);

            if (target) {

                target.scrollIntoView({
                    behavior: "smooth",
                    block: "start"
                });

            }

        });

    });

}


/* =========================================================
   UPLOAD
========================================================= */

function setupUpload() {

    const input =
        $("dgrFile");

    const dropZone =
        $("dropZone");


    if (input) {

        input.addEventListener(
            "change",
            event => {

                const file =
                    event.target.files?.[0];

                if (file) {

                    processFile(file);

                }

            }
        );

    }


    /*
       Drag/drop area may still exist in the HTML.
       It does not create an additional upload button.
    */

    if (dropZone) {

        dropZone.addEventListener(
            "click",
            () => {

                if (input) {
                    input.click();
                }

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
                    event.dataTransfer.files?.[0];


                if (file) {

                    processFile(file);

                }

            }
        );

    }

}


/* =========================================================
   REMOVE BUTTON
========================================================= */

function setupRemoveButton() {

    const button =
        $("removeFile");


    if (!button) return;


    button.addEventListener(
        "click",
        resetDashboard
    );

}


/* =========================================================
   PROCESS FILE
========================================================= */

function processFile(file) {

    if (!file) return;


    const extension =
        file.name
            .split(".")
            .pop()
            .toLowerCase();


    if (
        !["xlsx", "xls", "csv"].includes(
            extension
        )
    ) {

        alert(
            "Please upload a valid Excel file (.xlsx/.xls) or CSV file."
        );

        return;

    }


    if (
        typeof XLSX === "undefined"
    ) {

        alert(
            "Excel reader could not be loaded. Please check the SheetJS script."
        );

        return;

    }


    setStatus(
        "Reading DGR workbook..."
    );


    const reader =
        new FileReader();


    reader.onload =
        event => {

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
                            cellDates: true,
                            cellNF: true,
                            cellText: true
                        }
                    );


                if (
                    !workbook ||
                    !workbook.SheetNames ||
                    workbook.SheetNames.length === 0
                ) {

                    throw new Error(
                        "No worksheets found."
                    );

                }


                updateFileInformation(
                    file
                );


                showAnalytics();


                renderAll();


                setStatus(
                    `${file.name} loaded successfully.`
                );

            }

            catch (error) {

                console.error(
                    "DGR ERROR:",
                    error
                );


                setStatus(
                    "Could not read the DGR."
                );


                alert(
                    "Could not read the uploaded DGR.\n\n" +
                    error.message
                );

            }

        };


    reader.onerror =
        () => {

            setStatus(
                "Could not read the selected file."
            );

        };


    reader.readAsArrayBuffer(
        file
    );

}


/* =========================================================
   FILE INFORMATION
========================================================= */

function updateFileInformation(file) {

    if ($("fileName")) {

        $("fileName").textContent =
            file.name;

    }


    if ($("fileSheets")) {

        $("fileSheets").textContent =
            `${workbook.SheetNames.length} worksheets detected`;

    }


    if ($("sidebarFileName")) {

        $("sidebarFileName").textContent =
            file.name;

    }


    if ($("fileInfo")) {

        $("fileInfo").classList.remove(
            "hidden"
        );

    }


    if ($("workbookStatus")) {

        $("workbookStatus").classList.remove(
            "hidden"
        );

    }


    renderSheetBadges();

}


/* =========================================================
   WORKSHEET BADGES
========================================================= */

function renderSheetBadges() {

    const container =
        $("sheetBadges");


    if (
        !container ||
        !workbook
    ) {
        return;
    }


    const expected = [
        "Dashboard",
        "Annual_KPI",
        "Daily_KPI",
        "PA",
        "Curtailment records"
    ];


    container.innerHTML = "";


    expected.forEach(
        name => {

            const badge =
                document.createElement(
                    "span"
                );


            badge.className =
                "sheet-badge";


            if (
                workbook.SheetNames.includes(
                    name
                )
            ) {

                badge.textContent =
                    `${name} ✓`;

            }

            else {

                badge.textContent =
                    `${name} — missing`;

                badge.classList.add(
                    "missing"
                );

            }


            container.appendChild(
                badge
            );

        }
    );

}


/* =========================================================
   GET SHEET
========================================================= */

function getSheet(name) {

    if (
        !workbook ||
        !workbook.Sheets
    ) {

        return null;

    }


    if (
        workbook.Sheets[name]
    ) {

        return workbook.Sheets[name];

    }


    /*
       Case-insensitive fallback.
    */

    const actualName =
        workbook.SheetNames.find(
            sheetName =>
                sheetName.toLowerCase() ===
                name.toLowerCase()
        );


    if (!actualName) {
        return null;
    }


    return workbook.Sheets[
        actualName
    ];

}


/* =========================================================
   SHEET -> MATRIX
========================================================= */

function sheetToMatrix(
    sheet
) {

    if (!sheet) {
        return [];
    }


    return XLSX.utils.sheet_to_json(
        sheet,
        {
            header: 1,
            defval: null,
            raw: true,
            blankrows: false
        }
    );

}


/* =========================================================
   EXCEL COLUMN -> ARRAY INDEX
========================================================= */

function colIndex(column) {

    let result = 0;


    for (
        const character of column.toUpperCase()
    ) {

        result =
            result * 26 +
            character.charCodeAt(0) -
            64;

    }


    return result - 1;

}


/* =========================================================
   READ COLUMN FROM RAW ROW
========================================================= */

function getCell(
    row,
    column
) {

    if (!row) {
        return null;
    }


    return row[
        colIndex(column)
    ];

}


/* =========================================================
   NUMBER PARSER
========================================================= */

function parseNumber(
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

        return Number.isFinite(
            value
        )
            ? value
            : null;

    }


    if (
        value instanceof Date
    ) {

        return null;

    }


    const text =
        String(value)
            .trim()
            .replace(
                /,/g,
                ""
            )
            .replace(
                /%/g,
                ""
            );


    if (!text) {
        return null;
    }


    if (
        /^#(REF|DIV\/0|VALUE|N\/A|NAME|NUM|NULL)!?/i
            .test(text)
    ) {

        return null;

    }


    const number =
        Number(text);


    return Number.isFinite(
        number
    )
        ? number
        : null;

}


/* =========================================================
   DATE PARSER
========================================================= */

function parseDate(
    value
) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return null;

    }


    /*
       JavaScript Date
    */

    if (
        value instanceof Date
    ) {

        if (
            !isNaN(
                value.getTime()
            )
        ) {

            return new Date(
                value.getTime()
            );

        }

        return null;

    }


    /*
       Excel serial date
    */

    if (
        typeof value === "number"
    ) {

        try {

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

        catch (error) {

            console.warn(
                "Date conversion failed:",
                value
            );

        }


        return null;

    }


    /*
       String date
    */

    const text =
        String(value)
            .trim();


    /*
       DD/MM/YYYY
       DD-MM-YYYY
    */

    let match =
        text.match(
            /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
        );


    if (match) {

        const day =
            Number(
                match[1]
            );


        const month =
            Number(
                match[2]
            ) - 1;


        const year =
            Number(
                match[3]
            );


        const date =
            new Date(
                year,
                month,
                day
            );


        if (
            date.getFullYear() === year &&
            date.getMonth() === month &&
            date.getDate() === day
        ) {

            return date;

        }

    }


    /*
       DD-MMM-YYYY
    */

    match =
        text.match(
            /^(\d{1,2})[\/\-]([A-Za-z]{3,9})[\/\-](\d{2,4})/
        );


    if (match) {

        const months = [
            "jan",
            "feb",
            "mar",
            "apr",
            "may",
            "jun",
            "jul",
            "aug",
            "sep",
            "oct",
            "nov",
            "dec"
        ];


        const month =
            months.indexOf(
                match[2]
                    .toLowerCase()
                    .substring(
                        0,
                        3
                    )
            );


        let year =
            Number(
                match[3]
            );


        if (year < 100) {
            year += 2000;
        }


        if (month >= 0) {

            return new Date(
                year,
                month,
                Number(
                    match[1]
                )
            );

        }

    }


    /*
       ISO / browser date
    */

    const parsed =
        new Date(text);


    if (
        !isNaN(
            parsed.getTime()
        )
    ) {

        return parsed;

    }


    return null;

}


/* =========================================================
   DATE KEY
========================================================= */

function dateKey(
    date
) {

    return [
        date.getFullYear(),

        String(
            date.getMonth() + 1
        ).padStart(
            2,
            "0"
        ),

        String(
            date.getDate()
        ).padStart(
            2,
            "0"
        )
    ].join("-");

}


/* =========================================================
   DATE DISPLAY
========================================================= */

function displayDate(
    date
) {

    if (
        !(date instanceof Date)
    ) {

        return "";

    }


    return date.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short"
        }
    );

}


/* =========================================================
   TIME TO MINUTES
========================================================= */

function timeToMinutes(
    value
) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return null;

    }


    /*
       Excel time stored as Date
    */

    if (
        value instanceof Date
    ) {

        return (
            value.getHours() * 60 +
            value.getMinutes()
        );

    }


    /*
       Excel time fraction
    */

    if (
        typeof value === "number"
    ) {

        if (
            value >= 0 &&
            value < 1
        ) {

            return Math.round(
                value *
                24 *
                60
            );

        }


        const parsed =
            XLSX.SSF.parse_date_code(
                value
            );


        if (parsed) {

            return (
                parsed.H * 60 +
                parsed.M
            );

        }


        return null;

    }


    const text =
        String(value)
            .trim();


    /*
       HH:MM
       HH:MM:SS
       HH:MM AM/PM
    */

    const match =
        text.match(
            /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i
        );


    if (match) {

        let hour =
            Number(
                match[1]
            );


        const minute =
            Number(
                match[2]
            );


        if (
            match[4]
        ) {

            if (
                match[4].toUpperCase() ===
                "PM" &&
                hour < 12
            ) {

                hour += 12;

            }


            if (
                match[4].toUpperCase() ===
                "AM" &&
                hour === 12
            ) {

                hour = 0;

            }

        }


        if (
            hour <= 23 &&
            minute <= 59
        ) {

            return (
                hour * 60 +
                minute
            );

        }

    }


    return null;

}


/* =========================================================
   MINUTES -> HH:MM
========================================================= */

function minutesToTime(
    minutes
) {

    const safe =
        Math.max(
            0,
            Math.min(
                1439,
                Math.round(
                    minutes
                )
            )
        );


    const hours =
        Math.floor(
            safe / 60
        );


    const mins =
        safe % 60;


    return (
        String(hours).padStart(
            2,
            "0"
        ) +
        ":" +
        String(mins).padStart(
            2,
            "0"
        )
    );

}


/* =========================================================
   FORMAT PERCENT
========================================================= */

function formatPercent(
    value
) {

    if (
        value === null ||
        !Number.isFinite(value)
    ) {

        return "—";

    }


    /*
       Values such as 0.842 are
       interpreted as 84.2%.

       Values such as 84.2 are
       already percentages.
    */

    const percentage =
        Math.abs(value) <= 1.5
            ? value * 100
            : value;


    return (
        percentage.toFixed(
            2
        ) +
        "%"
    );

}


/* =========================================================
   CHART DESTROY
========================================================= */

function destroyChart(
    key
) {

    if (
        charts[key] &&
        typeof charts[key].destroy ===
        "function"
    ) {

        try {
            charts[key].destroy();
        }

        catch (error) {

            console.warn(
                "Chart destroy error:",
                key,
                error
            );

        }

    }


    charts[key] = null;

}


/* =========================================================
   DESTROY ALL
========================================================= */

function destroyAllCharts() {

    Object.keys(charts)
        .forEach(
            destroyChart
        );

}


/* =========================================================
   HIDE / SHOW ANALYTICS
========================================================= */

function hideAnalytics() {

    [
        "dashboardSection",
        "paSection",
        "performanceSection",
        "curtailmentSection",
        "energySection"
    ]
        .forEach(
            id => {

                const section =
                    $(id);

                if (section) {

                    section.style.display =
                        "none";

                }

            }
        );

}


function showAnalytics() {

    [
        "dashboardSection",
        "paSection",
        "performanceSection",
        "curtailmentSection",
        "energySection"
    ]
        .forEach(
            id => {

                const section =
                    $(id);

                if (section) {

                    section.style.display =
                        "";

                }

            }
        );


    if (emptyState) {

        emptyState.classList.add(
            "hidden"
        );

    }

}


/* =========================================================
   STATUS TEXT
========================================================= */

function setStatus(
    message
) {

    if (statusText) {

        statusText.textContent =
            message;

    }

}


/* =========================================================
   DAILY KPI
========================================================= */

function readDailyKPI() {

    const sheet =
        getSheet(
            "Daily_KPI"
        );


    if (!sheet) {

        console.warn(
            "Daily_KPI worksheet not found."
        );

        return [];

    }


    const rows =
        sheetToMatrix(
            sheet
        );


    const result = [];


    /*
       Exact positions:

       B = date
       I = operating hours
       S = PA
       V = PR
       AD = system loss

       We deliberately use the exact columns
       rather than guessing column positions.
    */

    rows.forEach(
        row => {

            const date =
                parseDate(
                    getCell(
                        row,
                        "B"
                    )
                );


            if (!date) {
                return;
            }


            const pa =
                parseNumber(
                    getCell(
                        row,
                        "S"
                    )
                );


            const pr =
                parseNumber(
                    getCell(
                        row,
                        "V"
                    )
                );


            const hours =
                parseNumber(
                    getCell(
                        row,
                        "I"
                    )
                );


            const loss =
                parseNumber(
                    getCell(
                        row,
                        "AD"
                    )
                );


            if (
                pa === null &&
                pr === null &&
                hours === null &&
                loss === null
            ) {

                return;

            }


            result.push({

                date,

                key:
                    dateKey(
                        date
                    ),

                pa:

                    pa !== null &&
                    Math.abs(pa) <= 1.5
                        ? pa * 100
                        : pa,

                pr:

                    pr !== null &&
                    Math.abs(pr) <= 1.5
                        ? pr * 100
                        : pr,

                hours,

                loss:

                    loss !== null &&
                    Math.abs(loss) <= 1.5
                        ? loss * 100
                        : loss

            });

        }
    );


    result.sort(
        (a, b) =>
            a.date -
            b.date
    );


    return result;

}


/* =========================================================
   DAILY KPI CHARTS
========================================================= */

function renderDailyKPICharts(
    rows
) {

    const labels =
        rows.map(
            row =>
                displayDate(
                    row.date
                )
        );


    /*
       PR
    */

    makeLineChart(
        "prChart",
        labels,
        rows.map(
            row =>
                row.pr
        ),
        "PR (%)",
        "PR (%)"
    );


    /*
       Operating Hours
    */

    makeLineChart(
        "hoursChart",
        labels,
        rows.map(
            row =>
                row.hours
        ),
        "Operating Hours",
        "Operating Hours"
    );


    /*
       System Loss
    */

    makeLineChart(
        "lossChart",
        labels,
        rows.map(
            row =>
                row.loss
        ),
        "System Losses (%)",
        "System Loss (%)"
    );


    /*
       Dashboard PR
    */

    makeLineChart(
        "dashboardPRChart",
        labels,
        rows.map(
            row =>
                row.pr
        ),
        "PR (%)",
        "PR (%)"
    );


    /*
       Dashboard System Loss
    */

    makeLineChart(
        "dashboardLossChart",
        labels,
        rows.map(
            row =>
                row.loss
        ),
        "System Loss (%)",
        "System Loss (%)"
    );

}


/* =========================================================
   GENERIC LINE CHART
========================================================= */

function makeLineChart(
    canvasId,
    labels,
    values,
    datasetLabel,
    yLabel
) {

    const canvas =
        $(canvasId);


    if (!canvas) {
        return;
    }


    destroyChart(
        canvasId
    );


    charts[canvasId] =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type: "line",

                data: {

                    labels,

                    datasets: [

                        {

                            label:
                                datasetLabel,

                            data:
                                values,

                            borderWidth:
                                2,

                            pointRadius:
                                3,

                            pointHoverRadius:
                                6,

                            tension:
                                0.25,

                            fill:
                                false

                        }

                    ]

                },


                options: {

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,

                    interaction: {

                        mode:
                            "index",

                        intersect:
                            false

                    },


                    plugins: {

                        legend: {

                            display:
                                false

                        },

                        tooltip: {

                            callbacks: {

                                label:
                                    context =>
                                        `${datasetLabel}: ${formatNumber(context.raw, 2)}`

                            }

                        }

                    },


                    scales: {

                        x: {

                            grid: {

                                display:
                                    false

                            },

                            ticks: {

                                autoSkip:
                                    true,

                                maxTicksLimit:
                                    12,

                                maxRotation:
                                    0,

                                minRotation:
                                    0

                            }

                        },


                        y: {

                            title: {

                                display:
                                    true,

                                text:
                                    yLabel

                            },

                            beginAtZero:
                                false,

                            ticks: {

                                maxTicksLimit:
                                    7,

                                callback:
                                    value =>
                                        formatNumber(
                                            value,
                                            1
                                        )

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   READ PLANT UNAVAILABILITY
========================================================= */

function readPlantUnavailability() {

    const rows =
        sheetToMatrix(
            getSheet(
                "PA"
            )
        );


    const result = [];


    rows.forEach(
        row => {

            const fault =
                getCell(
                    row,
                    "W"
                );


            const start =
                timeToMinutes(
                    getCell(
                        row,
                        "Z"
                    )
                );


            const end =
                timeToMinutes(
                    getCell(
                        row,
                        "AC"
                    )
                );


            if (
                fault === null ||
                fault === undefined ||
                String(
                    fault
                ).trim() === ""
            ) {

                return;

            }


            if (
                String(
                    fault
                ).trim() ===
                "#REF!"
            ) {

                return;

            }


            if (
                start === null ||
                end === null
            ) {

                return;

            }


            let finish =
                end;


            if (
                finish < start
            ) {

                finish +=
                    24 * 60;

            }


            result.push({

                fault:
                    String(
                        fault
                    ).trim(),

                start,

                end:
                    finish

            });

        }
    );


    return result;

}


/* =========================================================
   PLANT UNAVAILABILITY GANTT
========================================================= */

function renderPlantUnavailability() {

    const canvas =
        $("paChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "paChart"
    );


    const records =
        readPlantUnavailability();


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No plant unavailability records found."
        );

        return;

    }


    const labels =
        records.map(
            record =>
                record.fault
        );


    /*
       Floating horizontal bars.
       X = minutes after midnight.
       30-minute ticks.
    */

    const datasets =
        records.map(
            record => ({

                label:
                    record.fault,

                data: [

                    {

                        x: [
                            record.start,
                            record.end
                        ],

                        y:
                            record.fault

                    }

                ],

                backgroundColor:
                    "rgba(39,165,173,0.72)",

                borderColor:
                    "#27A5AD",

                borderWidth:
                    1,

                borderRadius:
                    4,

                barThickness:
                    22

            })
        );


    charts.paChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels,

                    datasets

                },


                options: {

                    indexAxis:
                        "y",

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,

                    parsing:
                        false,


                    plugins: {

                        legend: {

                            display:
                                false

                        },


                        tooltip: {

                            callbacks: {

                                title:
                                    context =>
                                        records[
                                            context[0]
                                                .dataIndex
                                        ]
                                            ?.fault ||
                                        "",


                                label:
                                    context => {

                                        const record =
                                            records[
                                                context
                                                    .dataIndex
                                            ];


                                        if (!record) {
                                            return "";
                                        }


                                        return [

                                            `Start: ${minutesToTime(record.start)}`,

                                            `End: ${minutesToTime(record.end)}`,

                                            `Duration: ${record.end - record.start} min`

                                        ];

                                    }

                            }

                        }

                    },


                    scales: {

                        x: {

                            type:
                                "linear",

                            min:
                                0,

                            max:
                                1439,

                            title: {

                                display:
                                    true,

                                text:
                                    "Time"

                            },

                            grid: {

                                color:
                                    "rgba(23,37,42,0.07)"

                            },

                            ticks: {

                                stepSize:
                                    30,

                                callback:
                                    value =>
                                        minutesToTime(
                                            value
                                        )

                            }

                        },


                        y: {

                            type:
                                "category",

                            labels,

                            grid: {

                                display:
                                    false

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   READ BREAKDOWN TIME
   PA B + AG
========================================================= */

function readBreakdownData() {

    const rows =
        sheetToMatrix(
            getSheet(
                "PA"
            )
        );


    const map =
        new Map();


    rows.forEach(
        row => {

            const date =
                parseDate(
                    getCell(
                        row,
                        "B"
                    )
                );


            const minutes =
                parseNumber(
                    getCell(
                        row,
                        "AG"
                    )
                );


            if (
                !date ||
                minutes === null
            ) {

                return;

            }


            const key =
                dateKey(
                    date
                );


            if (
                !map.has(
                    key
                )
            ) {

                map.set(
                    key,
                    {

                        date:
                            new Date(
                                date.getFullYear(),
                                date.getMonth(),
                                date.getDate()
                            ),

                        minutes:
                            0

                    }
                );

            }


            map.get(
                key
            ).minutes +=
                minutes;

        }
    );


    return Array.from(
        map.values()
    )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );

}


/* =========================================================
   CREATE BREAKDOWN CANVAS
========================================================= */

function ensureBreakdownCanvas() {

    const section =
        $("paSection");


    if (!section) {
        return null;
    }


    let card =
        $("breakdownTimelineCard");


    if (!card) {

        card =
            document.createElement(
                "div"
            );


        card.id =
            "breakdownTimelineCard";


        card.className =
            "chart-card full-card";


        card.style.marginTop =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        Breakdown Timeline
                    </h3>

                    <span>
                        Same-date breakdowns combined from PA · Column AG
                    </span>

                </div>

                <span class="chart-type">
                    MINUTES
                </span>

            </div>

            <div class="large-chart-container">

                <canvas id="breakdownChart"></canvas>

            </div>

        `;


        const mainCard =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (mainCard) {

            mainCard.insertAdjacentElement(
                "afterend",
                card
            );

        }

        else {

            section.appendChild(
                card
            );

        }

    }


    return $("breakdownChart");

}


/* =========================================================
   BREAKDOWN TIMELINE
========================================================= */

function renderBreakdownTimeline() {

    const canvas =
        ensureBreakdownCanvas();


    if (!canvas) {
        return;
    }


    const records =
        readBreakdownData();


    destroyChart(
        "breakdownChart"
    );


    if (!records.length) {
        return;
    }


    charts.breakdownChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels:
                        records.map(
                            record =>
                                displayDate(
                                    record.date
                                )
                        ),

                    datasets: [

                        {

                            label:
                                "Breakdown Time",

                            data:
                                records.map(
                                    record =>
                                        record.minutes
                                ),

                            borderWidth:
                                1,

                            borderRadius:
                                3

                        }

                    ]

                },


                options: {

                    indexAxis:
                        "y",

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,


                    plugins: {

                        legend: {

                            display:
                                false

                        }


                    },


                    scales: {

                        x: {

                            min:
                                0,

                            max:
                                13,

                            title: {

                                display:
                                    true,

                                text:
                                    "Breakdown Time (minutes)"

                            },

                            ticks: {

                                stepSize:
                                    1

                            }

                        },


                        y: {

                            grid: {

                                display:
                                    false

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   SYSTEM LOSS MWh
   PA B + AL
========================================================= */

function readSystemLossMWh() {

    const rows =
        sheetToMatrix(
            getSheet(
                "PA"
            )
        );


    const map =
        new Map();


    rows.forEach(
        row => {

            const date =
                parseDate(
                    getCell(
                        row,
                        "B"
                    )
                );


            const loss =
                parseNumber(
                    getCell(
                        row,
                        "AL"
                    )
                );


            if (
                !date ||
                loss === null
            ) {

                return;

            }


            const key =
                dateKey(
                    date
                );


            if (
                !map.has(
                    key
                )
            ) {

                map.set(
                    key,
                    {

                        date:
                            new Date(
                                date.getFullYear(),
                                date.getMonth(),
                                date.getDate()
                            ),

                        loss:
                            0

                    }
                );

            }


            map.get(
                key
            ).loss +=
                loss;

        }
    );


    return Array.from(
        map.values()
    )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );

}


/* =========================================================
   SYSTEM LOSS CANVAS
========================================================= */

function ensureSystemLossCanvas() {

    const section =
        $("paSection");


    if (!section) {
        return null;
    }


    let card =
        $("systemLossMwhCard");


    if (!card) {

        card =
            document.createElement(
                "div"
            );


        card.id =
            "systemLossMwhCard";


        card.className =
            "chart-card full-card";


        card.style.marginTop =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        System Loss
                    </h3>

                    <span>
                        Same-date losses combined from PA · Column AL
                    </span>

                </div>

                <span class="chart-type">
                    MWh
                </span>

            </div>

            <div class="large-chart-container">

                <canvas id="systemLossMwhChart"></canvas>

            </div>

        `;


        const breakdown =
            $("breakdownTimelineCard");


        if (breakdown) {

            breakdown.insertAdjacentElement(
                "afterend",
                card
            );

        }

        else {

            section.appendChild(
                card
            );

        }

    }


    return $("systemLossMwhChart");

}


/* =========================================================
   SYSTEM LOSS CHART
========================================================= */

function renderSystemLossMWh() {

    const canvas =
        ensureSystemLossCanvas();


    if (!canvas) {
        return;
    }


    const records =
        readSystemLossMWh();


    destroyChart(
        "systemLossMwhChart"
    );


    if (!records.length) {
        return;
    }


    charts.systemLossMwhChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels:
                        records.map(
                            record =>
                                displayDate(
                                    record.date
                                )
                        ),

                    datasets: [

                        {

                            label:
                                "System Loss (MWh)",

                            data:
                                records.map(
                                    record =>
                                        record.loss
                                ),

                            borderWidth:
                                1,

                            borderRadius:
                                3

                        }

                    ]

                },


                options: {

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,


                    plugins: {

                        legend: {

                            display:
                                false

                        }

                    },


                    scales: {

                        x: {

                            title: {

                                display:
                                    true,

                                text:
                                    "Date"

                            },

                            ticks: {

                                autoSkip:
                                    true,

                                maxTicksLimit:
                                    15

                            },

                            grid: {

                                display:
                                    false

                            }

                        },


                        y: {

                            beginAtZero:
                                true,

                            title: {

                                display:
                                    true,

                                text:
                                    "System Loss (MWh)"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   CURTAILMENT
========================================================= */

function readCurtailment() {

    const rows =
        sheetToMatrix(
            getSheet(
                "Curtailment records"
            )
        );


    const result = [];


    rows.forEach(
        row => {

            const date =
                parseDate(
                    getCell(
                        row,
                        "C"
                    )
                );


            const start =
                timeToMinutes(
                    getCell(
                        row,
                        "H"
                    )
                );


            const end =
                timeToMinutes(
                    getCell(
                        row,
                        "I"
                    )
                );


            const loss =
                parseNumber(
                    getCell(
                        row,
                        "R"
                    )
                );


            if (
                !date ||
                start === null ||
                end === null
            ) {

                return;

            }


            /*
               Keep zero-loss intervals for the
               duration Gantt, but only use the
               actual loss value for aggregation.
            */

            let finish =
                end;


            if (
                finish < start
            ) {

                finish +=
                    24 * 60;

            }


            result.push({

                date,

                key:
                    dateKey(
                        date
                    ),

                start,

                end:
                    finish,

                loss:
                    loss ?? 0

            });

        }
    );


    result.sort(
        (a, b) =>
            a.date -
            b.date ||
            a.start -
            b.start
    );


    return result;

}


/* =========================================================
   CURTAILMENT DAILY TABLE
========================================================= */

function aggregateCurtailment(
    records
) {

    const map =
        new Map();


    records.forEach(
        record => {

            if (
                !map.has(
                    record.key
                )
            ) {

                map.set(
                    record.key,
                    {

                        date:
                            record.date,

                        loss:
                            0,

                        intervals:
                            0

                    }
                );

            }


            const item =
                map.get(
                    record.key
                );


            item.loss +=
                record.loss;


            item.intervals++;

        }
    );


    return Array.from(
        map.values()
    )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );

}


/* =========================================================
   CREATE CURTAILMENT TABLE
========================================================= */

function ensureCurtailmentTable() {

    const section =
        $("curtailmentSection");


    if (!section) {
        return null;
    }


    let card =
        $("curtailmentTableCard");


    if (!card) {

        card =
            document.createElement(
                "div"
            );


        card.id =
            "curtailmentTableCard";


        card.className =
            "chart-card full-card";


        card.style.marginTop =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        Daily Curtailment Loss
                    </h3>

                    <span>
                        Loss of generation merged by date · Column R
                    </span>

                </div>

                <span class="chart-type">
                    TABLE
                </span>

            </div>

            <div id="curtailmentTable"></div>

        `;


        const main =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (main) {

            main.insertAdjacentElement(
                "afterend",
                card
            );

        }

        else {

            section.appendChild(
                card
            );

        }

    }


    return $("curtailmentTable");

}


/* =========================================================
   RENDER CURTAILMENT TABLE
========================================================= */

function renderCurtailmentTable(
    records
) {

    const target =
        ensureCurtailmentTable();


    if (!target) {
        return;
    }


    const daily =
        aggregateCurtailment(
            records
        );


    if (!daily.length) {

        target.innerHTML = `

            <div style="
                padding:16px;
                color:#879397;
                font-size:10px;
            ">
                No curtailment records found.
            </div>

        `;

        return;

    }


    let html = `

        <table style="
            width:100%;
            border-collapse:collapse;
            font-size:10px;
        ">

            <thead>

                <tr>

                    <th style="
                        padding:9px;
                        text-align:left;
                        border-bottom:1px solid #e1ebed;
                    ">
                        Date
                    </th>

                    <th style="
                        padding:9px;
                        text-align:right;
                        border-bottom:1px solid #e1ebed;
                    ">
                        Loss of Generation (MWh)
                    </th>

                    <th style="
                        padding:9px;
                        text-align:right;
                        border-bottom:1px solid #e1ebed;
                    ">
                        Intervals
                    </th>

                </tr>

            </thead>

            <tbody>

    `;


    daily.forEach(
        record => {

            html += `

                <tr>

                    <td style="
                        padding:9px;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${formatFullDate(
                            record.date
                        )}
                    </td>

                    <td style="
                        padding:9px;
                        text-align:right;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${record.loss.toFixed(2)}
                    </td>

                    <td style="
                        padding:9px;
                        text-align:right;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${record.intervals}
                    </td>

                </tr>

            `;

        }
    );


    html += `

            </tbody>

        </table>

    `;


    target.innerHTML =
        html;


    const total =
        daily.reduce(
            (sum, item) =>
                sum +
                item.loss,
            0
        );


    if ($("curtailmentSummary")) {

        $("curtailmentSummary").textContent =
            `${daily.length} date(s) · ${total.toFixed(2)} MWh total loss`;

    }

}


/* =========================================================
   CURTAILMENT GANTT CANVAS
========================================================= */

function ensureCurtailmentGantt() {

    const section =
        $("curtailmentSection");


    if (!section) {
        return null;
    }


    let card =
        $("curtailmentGanttCard");


    if (!card) {

        card =
            document.createElement(
                "div"
            );


        card.id =
            "curtailmentGanttCard";


        card.className =
            "chart-card full-card";


        card.style.marginTop =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        Curtailment Duration
                    </h3>

                    <span>
                        06:00–18:00 · 15-minute intervals
                    </span>

                </div>

                <span class="chart-type">
                    GANTT
                </span>

            </div>

            <div class="timeline-wrapper">

                <canvas id="curtailmentGanttChart"></canvas>

            </div>

        `;


        const tableCard =
            $("curtailmentTableCard");


        if (tableCard) {

            tableCard.insertAdjacentElement(
                "afterend",
                card
            );

        }

        else {

            section.appendChild(
                card
            );

        }

    }


    return $("curtailmentGanttChart");

}


/* =========================================================
   CURTAILMENT GANTT
========================================================= */

function renderCurtailmentGantt(
    records
) {

    const canvas =
        ensureCurtailmentGantt();


    if (!canvas) {
        return;
    }


    destroyChart(
        "curtailmentGanttChart"
    );


    if (!records.length) {
        return;
    }


    const daily =
        aggregateCurtailment(
            records
        );


    const labels =
        daily.map(
            item =>
                formatShortDate(
                    item.date
                )
        );


    /*
       Each individual curtailment interval
       is a separate floating bar.

       The Y-axis is the date.
       X-axis = minutes from midnight.

       06:00 = 360
       18:00 = 1080
    */

    const datasets =
        records
            .filter(
                record =>
                    record.end > 360 &&
                    record.start < 1080
            )
            .map(
                record => {

                    const start =
                        Math.max(
                            360,
                            record.start
                        );


                    const end =
                        Math.min(
                            1080,
                            record.end
                        );


                    return {

                        label:
                            `${formatShortDate(record.date)} ${minutesToTime(start)}-${minutesToTime(end)}`,

                        data: [

                            {

                                x: [
                                    start,
                                    end
                                ],

                                y:
                                    formatShortDate(
                                        record.date
                                    )

                            }

                        ],

                        backgroundColor:
                            "rgba(39,165,173,0.72)",

                        borderColor:
                            "#27A5AD",

                        borderWidth:
                            1,

                        borderRadius:
                            4,

                        barThickness:
                            20

                    };

                }
            );


    charts.curtailmentGanttChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels,

                    datasets

                },


                options: {

                    indexAxis:
                        "y",

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,

                    parsing:
                        false,


                    plugins: {

                        legend: {

                            display:
                                false

                        },


                        tooltip: {

                            callbacks: {

                                title:
                                    context => {

                                        return (
                                            context[0]
                                                ?.raw
                                                ?.y ||
                                            ""
                                        );

                                    },


                                label:
                                    context => {

                                        const raw =
                                            context.raw;


                                        if (!raw) {
                                            return "";
                                        }


                                        return (
                                            `Time: ${minutesToTime(raw.x[0])} – ${minutesToTime(raw.x[1])}`
                                        );

                                    }

                            }

                        }

                    },


                    scales: {

                        x: {

                            type:
                                "linear",

                            min:
                                360,

                            max:
                                1080,

                            title: {

                                display:
                                    true,

                                text:
                                    "Time"

                            },

                            ticks: {

                                stepSize:
                                    15,

                                callback:
                                    value =>
                                        minutesToTime(
                                            value
                                        )

                            },

                            grid: {

                                color:
                                    "rgba(23,37,42,0.07)"

                            }

                        },


                        y: {

                            type:
                                "category",

                            labels,

                            grid: {

                                display:
                                    false

                            }

                        }

                    }

                }

            }
        );


}


/* =========================================================
   ANNUAL KPI ENERGY
========================================================= */

function readAnnualEnergy() {

    const sheet =
        getSheet(
            "Annual_KPI"
        );


    if (!sheet) {
        return [];
    }


    const rows =
        sheetToMatrix(
            sheet
        );


    const result = [];


    const months = [
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
        "Jan",
        "Feb",
        "Mar"
    ];


    /*
       Exact requested rows:

       E10:E21
       F10:F21
    */

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
                getCell(
                    row,
                    "E"
                )
            );


        const measured =
            parseNumber(
                getCell(
                    row,
                    "F"
                )
            );


        if (
            budget === null &&
            measured === null
        ) {

            continue;

        }


        result.push({

            month:
                months[
                    result.length
                ] ||
                `Month ${result.length + 1}`,

            budget,

            measured

        });

    }


    return result;

}


/* =========================================================
   ENERGY CHART
========================================================= */

function renderEnergyChart() {

    const canvas =
        $("energyChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "energyChart"
    );


    const records =
        readAnnualEnergy();


    if (!records.length) {
        return;
    }


    const budget =
        records.map(
            record =>
                record.budget
        );


    const measured =
        records.map(
            record =>
                record.measured
        );


    const totalBudget =
        budget.reduce(
            (sum, value) =>
                sum +
                (value || 0),
            0
        );


    const totalMeasured =
        measured.reduce(
            (sum, value) =>
                sum +
                (value || 0),
            0
        );


    const variance =
        totalMeasured -
        totalBudget;


    if ($("totalBudget")) {

        $("totalBudget").textContent =
            formatNumber(
                totalBudget,
                2
            );

    }


    if ($("totalMeasured")) {

        $("totalMeasured").textContent =
            formatNumber(
                totalMeasured,
                2
            );

    }


    if ($("energyVariance")) {

        $("energyVariance").textContent =
            formatNumber(
                variance,
                2
            );

    }


    charts.energyChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels:
                        records.map(
                            record =>
                                record.month
                        ),

                    datasets: [

                        {

                            label:
                                "Budgeted Energy",

                            data:
                                budget,

                            borderWidth:
                                1,

                            borderRadius:
                                4

                        },

                        {

                            label:
                                "Measured Energy",

                            data:
                                measured,

                            borderWidth:
                                1,

                            borderRadius:
                                4

                        }

                    ]

                },


                options: {

                    responsive:
                        true,

                    maintainAspectRatio:
                        false,

                    animation:
                        false,


                    interaction: {

                        mode:
                            "index",

                        intersect:
                            false

                    },


                    plugins: {

                        legend: {

                            display:
                                true,

                            position:
                                "top"

                        }

                    },


                    scales: {

                        x: {

                            grid: {

                                display:
                                    false

                            }

                        },


                        y: {

                            beginAtZero:
                                true,

                            ticks: {

                                maxTicksLimit:
                                    7

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   FORMAT NUMBER
========================================================= */

function formatNumber(
    value,
    decimals = 2
) {

    if (
        value === null ||
        value === undefined ||
        !Number.isFinite(
            Number(value)
        )
    ) {

        return "—";

    }


    return Number(
        value
    ).toLocaleString(
        "en-IN",
        {
            minimumFractionDigits:
                decimals,

            maximumFractionDigits:
                decimals
        }
    );

}


/* =========================================================
   FORMAT FULL DATE
========================================================= */

function formatFullDate(
    date
) {

    if (
        !(date instanceof Date) ||
        isNaN(
            date.getTime()
        )
    ) {

        return "—";

    }


    return date.toLocaleDateString(
        "en-IN",
        {

            day:
                "2-digit",

            month:
                "short",

            year:
                "numeric"

        }
    );

}


/* =========================================================
   SHOW CANVAS MESSAGE
========================================================= */

function showCanvasMessage(
    canvas,
    message
) {

    if (!canvas) return;


    const ctx =
        canvas.getContext(
            "2d"
        );


    if (!ctx) return;


    ctx.clearRect(
        0,
        0,
        canvas.width,
        canvas.height
    );


    ctx.save();


    ctx.textAlign =
        "center";


    ctx.textBaseline =
        "middle";


    ctx.fillStyle =
        "#879397";


    ctx.font =
        "12px Inter, Arial";


    ctx.fillText(
        message,
        canvas.width / 2,
        canvas.height / 2
    );


    ctx.restore();

}


/* =========================================================
   RENDER ALL
========================================================= */

function renderAll() {

    destroyAllCharts();


    /*
       Daily KPI
    */

    const daily =
        readDailyKPI();


    if (daily.length) {

        renderDailyKPICharts(
            daily
        );

    }


    /*
       PA
    */

    renderPlantUnavailability();

    renderBreakdownTimeline();

    renderSystemLossMWh();


    /*
       Curtailment
    */

    const curtailment =
        readCurtailment();


    renderCurtailmentTable(
        curtailment
    );


    renderCurtailmentGantt(
        curtailment
    );


    /*
       Energy
    */

    renderEnergyChart();


    /*
       Dashboard PA card should come from
       Daily_KPI PA value, not an artificial
       calculation from downtime.
    */

    if (
        daily.length &&
        $("dashboardPA")
    ) {

        const latest =
            daily[
                daily.length - 1
            ];


        $("dashboardPA").textContent =
            latest.pa === null
                ? "—"
                : latest.pa.toFixed(
                    2
                ) +
                  "%";

    }

}


/* =========================================================
   RESET
========================================================= */

function resetDashboard() {

    workbook = null;


    destroyAllCharts();


    if (fileInput) {
        fileInput.value = "";
    }


    if ($("fileInfo")) {
        $("fileInfo").classList.add(
            "hidden"
        );
    }


    if ($("workbookStatus")) {
        $("workbookStatus").classList.add(
            "hidden"
        );
    }


    if ($("dropZone")) {
        $("dropZone").classList.remove(
            "hidden"
        );
    }


    if ($("emptyState")) {
        $("emptyState").classList.remove(
            "hidden"
        );
    }


    /*
       Remove dynamically created charts/tables.
    */

    [
        "breakdownTimelineCard",
        "systemLossMwhCard",
        "curtailmentTableCard",
        "curtailmentGanttCard"
    ]
        .forEach(
            id => {

                const element =
                    $(id);

                if (element) {
                    element.remove();
                }

            }
        );


    hideAnalytics();


    setStatus(
        "Upload a DGR to generate the analytics."
    );


    if ($("sidebarFileName")) {

        $("sidebarFileName").textContent =
            "No DGR uploaded";

    }


    [
        "dashboardPA",
        "dashboardPR",
        "dashboardLoss",
        "dashboardHours",
        "totalBudget",
        "totalMeasured",
        "energyVariance"
    ]
        .forEach(
            id => {

                if ($(id)) {
                    $(id).textContent =
                        "—";
                }

            }
        );

}
