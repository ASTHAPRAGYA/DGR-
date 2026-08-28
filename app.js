/* =========================================================
   SOLAR DGR ANALYTICS
   app.js
   =========================================================

   EXACT WORKBOOK MAPPING

   DAILY_KPI
   ---------------------------------------------------------
   B  = Date
   I  = Operating Hours
   S  = PA (%)
   V  = PR (%)
   AD = System Losses (%)


   PA
   ---------------------------------------------------------
   B  = Date
   W  = Issue / Fault
   Z  = Fault Start Time
   AC = Work Completion Time on Fault
   AG = Breakdown Time (minutes)
   AL = System Loss (MWh)


   CURTAILMENT RECORDS
   ---------------------------------------------------------
   C  = Date
   H  = Start Time
   I  = End Time
   R  = Loss of Generation MWh


   ANNUAL_KPI
   ---------------------------------------------------------
   E10:E21 = Budgeted Energy
   F10:F21 = Measured Energy


   ========================================================= */

"use strict";


/* =========================================================
   GLOBAL
========================================================= */

let workbook = null;

const charts = {};


/* =========================================================
   DOM SHORTCUT
========================================================= */

function $(id) {
    return document.getElementById(id);
}


/* =========================================================
   INITIALISE
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        installScrollStyles();

        setupNavigation();

        setupUpload();

        setupRemoveButton();

        hideAnalytics();

    }
);


/* =========================================================
   INJECT SCROLL STYLES
   IMPORTANT:
   Existing CSS contains:
   canvas {
       width: 100% !important;
   }

   That prevents the canvas from becoming wider.

   We therefore inject stronger rules here.
========================================================= */

function installScrollStyles() {

    if (
        document.getElementById(
            "dgrScrollStyles"
        )
    ) {

        return;

    }


    const style =
        document.createElement(
            "style"
        );


    style.id =
        "dgrScrollStyles";


    style.textContent = `

        .scroll-chart-container {
            width: 100% !important;
            height: 100% !important;
            overflow-x: auto !important;
            overflow-y: hidden !important;
            position: relative !important;
            box-sizing: border-box !important;
            scrollbar-width: thin !important;
            scrollbar-color: #b9cacc transparent !important;
        }

        .scroll-chart-container::-webkit-scrollbar {
            height: 8px;
        }

        .scroll-chart-container::-webkit-scrollbar-track {
            background: #eef3f4;
            border-radius: 10px;
        }

        .scroll-chart-container::-webkit-scrollbar-thumb {
            background: #b9cacc;
            border-radius: 10px;
        }

        .scroll-chart-container::-webkit-scrollbar-thumb:hover {
            background: #27a5ad;
        }

        .scroll-chart-container > canvas {
            display: block !important;
            width: auto !important;
            max-width: none !important;
            min-width: 0 !important;
            flex: none !important;
        }

        #paChart,
        #breakdownChart,
        #systemLossMwhChart,
        #curtailmentGanttChart,
        #paPercentageChart,
        #prChart,
        #hoursChart,
        #lossChart,
        #dashboardPRChart,
        #dashboardLossChart {
            max-width: none !important;
        }

    `;


    document.head.appendChild(
        style
    );

}


/* =========================================================
   NAVIGATION
========================================================= */

function setupNavigation() {

    const items =
        document.querySelectorAll(
            ".nav-item"
        );


    items.forEach(
        item => {

            item.addEventListener(
                "click",
                () => {

                    items.forEach(
                        nav =>
                            nav.classList.remove(
                                "active"
                            )
                    );


                    item.classList.add(
                        "active"
                    );


                    const target =
                        $(item.dataset.target);


                    if (target) {

                        target.scrollIntoView(
                            {
                                behavior:
                                    "smooth",

                                block:
                                    "start"
                            }
                        );

                    }

                }
            );

        }
    );

}


/* =========================================================
   UPLOAD
========================================================= */

function setupUpload() {

    const input =
        $("dgrFile");


    const dropZone =
        $("dropZone");


    if (!input) {

        console.error(
            "dgrFile input not found."
        );

        return;

    }


    input.addEventListener(
        "change",
        event => {

            const file =
                event.target.files &&
                event.target.files[0];


            if (file) {

                processFile(
                    file
                );

            }

        }
    );


    /*
       The old drop zone is optional.
       If it exists, it remains clickable.
    */

    if (dropZone) {

        dropZone.addEventListener(
            "click",
            () => {

                input.click();

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
                    event
                        .dataTransfer
                        ?.files
                        ?. [0];


                if (file) {

                    processFile(
                        file
                    );

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


    if (!button) {

        return;

    }


    button.addEventListener(
        "click",
        resetDashboard
    );

}


/* =========================================================
   FILE PROCESSING
========================================================= */

function processFile(
    file
) {

    if (!file) {
        return;
    }


    if (
        !/\.(xlsx|xls|csv)$/i.test(
            file.name
        )
    ) {

        alert(
            "Please upload an Excel file (.xlsx/.xls) or CSV file."
        );

        return;

    }


    if (
        typeof XLSX === "undefined"
    ) {

        alert(
            "SheetJS is not loaded. Please check index.html."
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

                workbook =
                    XLSX.read(
                        new Uint8Array(
                            event.target.result
                        ),
                        {
                            type:
                                "array",

                            cellDates:
                                true,

                            cellNF:
                                true
                        }
                    );


                if (
                    !workbook.SheetNames ||
                    !workbook.SheetNames.length
                ) {

                    throw new Error(
                        "No worksheets were found in the workbook."
                    );

                }


                updateFileInformation(
                    file
                );


                showAnalytics();


                renderAll();

            }

            catch (error) {

                console.error(
                    error
                );


                setStatus(
                    "Unable to read the DGR."
                );


                alert(
                    "Unable to read the uploaded DGR.\n\n" +
                    error.message
                );

            }

        };


    reader.onerror =
        () => {

            setStatus(
                "Unable to read the selected file."
            );

        };


    reader.readAsArrayBuffer(
        file
    );

}


/* =========================================================
   FILE INFORMATION
========================================================= */

function updateFileInformation(
    file
) {

    setText(
        "fileName",
        file.name
    );


    setText(
        "fileSheets",
        `${workbook.SheetNames.length} worksheets detected`
    );


    setText(
        "sidebarFileName",
        file.name
    );


    $("fileInfo")
        ?.classList
        .remove(
            "hidden"
        );


    $("workbookStatus")
        ?.classList
        .remove(
            "hidden"
        );


    $("emptyState")
        ?.classList
        .add(
            "hidden"
        );


    $("dropZone")
        ?.classList
        .add(
            "hidden"
        );


    renderSheetBadges();

}


/* =========================================================
   SHEET BADGES
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


    container.innerHTML =
        "";


    const sheets = [
        "Dashboard",
        "Annual_KPI",
        "Daily_KPI",
        "PA",
        "Curtailment records"
    ];


    sheets.forEach(
        name => {

            const badge =
                document.createElement(
                    "span"
                );


            badge.className =
                "sheet-badge";


            const found =
                workbook.SheetNames.some(
                    actual =>
                        normalizeSheet(
                            actual
                        ) ===
                        normalizeSheet(
                            name
                        )
                );


            badge.textContent =
                found
                    ? `${name} ✓`
                    : `${name} — missing`;


            if (!found) {

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
   NORMALIZE SHEET NAME
========================================================= */

function normalizeSheet(
    name
) {

    return String(
        name || ""
    )
        .trim()
        .toLowerCase()
        .replace(
            /[\s_-]+/g,
            ""
        );

}


/* =========================================================
   GET SHEET
========================================================= */

function getSheet(
    requestedName
) {

    if (
        !workbook ||
        !workbook.Sheets
    ) {

        return null;

    }


    if (
        workbook.Sheets[
            requestedName
        ]
    ) {

        return workbook.Sheets[
            requestedName
        ];

    }


    const normalized =
        normalizeSheet(
            requestedName
        );


    const actual =
        workbook.SheetNames.find(
            name =>
                normalizeSheet(
                    name
                ) === normalized
        );


    return actual
        ? workbook.Sheets[
            actual
        ]
        : null;

}


/* =========================================================
   SHEET MATRIX
========================================================= */

function sheetMatrix(
    sheet
) {

    if (!sheet) {
        return [];
    }


    return XLSX.utils.sheet_to_json(
        sheet,
        {
            header:
                1,

            raw:
                true,

            defval:
                null,

            blankrows:
                false
        }
    );

}


/* =========================================================
   EXCEL COLUMN INDEX
========================================================= */

function colIndex(
    column
) {

    let value =
        0;


    for (
        const char of column.toUpperCase()
    ) {

        value =
            value * 26 +
            char.charCodeAt(0) -
            64;

    }


    return value - 1;

}


/* =========================================================
   CELL
========================================================= */

function getCell(
    row,
    column
) {

    if (!row) {

        return null;

    }


    return row[
        colIndex(
            column
        )
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
        String(
            value
        )
            .trim()
            .replace(
                /,/g,
                ""
            )
            .replace(
                /%/g,
                ""
            );


    if (
        !text ||
        text.startsWith("#")
    ) {

        return null;

    }


    const result =
        Number(
            text
        );


    return Number.isFinite(
        result
    )
        ? result
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


    if (
        value instanceof Date
    ) {

        return isNaN(
            value.getTime()
        )
            ? null
            : new Date(
                value.getTime()
            );

    }


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

        catch (_) {}

        return null;

    }


    const text =
        String(
            value
        )
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

        const d =
            Number(
                match[1]
            );


        const m =
            Number(
                match[2]
            ) - 1;


        const y =
            Number(
                match[3]
            );


        const date =
            new Date(
                y,
                m,
                d
            );


        if (
            date.getFullYear() === y &&
            date.getMonth() === m &&
            date.getDate() === d
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
                    .substring(
                        0,
                        3
                    )
                    .toLowerCase()
            );


        let year =
            Number(
                match[3]
            );


        if (
            year < 100
        ) {

            year +=
                2000;

        }


        if (
            month >= 0
        ) {

            return new Date(
                year,
                month,
                Number(
                    match[1]
                )
            );

        }

    }


    const browserDate =
        new Date(
            text
        );


    return isNaN(
        browserDate.getTime()
    )
        ? null
        : browserDate;

}


/* =========================================================
   TIME PARSER
========================================================= */

function parseTimeMinutes(
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
       Date object
    */

    if (
        value instanceof Date
    ) {

        if (
            isNaN(
                value.getTime()
            )
        ) {

            return null;

        }


        return (
            value.getHours() * 60 +
            value.getMinutes()
        );

    }


    /*
       Excel fraction
    */

    if (
        typeof value === "number"
    ) {

        if (
            value >= 0 &&
            value < 1
        ) {

            return Math.round(
                value * 1440
            );

        }


        try {

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

        }

        catch (_) {}

    }


    const text =
        String(
            value
        )
            .trim();


    const match =
        text.match(
            /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i
        );


    if (
        !match
    ) {

        return null;

    }


    let hours =
        Number(
            match[1]
        );


    const minutes =
        Number(
            match[2]
        );


    const ampm =
        match[4];


    if (
        ampm
    ) {

        if (
            ampm.toUpperCase() ===
            "PM" &&
            hours < 12
        ) {

            hours +=
                12;

        }


        if (
            ampm.toUpperCase() ===
            "AM" &&
            hours === 12
        ) {

            hours =
                0;

        }

    }


    if (
        hours < 0 ||
        hours > 23 ||
        minutes < 0 ||
        minutes > 59
    ) {

        return null;

    }


    return (
        hours * 60 +
        minutes
    );

}


/* =========================================================
   MINUTES -> TIME
========================================================= */

function minutesToTime(
    minutes
) {

    let value =
        Math.round(
            minutes
        );


    value =
        Math.max(
            0,
            Math.min(
                1439,
                value
            )
        );


    const hours =
        Math.floor(
            value / 60
        );


    const mins =
        value % 60;


    return (
        String(
            hours
        ).padStart(
            2,
            "0"
        ) +
        ":" +
        String(
            mins
        ).padStart(
            2,
            "0"
        )
    );

}


/* =========================================================
   FORMAT DATE
========================================================= */

function formatDate(
    value
) {

    if (
        !(value instanceof Date)
    ) {

        return "";

    }


    return value.toLocaleDateString(
        "en-IN",
        {
            day:
                "2-digit",

            month:
                "short"
        }
    );

}


/* =========================================================
   FULL DATE
========================================================= */

function formatFullDate(
    value
) {

    if (
        !(value instanceof Date)
    ) {

        return "";

    }


    return value.toLocaleDateString(
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
   NUMBER FORMAT
========================================================= */

function formatNumber(
    value,
    decimals = 2
) {

    if (
        value === null ||
        !Number.isFinite(
            value
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
   SET TEXT
========================================================= */

function setText(
    id,
    value
) {

    const element =
        $(id);


    if (element) {

        element.textContent =
            value;

    }

}


/* =========================================================
   STATUS
========================================================= */

function setStatus(
    message
) {

    setText(
        "statusText",
        message
    );

}

/* =========================================================
   DESTROY SINGLE CHART
========================================================= */

function destroyChart(id) {

    if (
        charts[id] &&
        typeof charts[id].destroy === "function"
    ) {

        try {

            charts[id].destroy();

        } catch (error) {

            console.warn(
                `Could not destroy chart: ${id}`,
                error
            );

        }

    }

    charts[id] = null;

}
/* =========================================================
   SHOW / HIDE
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

                const element =
                    $(id);


                if (element) {

                    element.style.display =
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

                const element =
                    $(id);


                if (element) {

                    element.style.display =
                        "block";

                }

            }
        );

}


/* =========================================================
   SCROLLABLE CANVAS
========================================================= */

function makeScrollable(
    canvas,
    desiredWidth
) {

    if (!canvas) {

        return null;

    }


    const parent =
        canvas.parentElement;


    if (!parent) {

        return null;

    }


    let wrapper =
        parent.querySelector(
            ".scroll-chart-container"
        );


    /*
       If the canvas was already wrapped,
       move on without creating another wrapper.
    */

    if (!wrapper) {

        wrapper =
            document.createElement(
                "div"
            );


        wrapper.className =
            "scroll-chart-container";


        parent.insertBefore(
            wrapper,
            canvas
        );


        wrapper.appendChild(
            canvas
        );

    }


    const visibleWidth =
        parent.clientWidth ||
        600;


    const finalWidth =
        Math.max(
            visibleWidth,
            desiredWidth
        );


    wrapper.style.width =
        "100%";


    wrapper.style.height =
        "100%";


    wrapper.style.overflowX =
        "auto";


    wrapper.style.overflowY =
        "hidden";


    canvas.classList.add(
        "scroll-chart-canvas"
    );


    /*
       CSS !important is intentionally set
       here because the old stylesheet
       contains width:100%!important.
    */

    canvas.style.setProperty(
        "width",
        `${finalWidth}px`,
        "important"
    );


    canvas.style.setProperty(
        "max-width",
        "none",
        "important"
    );


    canvas.style.setProperty(
        "min-width",
        `${finalWidth}px`,
        "important"
    );


    canvas.style.height =
        "100%";


    return wrapper;

}


/* =========================================================
   GENERIC SCROLLABLE LINE CHART
========================================================= */

function createScrollableLineChart(
    canvasId,
    labels,
    values,
    label,
    yTitle,
    options = {}
) {

    const canvas =
        $(canvasId);


    if (!canvas) {

        return;

    }


    destroyChart(
        canvasId
    );


    const parent =
        canvas.parentElement;


    const visibleWidth =
        parent?.clientWidth ||
        700;


    const desiredWidth =
        Math.max(
            visibleWidth,
            labels.length *
            78
        );


    makeScrollable(
        canvas,
        desiredWidth
    );


    const yValues =
        values.filter(
            value =>
                Number.isFinite(
                    value
                )
        );


    let yMin =
        options.yMin;


    let yMax =
        options.yMax;


    if (
        yMin === undefined &&
        yValues.length
    ) {

        yMin =
            Math.min(
                ...yValues
            );


        if (
            options.beginAtZero
        ) {

            yMin =
                0;

        }

    }


    if (
        yMax === undefined &&
        yValues.length
    ) {

        yMax =
            Math.max(
                ...yValues
            );

    }


    charts[canvasId] =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "line",

                data: {

                    labels,

                    datasets: [

                        {

                            label,

                            data:
                                values,

                            borderWidth:
                                2,

                            pointRadius:
                                3,

                            pointHoverRadius:
                                6,

                            tension:
                                0.22,

                            fill:
                                false

                        }

                    ]

                },


                options: {

                    responsive:
                        false,

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
                                    false,

                                maxRotation:
                                    0,

                                minRotation:
                                    0,

                                font: {

                                    size:
                                        9

                                }

                            }

                        },


                        y: {

                            min:
                                yMin,

                            max:
                                yMax,

                            beginAtZero:
                                options.beginAtZero ||
                                false,

                            title: {

                                display:
                                    true,

                                text:
                                    yTitle

                            },

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
   DAILY KPI
========================================================= */

function readDailyKPI() {

    const rows =
        sheetMatrix(
            getSheet(
                "Daily_KPI"
            )
        );


    const result = [];


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


            const hours =
                parseNumber(
                    getCell(
                        row,
                        "I"
                    )
                );


            const paRaw =
                parseNumber(
                    getCell(
                        row,
                        "S"
                    )
                );


            const prRaw =
                parseNumber(
                    getCell(
                        row,
                        "V"
                    )
                );


            const lossRaw =
                parseNumber(
                    getCell(
                        row,
                        "AD"
                    )
                );


            /*
               Convert decimal percentages
               to actual percentage values.

               0.84 -> 84
               0.02 -> 2
            */

            const pa =
                paRaw === null
                    ? null
                    : (
                        Math.abs(
                            paRaw
                        ) <= 1.5
                            ? paRaw * 100
                            : paRaw
                    );


            const pr =
                prRaw === null
                    ? null
                    : (
                        Math.abs(
                            prRaw
                        ) <= 1.5
                            ? prRaw * 100
                            : prRaw
                    );


            const loss =
                lossRaw === null
                    ? null
                    : (
                        Math.abs(
                            lossRaw
                        ) <= 1.5
                            ? lossRaw * 100
                            : lossRaw
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

                pa,

                pr,

                hours,

                loss

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
   DATE LABELS
========================================================= */

function makeDailyLabels(
    rows
) {

    const count =
        rows.length;


    return rows.map(
        row => {

            const day =
                row.date.getDate();


            /*
               One-month datasets:
               only even dates are displayed.

               Example:
               2 4 6 8 10 12...
            */

            if (
                count >= 20 &&
                count <= 35
            ) {

                return (
                    day % 2 === 0
                        ? String(
                            day
                        )
                        : ""
                );

            }


            return formatDate(
                row.date
            );

        }
    );

}


/* =========================================================
   RENDER DAILY KPI CHARTS
========================================================= */

function renderDailyKPICharts(
    rows
) {

    if (!rows.length) {
        return;
    }


    const labels =
        makeDailyLabels(
            rows
        );


    /*
       PA %
    */

    renderPAPercentage(
        rows
    );


    /*
       PR
    */

    createScrollableLineChart(
        "prChart",
        labels,
        rows.map(
            row =>
                row.pr
        ),
        "Performance Ratio",
        "PR (%)",
        {
            yMin:
                0,

            yMax:
                100
        }
    );


    /*
       Operating Hours
    */

    createScrollableLineChart(
        "hoursChart",
        labels,
        rows.map(
            row =>
                row.hours
        ),
        "Operating Hours",
        "Operating Hours",
        {
            beginAtZero:
                true
        }
    );


    /*
       System Loss %
    */

    createScrollableLineChart(
        "lossChart",
        labels,
        rows.map(
            row =>
                row.loss
        ),
        "System Loss",
        "System Loss (%)",
        {
            beginAtZero:
                true
        }
    );


    /*
       Dashboard PR
    */

    createScrollableLineChart(
        "dashboardPRChart",
        labels,
        rows.map(
            row =>
                row.pr
        ),
        "Performance Ratio",
        "PR (%)",
        {
            yMin:
                0,

            yMax:
                100
        }
    );


    /*
       Dashboard System Loss
    */

    createScrollableLineChart(
        "dashboardLossChart",
        labels,
        rows.map(
            row =>
                row.loss
        ),
        "System Loss",
        "System Loss (%)",
        {
            beginAtZero:
                true
        }
    );


    /*
       KPI cards
    */

    const latest =
        rows[
            rows.length - 1
        ];


    setText(
        "dashboardPA",
        latest.pa === null
            ? "—"
            : `${latest.pa.toFixed(2)}%`
    );


    setText(
        "dashboardPR",
        latest.pr === null
            ? "—"
            : `${latest.pr.toFixed(2)}%`
    );


    setText(
        "dashboardLoss",
        latest.loss === null
            ? "—"
            : `${latest.loss.toFixed(2)}%`
    );


    setText(
        "dashboardHours",
        latest.hours === null
            ? "—"
            : `${latest.hours.toFixed(2)} h`
    );

}


/* =========================================================
   PA PERCENTAGE
========================================================= */

function ensurePAPercentageCard() {

    const section =
        $("paSection");


    if (!section) {
        return null;
    }


    let card =
        $("paPercentageCard");


    if (!card) {

        card =
            document.createElement(
                "div"
            );


        card.id =
            "paPercentageCard";


        card.className =
            "chart-card full-card";


        card.style.marginTop =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        Plant Availability
                    </h3>

                    <span>
                        Daily PA (%) from Daily_KPI
                    </span>

                </div>

                <span class="chart-type">
                    PA %
                </span>

            </div>

            <div class="chart-container">

                <canvas id="paPercentageChart"></canvas>

            </div>

        `;


        const mainCard =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (mainCard) {

            section.insertBefore(
                card,
                mainCard
            );

        }

        else {

            section.appendChild(
                card
            );

        }

    }


    return card;

}


/* =========================================================
   PA % CHART
========================================================= */

function renderPAPercentage(
    rows
) {

    const card =
        ensurePAPercentageCard();


    if (!card) {
        return;
    }


    const canvas =
        $("paPercentageChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "paPercentageChart"
    );


    const labels =
        makeDailyLabels(
            rows
        );


    const width =
        Math.max(
            card.clientWidth || 700,
            rows.length * 78
        );


    makeScrollable(
        canvas,
        width
    );


    chartRegistryFix(
        "paPercentageChart",
        canvas
    );


    charts.paPercentageChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "line",

                data: {

                    labels,

                    datasets: [

                        {

                            label:
                                "PA (%)",

                            data:
                                rows.map(
                                    row =>
                                        row.pa
                                ),

                            borderWidth:
                                2,

                            pointRadius:
                                3,

                            pointHoverRadius:
                                6,

                            tension:
                                0.20,

                            fill:
                                false

                        }

                    ]

                },


                options: {

                    responsive:
                        false,

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

                            grid: {

                                display:
                                    false

                            },

                            ticks: {

                                autoSkip:
                                    false,

                                maxRotation:
                                    0,

                                minRotation:
                                    0

                            }

                        },


                        y: {

                            min:
                                80,

                            max:
                                100,

                            title: {

                                display:
                                    true,

                                text:
                                    "Plant Availability (%)"

                            },

                            ticks: {

                                stepSize:
                                    5,

                                callback:
                                    value =>
                                        `${value}%`

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   CHART REGISTRY FIX
========================================================= */

function chartRegistryFix(
    id,
    canvas
) {

    if (
        !canvas
    ) {

        return;

    }


    /*
       Set a real canvas drawing size while
       preserving the browser display size.
    */

    const parent =
        canvas.parentElement;


    if (
        !parent
    ) {

        return;

    }


    canvas.style.height =
        "100%";

}


/* =========================================================
   PLANT UNAVAILABILITY
========================================================= */

function readPlantUnavailability() {

    const rows =
        sheetMatrix(
            getSheet(
                "PA"
            )
        );


    const records = [];


    rows.forEach(
        row => {

            const issue =
                cell(
                    row,
                    "W"
                );


            const start =
                parseTimeMinutes(
                    cell(
                        row,
                        "Z"
                    )
                );


            const end =
                parseTimeMinutes(
                    cell(
                        row,
                        "AC"
                    )
                );


            if (
                issue === null ||
                issue === undefined ||
                String(
                    issue
                ).trim() === ""
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


            /*
               If end crosses midnight.
            */

            if (
                finish < start
            ) {

                finish +=
                    1440;

            }


            records.push({

                issue:
                    String(
                        issue
                    ).trim(),

                start,

                end:
                    finish

            });

        }
    );


    return records;

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


    const parent =
        canvas.parentElement;


    const width =
        Math.max(
            parent?.clientWidth || 700,
            2304
        );


    makeScrollable(
        canvas,
        width
    );


    const labels =
        records.map(
            record =>
                record.issue
        );


    const datasets =
        records.map(
            record => ({

                label:
                    record.issue,

                data: [

                    {

                        x: [
                            record.start,
                            record.end
                        ],

                        y:
                            record.issue

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
                        false,

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
                                        ]?.issue ||
                                        "",


                                label:
                                    context => {

                                        const item =
                                            records[
                                                context
                                                    .dataIndex
                                            ];


                                        return [

                                            `Start: ${minutesToTime(item.start)}`,

                                            `End: ${minutesToTime(item.end)}`,

                                            `Duration: ${item.end - item.start} min`

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

                            ticks: {

                                stepSize:
                                    30,

                                callback:
                                    value =>
                                        minutesToTime(
                                            value
                                        )

                            },

                            grid: {

                                color:
                                    "rgba(23,37,42,0.08)"

                            }

                        },


                        y: {

                            type:
                                "category",

                            labels,

                            title: {

                                display:
                                    true,

                                text:
                                    "Issue / Fault"

                            },

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
   BREAKDOWN TIMELINE
========================================================= */

function readBreakdownData() {

    const rows =
        sheetMatrix(
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
                    cell(
                        row,
                        "B"
                    )
                );


            const breakdown =
                parseNumber(
                    cell(
                        row,
                        "AG"
                    )
                );


            if (
                !date ||
                breakdown === null
            ) {

                return;

            }


            const key =
                dayKey(
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

                        total:
                            0

                    }
                );

            }


            map.get(
                key
            ).total +=
                breakdown;

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
   DAY KEY
========================================================= */

function dayKey(
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

    ].join(
        "-"
    );

}


/* =========================================================
   BREAKDOWN CARD
========================================================= */

function ensureBreakdownCard() {

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
                        Same-date breakdown time combined from PA · Column AG
                    </span>

                </div>

                <span class="chart-type">
                    MINUTES
                </span>

            </div>

            <div class="chart-container">

                <canvas id="breakdownChart"></canvas>

            </div>

        `;


        const firstCard =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (firstCard) {

            firstCard.insertAdjacentElement(
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


    return card;

}


/* =========================================================
   BREAKDOWN CHART
========================================================= */

function renderBreakdownTimeline() {

    const card =
        ensureBreakdownCard();


    if (!card) {
        return;
    }


    const canvas =
        $("breakdownChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "breakdownChart"
    );


    const records =
        readBreakdownData();


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No breakdown timeline data found."
        );

        return;

    }


    /*
       Fixed requested X-axis:
       0 to 13 minutes
       interval 1 minute.
    */

    chartRegistry.breakdownChart =
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
                                formatDate(
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
                                        record.total
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
========================================================= */

function readSystemLossMWh() {

    const rows =
        sheetMatrix(
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
                    cell(
                        row,
                        "B"
                    )
                );


            const loss =
                parseNumber(
                    cell(
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
                dayKey(
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
   SYSTEM LOSS CARD
========================================================= */

function ensureSystemLossCard() {

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


        const breakdownCard =
            $("breakdownTimelineCard");


        if (breakdownCard) {

            breakdownCard.insertAdjacentElement(
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


    return card;

}


/* =========================================================
   SYSTEM LOSS CHART
========================================================= */

function renderSystemLossMWh() {

    const card =
        ensureSystemLossCard();


    if (!card) {
        return;
    }


    const canvas =
        $("systemLossMwhChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "systemLossMwhChart"
    );


    const records =
        readSystemLossMWh();


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No system loss MWh records found."
        );

        return;

    }


    const width =
        Math.max(
            card.clientWidth || 700,
            records.length * 80
        );


    makeScrollable(
        canvas,
        width
    );


    chartRegistry.systemLossMwhChart =
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
                                shortDate(
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
                        false,

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

                            grid: {

                                display:
                                    false

                            },

                            ticks: {

                                autoSkip:
                                    false,

                                maxTicksLimit:
                                    15

                            },

                            title: {

                                display:
                                    true,

                                text:
                                    "Date"

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
        sheetMatrix(
            getSheet(
                "Curtailment records"
            )
        );


    const records = [];


    rows.forEach(
        row => {

            const date =
                parseDate(
                    cell(
                        row,
                        "C"
                    )
                );


            const start =
                parseTimeMinutes(
                    cell(
                        row,
                        "H"
                    )
                );


            const end =
                parseTimeMinutes(
                    cell(
                        row,
                        "I"
                    )
                );


            const loss =
                parseNumber(
                    cell(
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


            let finish =
                end;


            if (
                finish < start
            ) {

                finish +=
                    1440;

            }


            records.push({

                date,

                key:
                    dayKey(
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


    records.sort(
        (a, b) =>
            a.date -
            b.date ||
            a.start -
            b.start
    );


    return records;

}


/* =========================================================
   CURTAILMENT TABLE
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
                        Loss of generation merged for each date · Column R
                    </span>

                </div>

                <span class="chart-type">
                    TABLE
                </span>

            </div>

            <div
                id="curtailmentTable"
                style="overflow-x:auto;"
            ></div>

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


            map.get(
                record.key
            ).loss +=
                record.loss;


            map.get(
                record.key
            ).intervals++;

        }
    );


    const daily =
        Array.from(
            map.values()
        )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );


    if (!daily.length) {

        target.innerHTML = `

            <div style="
                padding:15px;
                color:#879397;
                font-size:10px;
            ">
                No curtailment records found.
            </div>

        `;

        setText(
            "curtailmentSummary",
            "No curtailment records found"
        );

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
                        padding:10px;
                        text-align:left;
                        border-bottom:1px solid #e1ebed;
                    ">
                        Date
                    </th>

                    <th style="
                        padding:10px;
                        text-align:right;
                        border-bottom:1px solid #e1ebed;
                    ">
                        Loss of Generation (MWh)
                    </th>

                    <th style="
                        padding:10px;
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
        item => {

            html += `

                <tr>

                    <td style="
                        padding:10px;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${formatFullDate(
                            item.date
                        )}
                    </td>

                    <td style="
                        padding:10px;
                        text-align:right;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${item.loss.toFixed(2)}
                    </td>

                    <td style="
                        padding:10px;
                        text-align:right;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${item.intervals}
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


    setText(
        "curtailmentSummary",
        `${daily.length} date(s) · ${total.toFixed(2)} MWh total generation loss`
    );

}


/* =========================================================
   CURTAILMENT GANTT CARD
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


    return card;

}


/* =========================================================
   CURTAILMENT GANTT
========================================================= */

function renderCurtailmentGantt(
    records
) {

    const card =
        ensureCurtailmentGantt();


    if (!card) {
        return;
    }


    const canvas =
        $("curtailmentGanttChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "curtailmentGanttChart"
    );


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No curtailment intervals found."
        );

        return;

    }


    const width =
        Math.max(
            card.clientWidth || 700,
            3120
        );


    makeScrollable(
        canvas,
        width
    );


    /*
       Unique dates on Y axis.
    */

    const labels = [];


    records.forEach(
        record => {

            const label =
                shortDate(
                    record.date
                );


            if (
                !labels.includes(
                    label
                )
            ) {

                labels.push(
                    label
                );

            }

        }
    );


    /*
       Create one floating bar per
       curtailment interval.
    */

    const datasets =
        records
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


                    if (
                        end <= start
                    ) {

                        return null;

                    }


                    return {

                        label:
                            `${shortDate(record.date)} ${minutesToTime(start)}–${minutesToTime(end)}`,

                        data: [

                            {

                                x: [
                                    start,
                                    end
                                ],

                                y:
                                    shortDate(
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
            )
            .filter(
                Boolean
            );


    chartRegistry.curtailmentGanttChart =
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
                        false,

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
                                        context[0]
                                            ?.raw
                                            ?.y ||
                                        "",


                                label:
                                    context => {

                                        const raw =
                                            context.raw;


                                        if (!raw) {
                                            return "";
                                        }


                                        return (
                                            `Time: ${minutesToTime(
                                                raw.x[0]
                                            )} – ${minutesToTime(
                                                raw.x[1]
                                            )}`
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
                                    "rgba(23,37,42,0.08)"

                            }

                        },


                        y: {

                            type:
                                "category",

                            labels,

                            title: {

                                display:
                                    true,

                                text:
                                    "Date"

                            },

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
   ENERGY
========================================================= */

function renderEnergyChart() {

    const sheet =
        getSheet(
            "Annual_KPI"
        );


    if (!sheet) {
        return;
    }


    const rows =
        sheetMatrix(
            sheet
        );


    const labels = [
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


    const budget = [];
    const measured = [];


    /*
       Exact:
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
            ] ||
            [];


        budget.push(
            parseNumber(
                getCell(
                    row,
                    "E"
                )
            )
        );


        measured.push(
            parseNumber(
                getCell(
                    row,
                    "F"
                )
            )
        );

    }


    /*
       Summary
    */

    const totalBudget =
        budget.reduce(
            (sum, value) =>
                sum +
                (
                    value || 0
                ),
            0
        );


    const totalMeasured =
        measured.reduce(
            (sum, value) =>
                sum +
                (
                    value || 0
                ),
            0
        );


    const variance =
        totalMeasured -
        totalBudget;


    setText(
        "totalBudget",
        `${formatNumber(
            totalBudget
        )} MWh`
    );


    setText(
        "totalMeasured",
        `${formatNumber(
            totalMeasured
        )} MWh`
    );


    setText(
        "energyVariance",
        `${formatNumber(
            variance
        )} MWh`
    );


    const canvas =
        $("energyChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "energyChart"
    );


    charts.energyChart =
        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels,

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

                            title: {

                                display:
                                    true,

                                text:
                                    "Energy (MWh)"

                            }

                        }

                    }

                }

            }
        );

}


/* =========================================================
   CANVAS MESSAGE
========================================================= */

function showCanvasMessage(
    canvas,
    message
) {

    if (!canvas) {
        return;
    }


    const ctx =
        canvas.getContext(
            "2d"
        );


    if (!ctx) {
        return;
    }


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


    ctx.font =
        "12px Inter, Arial";


    ctx.fillStyle =
        "#879397";


    ctx.fillText(
        message,
        canvas.width / 2,
        canvas.height / 2
    );


    ctx.restore();

}


/* =========================================================
   RESET
========================================================= */

function resetDashboard() {

    workbook =
        null;


    destroyAllCharts();


    removeDynamicElements();


    if (fileInput) {

        fileInput.value =
            "";

    }


    $("fileInfo")
        ?.classList
        .add(
            "hidden"
        );


    $("workbookStatus")
        ?.classList
        .add(
            "hidden"
        );


    $("dropZone")
        ?.classList
        .remove(
            "hidden"
        );


    $("emptyState")
        ?.classList
        .remove(
            "hidden"
        );


    hideAnalytics();


    setText(
        "fileName",
        "—"
    );


    setText(
        "fileSheets",
        "—"
    );


    setText(
        "sidebarFileName",
        "No DGR uploaded"
    );


    setStatus(
        "Upload a DGR to generate the analytics."
    );


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
            id =>
                setText(
                    id,
                    "—"
                )
        );


    setText(
        "curtailmentSummary",
        "Waiting for DGR data"
    );

}


/* =========================================================
   REMOVE DYNAMIC CARDS
========================================================= */

function removeDynamicElements() {

    [
        "paPercentageCard",
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

}


/* =========================================================
   DESTROY CHARTS
========================================================= */

function destroyAllCharts() {

    Object.keys(
        charts
    )
        .forEach(
            id => {

                if (
                    charts[id] &&
                    typeof charts[id].destroy ===
                    "function"
                ) {

                    try {

                        charts[id].destroy();

                    }

                    catch (_) {}

                }


                charts[id] =
                    null;

            }
        );

}


/* =========================================================
   RENDER ALL
========================================================= */

function renderAll() {

    if (!workbook) {
        return;
    }


    destroyAllCharts();


    removeDynamicElements();


    /*
       DAILY KPI
    */

    const daily =
        readDailyKPI();


    if (
        daily.length
    ) {

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
       CURTAILMENT
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
       ENERGY
    */

    renderEnergyChart();


    /*
       Status
    */

    if (
        daily.length
    ) {

        setStatus(
            `DGR loaded successfully • ${daily.length} daily records analysed`
        );

    }

}


/* =========================================================
   END
========================================================= */
