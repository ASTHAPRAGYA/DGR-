/* =========================================================
   SOLAR DGR ANALYTICS
   app.js

   =========================================================
   EXACT WORKBOOK MAPPING
   =========================================================

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
   AC = Fault End / Work Completion Time
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


/* =========================================================
   GLOBAL
========================================================= */

"use strict";

let workbook = null;

const charts = {};


/* =========================================================
   DOM SHORTCUT
========================================================= */

function $(id) {
    return document.getElementById(id);
}


/* =========================================================
   PAGE LOAD
========================================================= */

document.addEventListener(
    "DOMContentLoaded",
    function () {

        setupNavigation();

        setupUpload();

        setupRemoveButton();

        hideAnalytics();

    }
);


/* =========================================================
   NAVIGATION
========================================================= */

function setupNavigation() {

    const buttons =
        document.querySelectorAll(".nav-item");


    buttons.forEach(
        function (button) {

            button.addEventListener(
                "click",
                function () {

                    buttons.forEach(
                        function (item) {

                            item.classList.remove(
                                "active"
                            );

                        }
                    );


                    button.classList.add(
                        "active"
                    );


                    const target =
                        $(button.dataset.target);


                    if (target) {

                        target.scrollIntoView({
                            behavior: "smooth",
                            block: "start"
                        });

                    }

                }
            );

        }
    );

}


/* =========================================================
   FILE UPLOAD
========================================================= */

function setupUpload() {

    const input =
        $("dgrFile");


    if (!input) {

        console.error(
            "ERROR: #dgrFile does not exist."
        );

        return;

    }


    input.addEventListener(
        "change",
        function (event) {

            const file =
                event.target.files &&
                event.target.files[0];


            if (file) {

                processDGR(
                    file
                );

            }

        }
    );


    /*
       Optional drag/drop support.
       Works only if dropZone exists.
    */

    const dropZone =
        $("dropZone");


    if (!dropZone) {
        return;
    }


    dropZone.addEventListener(
        "dragover",
        function (event) {

            event.preventDefault();

            dropZone.classList.add(
                "dragging"
            );

        }
    );


    dropZone.addEventListener(
        "dragleave",
        function () {

            dropZone.classList.remove(
                "dragging"
            );

        }
    );


    dropZone.addEventListener(
        "drop",
        function (event) {

            event.preventDefault();

            dropZone.classList.remove(
                "dragging"
            );


            const file =
                event.dataTransfer.files &&
                event.dataTransfer.files[0];


            if (file) {

                processDGR(
                    file
                );

            }

        }
    );

}


/* =========================================================
   REMOVE FILE BUTTON
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
   PROCESS DGR
========================================================= */

function processDGR(file) {

    if (!file) {
        return;
    }


    if (
        !/\.(xlsx|xls|csv)$/i.test(
            file.name
        )
    ) {

        alert(
            "Please upload a valid Excel file (.xlsx, .xls, or .csv)."
        );

        return;

    }


    /*
       Confirm SheetJS exists BEFORE trying
       to read the workbook.
    */

    if (
        typeof XLSX === "undefined"
    ) {

        alert(
            "SheetJS is not loaded.\n\n" +
            "Please check the XLSX script in index.html."
        );

        return;

    }


    setStatus(
        "Reading DGR workbook..."
    );


    const reader =
        new FileReader();


    reader.onload =
        function (event) {

            try {

                workbook =
                    XLSX.read(
                        new Uint8Array(
                            event.target.result
                        ),
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
                        "No worksheets were found in the uploaded workbook."
                    );

                }


                updateFileInformation(
                    file
                );


                showAnalytics();


                renderDashboard();


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
                    "Unable to read the DGR."
                );


                alert(
                    "Unable to read the DGR.\n\n" +
                    error.message
                );

            }

        };


    reader.onerror =
        function () {

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


    const required =
        [
            "Dashboard",
            "Annual_KPI",
            "Daily_KPI",
            "PA",
            "Curtailment records"
        ];


    container.innerHTML =
        "";


    required.forEach(
        function (name) {

            const badge =
                document.createElement(
                    "span"
                );


            badge.className =
                "sheet-badge";


            const exists =
                workbook.SheetNames.some(
                    function (actual) {

                        return (
                            normalizeSheetName(actual) ===
                            normalizeSheetName(name)
                        );

                    }
                );


            if (exists) {

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
   NORMALIZE SHEET NAME
========================================================= */

function normalizeSheetName(
    value
) {

    return String(
        value || ""
    )
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


    const wanted =
        normalizeSheetName(
            requestedName
        );


    const actual =
        workbook.SheetNames.find(
            function (sheetName) {

                return (
                    normalizeSheetName(
                        sheetName
                    ) ===
                    wanted
                );

            }
        );


    if (!actual) {
        return null;
    }


    return workbook.Sheets[
        actual
    ];

}


/* =========================================================
   SHEET TO RAW MATRIX
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
            raw: true,
            defval: null,
            blankrows: false
        }
    );

}


/* =========================================================
   EXCEL COLUMN LETTER -> ARRAY INDEX
========================================================= */

function columnIndex(
    column
) {

    let result =
        0;


    for (
        const character
        of column.toUpperCase()
    ) {

        result =
            result * 26 +
            character.charCodeAt(0) -
            64;

    }


    return result - 1;

}


/* =========================================================
   GET CELL USING EXCEL COLUMN LETTER
========================================================= */

function getCell(
    row,
    column
) {

    if (!row) {
        return null;
    }


    return row[
        columnIndex(
            column
        )
    ];

}


/* =========================================================
   NUMBER PARSER
========================================================= */

function numberValue(
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


    if (
        !text ||
        text.startsWith("#")
    ) {

        return null;

    }


    const valueNumber =
        Number(text);


    if (
        !Number.isFinite(
            valueNumber
        )
    ) {

        return null;

    }


    return valueNumber;

}


/* =========================================================
   DATE PARSER
========================================================= */

function dateValue(
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
       Already a JavaScript Date
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


        return new Date(
            value.getTime()
        );

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


            if (
                parsed
            ) {

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

        catch (_) {

            return null;

        }

    }


    const text =
        String(
            value
        ).trim();


    /*
       DD/MM/YYYY
       DD-MM-YYYY
    */

    let match =
        text.match(
            /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/
        );


    if (
        match
    ) {

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


        const result =
            new Date(
                year,
                month,
                day
            );


        if (
            result.getFullYear() === year &&
            result.getMonth() === month &&
            result.getDate() === day
        ) {

            return result;

        }

    }


    /*
       DD-MMM-YYYY
    */

    match =
        text.match(
            /^(\d{1,2})[\/\-]([A-Za-z]{3,9})[\/\-](\d{2,4})/
        );


    if (
        match
    ) {

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


    /*
       Browser fallback
    */

    const result =
        new Date(
            text
        );


    if (
        !isNaN(
            result.getTime()
        )
    ) {

        return result;

    }


    return null;

}


/* =========================================================
   SHORT DATE
========================================================= */

function formatShortDate(
    value
) {

    const date =
        value instanceof Date
            ? value
            : dateValue(
                value
            );


    if (!date) {
        return "—";
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
   FULL DATE
========================================================= */

function formatFullDate(
    value
) {

    const date =
        value instanceof Date
            ? value
            : dateValue(
                value
            );


    if (!date) {
        return "—";
    }


    return date.toLocaleDateString(
        "en-IN",
        {
            day: "2-digit",
            month: "short",
            year: "numeric"
        }
    );

}


/* =========================================================
   TIME PARSER
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
       Excel time represented as Date
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
       Excel time represented as fraction
       of a day.

       0.5 = 12:00
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


        /*
           Full Excel datetime serial
        */

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
        ).trim();


    /*
       HH:MM
       HH:MM:SS
       HH:MM AM
       HH:MM PM
    */

    const match =
        text.match(
            /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i
        );


    if (
        match
    ) {

        let hours =
            Number(
                match[1]
            );


        const minutes =
            Number(
                match[2]
            );


        const period =
            match[4]
                ? match[4].toUpperCase()
                : null;


        if (
            period === "PM" &&
            hours < 12
        ) {

            hours += 12;

        }


        if (
            period === "AM" &&
            hours === 12
        ) {

            hours = 0;

        }


        if (
            hours >= 0 &&
            hours <= 23 &&
            minutes >= 0 &&
            minutes <= 59
        ) {

            return (
                hours * 60 +
                minutes
            );

        }

    }


    return null;

}


/* =========================================================
   FORMAT TIME
========================================================= */

function minutesToTime(
    minutes
) {

    if (
        !Number.isFinite(
            Number(
                minutes
            )
        )
    ) {

        return "—";

    }


    let value =
        Math.round(
            Number(
                minutes
            )
        );


    /*
       Keep within one day.
    */

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
   NUMBER DISPLAY
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
   TEXT HELPER
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
   ANALYTICS VISIBILITY
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
        function (id) {

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
        function (id) {

            const element =
                $(id);


            if (element) {

                element.style.display =
                    "";

            }

        }
    );


    if ($("emptyState")) {

        $("emptyState").classList.add(
            "hidden"
        );

    }

}


/* =========================================================
   CHART MANAGEMENT
========================================================= */

function destroyChart(
    id
) {

    if (
        charts[id]
    ) {

        try {

            charts[id].destroy();

        }

        catch (_) {}

    }


    charts[id] =
        null;

}


function destroyAllCharts() {

    Object.keys(
        charts
    )
    .forEach(
        function (id) {

            destroyChart(
                id
            );

        }
    );

}


/* =========================================================
   DYNAMIC CARDS
========================================================= */

function removeDynamicCards() {

    [
        "paPercentageCard",
        "breakdownTimelineCard",
        "systemLossMwhCard",
        "curtailmentTableCard",
        "curtailmentGanttCard"
    ]
    .forEach(
        function (id) {

            const element =
                $(id);


            if (element) {

                element.remove();

            }

        }
    );

}


/* =========================================================
   PREPARE SCROLLABLE CHART
========================================================= */

function prepareScrollableCanvas(
    canvas,
    width
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


    if (!wrapper) {

        wrapper =
            document.createElement(
                "div"
            );


        wrapper.className =
            "scroll-chart-container";


        wrapper.style.width =
            "100%";


        wrapper.style.height =
            "100%";


        wrapper.style.overflowX =
            "auto";


        wrapper.style.overflowY =
            "hidden";


        wrapper.style.position =
            "relative";


        wrapper.style.scrollbarWidth =
            "thin";


        parent.insertBefore(
            wrapper,
            canvas
        );


        wrapper.appendChild(
            canvas
        );

    }


    canvas.style.display =
        "block";


    canvas.style.width =
        `${width}px`;


    canvas.style.height =
        "100%";


    canvas.style.maxWidth =
        "none";


    return wrapper;

}


/* =========================================================
   RENDER PA PERCENTAGE
========================================================= */

function renderPAPercentage(
    records
) {

    if (
        !records ||
        !records.length
    ) {

        return;

    }


    const section =
        $("paSection");


    if (!section) {
        return;
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


        card.style.marginBottom =
            "14px";


        card.innerHTML = `

            <div class="chart-heading">

                <div>

                    <h3>
                        Plant Availability
                    </h3>

                    <span>
                        Daily PA (%) · Daily_KPI Column S
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


        const firstCard =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (firstCard) {

            section.insertBefore(
                card,
                firstCard
            );

        }

        else {

            section.appendChild(
                card
            );

        }

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
        records.map(
            function (row) {

                const day =
                    row.date.getDate();


                /*
                   For a monthly DGR show
                   2,4,6,8,10...
                */

                if (
                    records.length >= 20 &&
                    records.length <= 35
                ) {

                    return (
                        day % 2 === 0
                            ? String(day)
                            : ""
                    );

                }


                return formatShortDate(
                    row.date
                );

            }
        );


    const width =
        Math.max(
            650,
            records.length * 60
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    chartRegistrySet(
        "paPercentageChart",

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
                                "Plant Availability",

                            data:
                                records.map(
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
                                0.2,

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
        )
    );

}


/* =========================================================
   DAILY KPI CHARTS
========================================================= */

function renderDailyKPICharts(
    records
) {

    if (!records.length) {
        return;
    }


    const labels =
        records.map(
            function (row) {

                const day =
                    row.date.getDate();


                /*
                   Do not allow date labels
                   to become 29,31,31,29.

                   For one month, show even days.
                */

                if (
                    records.length >= 20 &&
                    records.length <= 35
                ) {

                    return (
                        day % 2 === 0
                            ? String(day)
                            : ""
                    );

                }


                return formatShortDate(
                    row.date
                );

            }
        );


    /*
       PR
    */

    createScrollableLineChart(
        "prChart",
        labels,
        records.map(
            row =>
                row.pr
        ),
        "Performance Ratio",
        "PR (%)"
    );


    /*
       OPERATING HOURS
    */

    createScrollableLineChart(
        "hoursChart",
        labels,
        records.map(
            row =>
                row.hours
        ),
        "Operating Hours",
        "Operating Hours"
    );


    /*
       SYSTEM LOSS
    */

    createScrollableLineChart(
        "lossChart",
        labels,
        records.map(
            row =>
                row.loss
        ),
        "System Losses",
        "System Loss (%)"
    );


    /*
       DASHBOARD PR
    */

    createScrollableLineChart(
        "dashboardPRChart",
        labels,
        records.map(
            row =>
                row.pr
        ),
        "Performance Ratio",
        "PR (%)"
    );


    /*
       DASHBOARD LOSS
    */

    createScrollableLineChart(
        "dashboardLossChart",
        labels,
        records.map(
            row =>
                row.loss
        ),
        "System Losses",
        "System Loss (%)"
    );

}


/* =========================================================
   SCROLLABLE LINE CHART
========================================================= */

function createScrollableLineChart(
    canvasId,
    labels,
    values,
    datasetLabel,
    yTitle
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


    const width =
        Math.max(
            650,
            labels.length * 60
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    chartRegistrySet(
        canvasId,

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
                                    0

                            }

                        },


                        y: {

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
        )
    );

}


/* =========================================================
   PA PLANT UNAVAILABILITY DATA
========================================================= */

function readPlantUnavailability() {

    const rows =
        sheetToMatrix(
            getSheet(
                "PA"
            )
        );


    const records = [];


    rows.forEach(
        function (row) {

            const issue =
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
                issue === null ||
                issue === undefined
            ) {

                return;

            }


            const text =
                String(
                    issue
                ).trim();


            if (
                !text ||
                text === "#REF!"
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
                finish <
                start
            ) {

                finish +=
                    1440;

            }


            records.push({

                issue:
                    text,

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
            "No Plant Unavailability records found."
        );

        return;

    }


    /*
       48 half-hour periods × 48px.
       This makes the chart wide enough to
       scroll without enlarging the card.
    */

    prepareScrollableCanvas(
        canvas,
        2304
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


    chartRegistrySet(
        "paChart",

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
                                    function (items) {

                                        const index =
                                            items[0]
                                                ?.dataIndex;


                                        return (
                                            records[
                                                index
                                            ]?.issue ||
                                            ""
                                        );

                                    },


                                label:
                                    function (context) {

                                        const item =
                                            records[
                                                context
                                                    .dataIndex
                                            ];


                                        if (!item) {
                                            return "";
                                        }


                                        return [

                                            `Start: ${minutesToTime(
                                                item.start
                                            )}`,

                                            `End: ${minutesToTime(
                                                item.end
                                            )}`,

                                            `Duration: ${
                                                item.end -
                                                item.start
                                            } min`

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
                                    function (value) {

                                        return minutesToTime(
                                            value
                                        );

                                    }

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
        )
    );

}


/* =========================================================
   BREAKDOWN TIMELINE DATA
========================================================= */

function readBreakdownTimeline() {

    const rows =
        sheetToMatrix(
            getSheet(
                "PA"
            )
        );


    const grouped =
        new Map();


    rows.forEach(
        function (row) {

            const date =
                dateValue(
                    getCell(
                        row,
                        "B"
                    )
                );


            const breakdown =
                numberValue(
                    getCell(
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
                !grouped.has(
                    key
                )
            ) {

                grouped.set(
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


            grouped.get(
                key
            ).total +=
                breakdown;

        }
    );


    return Array.from(
        grouped.values()
    )
    .sort(
        (a, b) =>
            a.date -
            b.date
    );

}


/* =========================================================
   BREAKDOWN TIMELINE CARD
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
                        Same-date breakdown durations combined from PA · Column AG
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


        section.appendChild(
            card
        );

    }


    return card;

}


/* =========================================================
   BREAKDOWN TIMELINE CHART
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
        readBreakdownTimeline();


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No breakdown timeline data found."
        );

        return;

    }


    chartRegistrySet(
        "breakdownChart",

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
                                formatShortDate(
                                    record.date
                                )
                        ),

                    datasets: [

                        {

                            label:
                                "Breakdown Time (min)",

                            data:
                                records.map(
                                    record =>
                                        record.total
                                ),

                            borderWidth:
                                1,

                            borderRadius:
                                4

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

                            ticks: {

                                stepSize:
                                    1

                            },

                            title: {

                                display:
                                    true,

                                text:
                                    "Breakdown Time (minutes)"

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
        )
    );

}


/* =========================================================
   SYSTEM LOSS MWh
========================================================= */

function readSystemLossMWh() {

    const rows =
        sheetToMatrix(
            getSheet(
                "PA"
            )
        );


    const grouped =
        new Map();


    rows.forEach(
        function (row) {

            const date =
                dateValue(
                    getCell(
                        row,
                        "B"
                    )
                );


            const loss =
                numberValue(
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
                dayKey(
                    date
                );


            if (
                !grouped.has(
                    key
                )
            ) {

                grouped.set(
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


            grouped.get(
                key
            ).total +=
                loss;

        }
    );


    return Array.from(
        grouped.values()
    )
    .sort(
        (a, b) =>
            a.date -
            b.date
    );

}


/* =========================================================
   SYSTEM LOSS MWh CARD
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


        section.appendChild(
            card
        );

    }


    return card;

}


/* =========================================================
   SYSTEM LOSS MWh CHART
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


    const labels =
        records.map(
            record =>
                formatShortDate(
                    record.date
                )
        );


    const width =
        Math.max(
            700,
            records.length * 60
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    chartRegistrySet(
        "systemLossMwhChart",

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
                                "System Loss (MWh)",

                            data:
                                records.map(
                                    record =>
                                        record.total
                                ),

                            borderWidth:
                                1,

                            borderRadius:
                                4

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
                                    true,

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
        )
    );

}


/* =========================================================
   CURTAILMENT DATA
========================================================= */

function readCurtailment() {

    const rows =
        sheetToMatrix(
            getSheet(
                "Curtailment records"
            )
        );


    const records = [];


    rows.forEach(
        function (row) {

            const date =
                dateValue(
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
                numberValue(
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


            let actualEnd =
                end;


            if (
                actualEnd <
                start
            ) {

                actualEnd +=
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
                    actualEnd,

                loss:
                    loss === null
                        ? 0
                        : loss

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
   CURTAILMENT TABLE CARD
========================================================= */

function ensureCurtailmentTableCard() {

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
                        Loss of generation merged across the same date · Column R
                    </span>

                </div>

                <span class="chart-type">
                    TABLE
                </span>

            </div>

            <div id="curtailmentTable"></div>

        `;


        section.appendChild(
            card
        );

    }


    return $("curtailmentTable");

}


/* =========================================================
   CURTAILMENT TABLE
========================================================= */

function renderCurtailmentTable(
    records
) {

    const container =
        ensureCurtailmentTableCard();


    if (!container) {
        return;
    }


    const grouped =
        new Map();


    records.forEach(
        function (record) {

            if (
                !grouped.has(
                    record.key
                )
            ) {

                grouped.set(
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
                grouped.get(
                    record.key
                );


            item.loss +=
                record.loss;


            item.intervals++;

        }
    );


    const daily =
        Array.from(
            grouped.values()
        )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );


    if (!daily.length) {

        container.innerHTML = `

            <div style="
                padding:16px;
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
                        Curtailment Intervals
                    </th>

                </tr>

            </thead>

            <tbody>
    `;


    daily.forEach(
        function (item) {

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


    container.innerHTML =
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

function ensureCurtailmentGanttCard() {

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


        section.appendChild(
            card
        );

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
        ensureCurtailmentGanttCard();


    if (!canvas) {
        return;
    }


    destroyChart(
        "curtailmentGanttChart"
    );


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No curtailment duration records found."
        );

        return;

    }


    /*
       48 × 15-minute intervals from
       06:00 to 18:00.

       2400px internal chart width.
    */

    prepareScrollableCanvas(
        canvas,
        2400
    );


    const uniqueDates = [];


    records.forEach(
        function (record) {

            const label =
                formatShortDate(
                    record.date
                );


            if (
                !uniqueDates.includes(
                    label
                )
            ) {

                uniqueDates.push(
                    label
                );

            }

        }
    );


    const datasets =
        records
            .filter(
                record =>
                    record.end > 360 &&
                    record.start < 1080
            )
            .map(
                function (record) {

                    return {

                        label:
                            `${formatShortDate(
                                record.date
                            )} ${minutesToTime(
                                record.start
                            )}`,

                        data: [

                            {

                                x: [

                                    Math.max(
                                        360,
                                        record.start
                                    ),

                                    Math.min(
                                        1080,
                                        record.end
                                    )

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


    chartRegistrySet(
        "curtailmentGanttChart",

        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels:
                        uniqueDates,

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

                            labels:
                                uniqueDates,

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
        )
    );

}


/* =========================================================
   ANNUAL ENERGY
========================================================= */

function renderEnergyChart() {

    const sheet =
        getSheet(
            "Annual_KPI"
        );


    if (!sheet) {

        console.warn(
            "Annual_KPI worksheet not found."
        );

        return;

    }


    const rows =
        sheetToMatrix(
            sheet
        );


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
            numberValue(
                getCell(
                    row,
                    "E"
                )
            )
        );


        measured.push(
            numberValue(
                getCell(
                    row,
                    "F"
                )
            )
        );

    }


    const monthLabels = [

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


    chartRegistrySet(
        "energyChart",

        new Chart(
            canvas.getContext(
                "2d"
            ),
            {

                type:
                    "bar",

                data: {

                    labels:
                        monthLabels,

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
        )
    );

}


/* =========================================================
   MAIN RENDER
========================================================= */

function renderDashboard() {

    destroyAllCharts();

    removeDynamicCards();


    /*
       DAILY KPI
    */

    const daily =
        readDailyKPI();


    if (
        daily.length
    ) {

        renderDashboardKPI(
            daily
        );


        renderDailyCharts(
            daily
        );

    }


    /*
       PA
    */

    renderPAPercentage(
        daily
    );


    renderPlantUnavailability();


    renderBreakdownTimeline();


    renderSystemLossMWh();


    /*
       CURTAILMENT
    */

    const curtailment =
        readCurtailment();


    renderCurtailmentLossChart(
        curtailment
    );


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

}


/* =========================================================
   DASHBOARD KPI
========================================================= */

function renderDashboardKPI(
    records
) {

    if (
        !records ||
        !records.length
    ) {

        return;

    }


    const latest =
        records[
            records.length - 1
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
   READ DAILY KPI
========================================================= */

function readDailyKPI() {

    const rows =
        sheetToMatrix(
            getSheet(
                "Daily_KPI"
            )
        );


    const records = [];


    rows.forEach(
        function (row) {

            const date =
                dateValue(
                    getCell(
                        row,
                        "B"
                    )
                );


            if (!date) {
                return;
            }


            const hours =
                numberValue(
                    getCell(
                        row,
                        "I"
                    )
                );


            const paRaw =
                numberValue(
                    getCell(
                        row,
                        "S"
                    )
                );


            const prRaw =
                numberValue(
                    getCell(
                        row,
                        "V"
                    )
                );


            const lossRaw =
                numberValue(
                    getCell(
                        row,
                        "AD"
                    )
                );


            /*
               Convert decimals to percentages.

               0.95 → 95
               0.84 → 84
               0.019 → 1.9
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


            /*
               Keep rows with valid dates.
            */

            records.push({

                date,

                pa,

                pr,

                hours,

                loss

            });

        }
    );


    records.sort(
        (a, b) =>
            a.date -
            b.date
    );


    return records;

}


/* =========================================================
   GENERIC CANVAS MESSAGE
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


    ctx.font =
        "12px Inter, Arial";


    ctx.fillStyle =
        "#879397";


    ctx.textAlign =
        "center";


    ctx.textBaseline =
        "middle";


    ctx.fillText(
        message,
        canvas.width / 2,
        canvas.height / 2
    );


    ctx.restore();

}


/* =========================================================
   CURTAILMENT LOSS CHART
========================================================= */

function renderCurtailmentLossChart(
    records
) {

    const canvas =
        $("curtailmentChart");


    if (!canvas) {
        return;
    }


    destroyChart(
        "curtailmentChart"
    );


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No curtailment loss records found."
        );

        return;

    }


    const grouped =
        new Map();


    records.forEach(
        function (record) {

            if (
                !grouped.has(
                    record.key
                )
            ) {

                grouped.set(
                    record.key,
                    {

                        date:
                            record.date,

                        loss:
                            0

                    }
                );

            }


            grouped.get(
                record.key
            ).loss +=
                record.loss;

        }
    );


    const daily =
        Array.from(
            grouped.values()
        )
        .sort(
            (a, b) =>
                a.date -
                b.date
        );


    const labels =
        daily.map(
            record =>
                formatShortDate(
                    record.date
                )
        );


    const values =
        daily.map(
            record =>
                record.loss
        );


    const width =
        Math.max(
            750,
            labels.length * 65
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    chartRegistrySet(
        "curtailmentChart",

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
                                "Curtailment Loss (MWh)",

                            data:
                                values,

                            borderWidth:
                                2,

                            pointRadius:
                                3,

                            pointHoverRadius:
                                6,

                            tension:
                                0.2,

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
                                    true,

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
                                    "Curtailment Loss (MWh)"

                            }

                        }

                    }

                }

            }
        )
    );


    const total =
        values.reduce(
            (sum, value) =>
                sum +
                value,
            0
        );


    setText(
        "curtailmentSummary",
        `${daily.length} date(s) · ${total.toFixed(2)} MWh total generation loss`
    );

}


/* =========================================================
   RESET
========================================================= */

function resetDashboard() {

    destroyAllCharts();


    removeDynamicCards();


    workbook =
        null;


    if ($("dgrFile")) {

        $("dgrFile").value =
            "";

    }


    $("fileInfo")?.classList.add(
        "hidden"
    );


    $("workbookStatus")?.classList.add(
        "hidden"
    );


    $("emptyState")?.classList.remove(
        "hidden"
    );


    $("dropZone")?.classList.remove(
        "hidden"
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
        function (id) {

            setText(
                id,
                "—"
            );

        }
    );


    setText(
        "sidebarFileName",
        "No DGR uploaded"
    );


    setText(
        "curtailmentSummary",
        "Waiting for DGR data"
    );


    setStatus(
        "Upload a DGR to generate the analytics."
    );


    hideAnalytics();

}


/* =========================================================
   END
========================================================= */
