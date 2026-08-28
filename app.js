/* =========================================================
   SOLAR DGR ANALYTICS
   app.js

   EXACT DGR MAPPINGS
   ---------------------------------------------------------

   DAILY_KPI
      B  = Date
      I  = Operating Hours
      S  = PA (%)
      V  = PR (%)
      AD = System Losses (%)

   PA
      B  = Date
      W  = Issue / Fault
      Z  = Start / Fault Time
      AC = Work Completion / End Time
      AG = Breakdown Time (minutes)
      AL = System Loss (MWh)

   CURTAILMENT RECORDS
      C  = Date
      H  = Start Time
      I  = End Time
      R  = Loss of Generation MWh

   ANNUAL_KPI
      E10:E21 = Budgeted Energy
      F10:F21 = Measured Energy

   ========================================================= */

"use strict";


/* =========================================================
   GLOBAL
========================================================= */

let workbook = null;

const chartRegistry = {};

const dynamicElements = [
    "paPercentageCard",
    "breakdownTimelineCard",
    "systemLossMwhCard",
    "curtailmentTableCard",
    "curtailmentGanttCard"
];


/* =========================================================
   DOM HELPER
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
        document.querySelectorAll(
            ".nav-item"
        );


    buttons.forEach(
        button => {

            button.addEventListener(
                "click",
                () => {

                    buttons.forEach(
                        item =>
                            item.classList.remove(
                                "active"
                            )
                    );


                    button.classList.add(
                        "active"
                    );


                    const target =
                        $(button.dataset.target);


                    if (target) {

                        target.scrollIntoView(
                            {
                                behavior: "smooth",
                                block: "start"
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
            "dgrFile input was not found."
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
       Drag/drop support if the drop zone
       is still present in index.html.
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
                    event.dataTransfer.files &&
                    event.dataTransfer.files[0];


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
   PROCESS FILE
========================================================= */

function processFile(file) {

    const validExtension =
        /\.(xlsx|xls|csv)$/i.test(
            file.name
        );


    if (!validExtension) {

        alert(
            "Please upload an Excel file (.xlsx/.xls) or CSV file."
        );

        return;

    }


    if (
        typeof XLSX === "undefined"
    ) {

        alert(
            "SheetJS is not loaded. Check the XLSX script in index.html."
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
                            type: "array",
                            cellDates: true,
                            cellNF: true
                        }
                    );


                if (
                    !workbook.SheetNames ||
                    !workbook.SheetNames.length
                ) {

                    throw new Error(
                        "No worksheets were found."
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
                    "DGR processing error:",
                    error
                );


                setStatus(
                    "Unable to read the DGR."
                );


                alert(
                    "The DGR could not be read.\n\n" +
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


    $("fileInfo")?.classList.remove(
        "hidden"
    );


    $("workbookStatus")?.classList.remove(
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


    container.innerHTML = "";


    const requiredSheets = [
        "Dashboard",
        "Annual_KPI",
        "Daily_KPI",
        "PA",
        "Curtailment records"
    ];


    requiredSheets.forEach(
        sheetName => {

            const badge =
                document.createElement(
                    "span"
                );


            badge.className =
                "sheet-badge";


            const exists =
                workbook.SheetNames.includes(
                    sheetName
                );


            if (exists) {

                badge.textContent =
                    `${sheetName} ✓`;

            }

            else {

                badge.textContent =
                    `${sheetName} — missing`;

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
   GET WORKSHEET
========================================================= */

function getSheet(
    name
) {

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


    const target =
        name
            .toLowerCase()
            .replace(
                /[\s_-]+/g,
                ""
            );


    const actual =
        workbook.SheetNames.find(
            sheetName =>
                sheetName
                    .toLowerCase()
                    .replace(
                        /[\s_-]+/g,
                        ""
                    ) ===
                target
        );


    return actual
        ? workbook.Sheets[actual]
        : null;

}


/* =========================================================
   RAW MATRIX
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
            header: 1,
            raw: true,
            defval: null,
            blankrows: false
        }
    );

}


/* =========================================================
   COLUMN LETTER
========================================================= */

function columnIndex(
    column
) {

    let result =
        0;


    for (
        const char of column.toUpperCase()
    ) {

        result =
            result * 26 +
            char.charCodeAt(0) -
            64;

    }


    return result - 1;

}


/* =========================================================
   GET CELL BY EXCEL COLUMN
========================================================= */

function cell(
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
   NUMBER
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


    const number =
        Number(text);


    return Number.isFinite(
        number
    )
        ? number
        : null;

}


/* =========================================================
   DATE
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

    }


    if (
        typeof value === "string"
    ) {

        const text =
            value.trim();


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


            const result =
                new Date(
                    y,
                    m,
                    d
                );


            if (
                result.getFullYear() === y &&
                result.getMonth() === m &&
                result.getDate() === d
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


        const parsed =
            new Date(
                text
            );


        if (
            !isNaN(
                parsed.getTime()
            )
        ) {

            return parsed;

        }

    }


    return null;

}


/* =========================================================
   DATE KEY
========================================================= */

function dayKey(
    date
) {

    if (
        !(date instanceof Date)
    ) {

        return null;

    }


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
   TIME -> MINUTES
========================================================= */

function timeMinutes(
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
       Excel time as Date
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
                1440
            );

        }


        const parsed =
            XLSX.SSF.parse_date_code(
                value
            );


        if (parsed) {

            return (
                parsed.H * 60 +
                parsed.M +
                Math.floor(
                    (parsed.S || 0) / 60
                )
            );

        }

    }


    const text =
        String(value)
            .trim();


    const match =
        text.match(
            /^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i
        );


    if (match) {

        let h =
            Number(
                match[1]
            );


        const m =
            Number(
                match[2]
            );


        if (
            match[4]
        ) {

            const ap =
                match[4].toUpperCase();


            if (
                ap === "PM" &&
                h < 12
            ) {

                h += 12;

            }


            if (
                ap === "AM" &&
                h === 12
            ) {

                h = 0;

            }

        }


        return (
            h * 60 +
            m
        );

    }


    return null;

}


/* =========================================================
   MINUTES -> TIME
========================================================= */

function minuteTime(
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


    const h =
        Math.floor(
            safe / 60
        );


    const m =
        safe % 60;


    return (
        String(
            h
        ).padStart(
            2,
            "0"
        ) +
        ":" +
        String(
            m
        ).padStart(
            2,
            "0"
        )
    );

}


/* =========================================================
   FORMAT DATE
========================================================= */

function shortDate(
    date
) {

    if (
        !(date instanceof Date)
    ) {

        return "";

    }


    return date.toLocaleDateString(
        "en-GB",
        {
            day: "2-digit",
            month: "short"
        }
    );

}


function longDate(
    date
) {

    if (
        !(date instanceof Date)
    ) {

        return "";

    }


    return date.toLocaleDateString(
        "en-GB",
        {
            day: "2-digit",
            month: "short",
            year: "numeric"
        }
    );

}


/* =========================================================
   FORMAT NUMBER
========================================================= */

function formattedNumber(
    value,
    decimals = 2
) {

    if (
        value === null ||
        value === undefined ||
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
   HIDE ANALYTICS
========================================================= */

function hideAnalytics() {

    [
        "dashboardSection",
        "paSection",
        "performanceSection",
        "curtailmentSection",
        "energySection"
    ].forEach(
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


/* =========================================================
   SHOW ANALYTICS
========================================================= */

function showAnalytics() {

    [
        "dashboardSection",
        "paSection",
        "performanceSection",
        "curtailmentSection",
        "energySection"
    ].forEach(
        id => {

            const element =
                $(id);


            if (element) {

                element.style.display =
                    "block";

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
   DESTROY CHART
========================================================= */

function destroyChart(
    id
) {

    if (
        chartRegistry[id]
    ) {

        try {

            chartRegistry[id].destroy();

        }

        catch (_) {}

    }


    delete chartRegistry[id];

}


/* =========================================================
   DESTROY ALL CHARTS
========================================================= */

function destroyAllCharts() {

    Object.keys(
        chartRegistry
    ).forEach(
        destroyChart
    );

}


/* =========================================================
   REMOVE DYNAMIC ELEMENTS
========================================================= */

function removeDynamicElements() {

    dynamicElements.forEach(
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
   MAKE SCROLLABLE CANVAS
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


    /*
       Create the scrolling viewport.
    */

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


        parent.insertBefore(
            wrapper,
            canvas
        );


        wrapper.appendChild(
            canvas
        );

    }


    /*
       The wrapper occupies exactly the
       available chart-card space.
    */

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

    wrapper.style.boxSizing =
        "border-box";


    /*
       Wide internal canvas.

       This is what creates the horizontal
       scrolling area.
    */

    const finalWidth =
        Math.max(
            width || 800,
            parent.clientWidth || 800
        );


    canvas.classList.add(
        "scroll-chart-canvas"
    );


    canvas.style.display =
        "block";


    canvas.style.width =
        `${finalWidth}px`;


    canvas.style.minWidth =
        `${finalWidth}px`;


    canvas.style.maxWidth =
        "none";


    canvas.style.height =
        "100%";


    /*
       Prevent Chart.js from forcing the
       canvas back to the card width.
    */

    canvas.removeAttribute(
        "width"
    );

    canvas.removeAttribute(
        "height"
    );


    return wrapper;

}
/* =========================================================
   CREATE PA PERCENTAGE CARD
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


    return card;

}


/* =========================================================
   READ DAILY KPI
========================================================= */

function readDailyKPI() {

    const rows =
        sheetMatrix(
            getSheet(
                "Daily_KPI"
            )
        );


    const records = [];


    rows.forEach(
        row => {

            const date =
                dateValue(
                    cell(
                        row,
                        "B"
                    )
                );


            if (!date) {
                return;
            }


            const paRaw =
                numberValue(
                    cell(
                        row,
                        "S"
                    )
                );


            const prRaw =
                numberValue(
                    cell(
                        row,
                        "V"
                    )
                );


            const hours =
                numberValue(
                    cell(
                        row,
                        "I"
                    )
                );


            const lossRaw =
                numberValue(
                    cell(
                        row,
                        "AD"
                    )
                );


            /*
               DGR percentage fields are decimal
               values such as 0.84.
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
   RENDER DAILY CHARTS
========================================================= */

function renderDailyCharts(
    records
) {

    if (!records.length) {
        return;
    }


    /*
       Date labels.

       For a normal monthly dataset:
       odd dates are hidden so the visible
       sequence becomes:

       2 4 6 8 10 12 14 ...

       The data points themselves remain
       daily; only the visible labels are
       reduced.
    */

    const labels =
        records.map(
            record =>
                shortDate(
                    record.date
                )
        );


    const axisLabels =
        records.map(
            record => {

                const day =
                    record.date.getDate();


                const total =
                    records.length;


                /*
                   If approximately a month,
                   show every second day.

                   Otherwise let Chart.js
                   auto skip.
                */

                if (
                    total >= 20 &&
                    total <= 35
                ) {

                    return (
                        day % 2 === 0
                            ? day
                            : ""
                    );

                }


                return shortDate(
                    record.date
                );

            }
        );


    /*
       PA
    */

    renderPAPercentage(
        records
    );


    /*
       PR
    */

    makeScrollableLineChart(
        "prChart",
        axisLabels,
        records.map(
            r => r.pr
        ),
        "PR (%)",
        "Performance Ratio (%)",
        false
    );


    /*
       Operating Hours
    */

    makeScrollableLineChart(
        "hoursChart",
        axisLabels,
        records.map(
            r => r.hours
        ),
        "Operating Hours",
        "Operating Hours",
        true
    );


    /*
       System Loss %
    */

    makeScrollableLineChart(
        "lossChart",
        axisLabels,
        records.map(
            r => r.loss
        ),
        "System Losses (%)",
        "System Loss (%)",
        true
    );


    /*
       Dashboard PR
    */

    makeScrollableLineChart(
        "dashboardPRChart",
        axisLabels,
        records.map(
            r => r.pr
        ),
        "PR (%)",
        "Performance Ratio (%)",
        false
    );


    /*
       Dashboard Loss
    */

    makeScrollableLineChart(
        "dashboardLossChart",
        axisLabels,
        records.map(
            r => r.loss
        ),
        "System Losses (%)",
        "System Loss (%)",
        true
    );

}


/* =========================================================
   PA PERCENTAGE CHART
========================================================= */

function renderPAPercentage(
    records
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
        records.map(
            record => {

                const day =
                    record.date.getDate();


                return (
                    day % 2 === 0
                        ? day
                        : ""
                );

            }
        );


    const width =
        Math.max(
            card.clientWidth || 500,
            records.length * 58
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    chartRegistry.paPercentageChart =
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
                                records.map(
                                    r =>
                                        r.pa
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
        );

}


/* =========================================================
   SCROLLABLE LINE CHART
========================================================= */

function makeScrollableLineChart(
    canvasId,
    labels,
    values,
    datasetLabel,
    yAxisLabel,
    beginAtZero
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
            parent?.clientWidth || 500,
            labels.length * 58
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    chartRegistry[canvasId] =
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

                            beginAtZero,

                            title: {

                                display:
                                    true,

                                text:
                                    yAxisLabel

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
   READ PLANT UNAVAILABILITY
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

            const fault =
                cell(
                    row,
                    "W"
                );


            const start =
                timeMinutes(
                    cell(
                        row,
                        "Z"
                    )
                );


            const end =
                timeMinutes(
                    cell(
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


    const parent =
        canvas.parentElement;


    /*
       48 half-hour intervals × 45px
       creates a wide internal timeline.

       The visible card does NOT grow.
       The user scrolls horizontally.
    */

    const width =
        Math.max(
            parent?.clientWidth || 700,
            2304
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    const labels =
        records.map(
            record =>
                record.fault
        );


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


    chartRegistry.paChart =
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
                                    items =>
                                        records[
                                            items[0]
                                                .dataIndex
                                        ]?.fault ||
                                        "",


                                label:
                                    context => {

                                        const item =
                                            records[
                                                context
                                                    .dataIndex
                                            ];


                                        return [

                                            `Start: ${minuteTime(item.start)}`,

                                            `End: ${minuteTime(item.end)}`,

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
                                        minuteTime(
                                            value
                                        ),

                                maxRotation:
                                    0

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

function readBreakdownTimeline() {

    const rows =
        sheetMatrix(
            getSheet(
                "PA"
            )
        );


    const byDate =
        new Map();


    rows.forEach(
        row => {

            const date =
                dateValue(
                    cell(
                        row,
                        "B"
                    )
                );


            const minutes =
                numberValue(
                    cell(
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
                dayKey(
                    date
                );


            if (
                !byDate.has(
                    key
                )
            ) {

                byDate.set(
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


            byDate.get(
                key
            ).total +=
                minutes;

        }
    );


    return Array.from(
        byDate.values()
    )
        .sort(
            (a, b) =>
                a.date -
                b.date
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
                        Same-date breakdown times combined from PA · Column AG
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


        const ganttCard =
            section.querySelector(
                ".chart-card.full-card"
            );


        if (ganttCard) {

            ganttCard.insertAdjacentElement(
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
        readBreakdownTimeline();


    if (!records.length) {

        showCanvasMessage(
            canvas,
            "No breakdown timeline data found."
        );

        return;

    }


    /*
       Fixed requested range:
       0–13 minutes
       interval = 1 minute.
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
                                shortDate(
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


    const byDate =
        new Map();


    rows.forEach(
        row => {

            const date =
                dateValue(
                    cell(
                        row,
                        "B"
                    )
                );


            const loss =
                numberValue(
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
                !byDate.has(
                    key
                )
            ) {

                byDate.set(
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


            byDate.get(
                key
            ).loss +=
                loss;

        }
    );


    return Array.from(
        byDate.values()
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
        );

}


/* =========================================================
   CURTAILMENT DATA
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
                dateValue(
                    cell(
                        row,
                        "C"
                    )
                );


            const start =
                timeMinutes(
                    cell(
                        row,
                        "H"
                    )
                );


            const end =
                timeMinutes(
                    cell(
                        row,
                        "I"
                    )
                );


            const loss =
                numberValue(
                    cell(
                        row,
                        "R"
                    )
                );


            if (!date) {
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
                        Loss of Generation merged for each date · Column R
                    </span>

                </div>

                <span class="chart-type">
                    TABLE
                </span>

            </div>

            <div id="curtailmentTable"></div>

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


    const byDate =
        new Map();


    records.forEach(
        record => {

            if (
                !byDate.has(
                    record.key
                )
            ) {

                byDate.set(
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
                byDate.get(
                    record.key
                );


            item.loss +=
                record.loss;


            item.intervals++;

        }
    );


    const daily =
        Array.from(
            byDate.values()
        )
        .sort(
            (a, b) =>
                a.date -
                b.date
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
        item => {

            html += `

                <tr>

                    <td style="
                        padding:10px;
                        border-bottom:1px solid #edf2f3;
                    ">
                        ${longDate(
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
        ensureCurtailmentGanttCard();


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


    const parent =
        canvas.parentElement;


    const width =
        Math.max(
            parent?.clientWidth || 700,
            2400
        );


    prepareScrollableCanvas(
        canvas,
        width
    );


    /*
       Group the dates for Y-axis.
    */

    const uniqueDates = [];


    records.forEach(
        record => {

            const label =
                shortDate(
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
        records.map(
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
                        `${shortDate(record.date)} ${minuteTime(start)}–${minuteTime(end)}`,

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
            dataset =>
                dataset.data[0].x[1] >
                dataset.data[0].x[0]
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


                                        return (

                                            `Time: ${minuteTime(
                                                raw.x[0]
                                            )} – ${minuteTime(
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
                                        minuteTime(
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
                cell(
                    row,
                    "E"
                )
            )
        );


        measured.push(
            numberValue(
                cell(
                    row,
                    "F"
                )
            )
        );

    }


    const validBudget =
        budget.filter(
            value =>
                value !== null
        );


    const validMeasured =
        measured.filter(
            value =>
                value !== null
        );


    const totalBudget =
        validBudget.reduce(
            (sum, value) =>
                sum + value,
            0
        );


    const totalMeasured =
        validMeasured.reduce(
            (sum, value) =>
                sum + value,
            0
        );


    const variance =
        totalMeasured -
        totalBudget;


    setText(
        "totalBudget",
        `${formattedNumber(
            totalBudget
        )} MWh`
    );


    setText(
        "totalMeasured",
        `${formattedNumber(
            totalMeasured
        )} MWh`
    );


    setText(
        "energyVariance",
        `${formattedNumber(
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


    chartRegistry.energyChart =
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
        );

}


/* =========================================================
   DASHBOARD KPI VALUES
========================================================= */

function renderKPIs() {

    const rows =
        readDailyKPI();


    if (!rows.length) {
        return;
    }


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


    ctx.fillStyle =
        "#879397";


    ctx.font =
        "12px Inter, Arial";


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
   RENDER DASHBOARD
========================================================= */

function renderDashboard() {

    destroyAllCharts();

    removeDynamicElements();


    /*
       1. Daily KPI
    */

    const daily =
        readDailyKPI();


    if (daily.length) {

        renderKPIs();

        renderDailyCharts(
            daily
        );

    }


    /*
       2. PA
    */

    renderPlantUnavailability();

    renderBreakdownTimeline();

    renderSystemLossMWh();


    /*
       3. Curtailment
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
       4. Energy
    */

    renderEnergyChart();

}


/* =========================================================
   RESET
========================================================= */

function resetDashboard() {

    workbook = null;


    destroyAllCharts();

    removeDynamicElements();


    if (fileInput) {

        fileInput.value =
            "";

    }


    $("fileInfo")?.classList.add(
        "hidden"
    );


    $("workbookStatus")?.classList.add(
        "hidden"
    );


    $("dropZone")?.classList.remove(
        "hidden"
    );


    $("emptyState")?.classList.remove(
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
    ].forEach(
        id => {

            setText(
                id,
                "—"
            );

        }
    );


    setText(
        "curtailmentSummary",
        "Waiting for DGR data"
    );

}


/* =========================================================
   END
========================================================= */
