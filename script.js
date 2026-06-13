document.getElementById("fileUpload").addEventListener("change", handleFile);

let allData = [];

function handleFile(e) {
    const file = e.target.files[0];

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        let raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });

        allData = raw.map(r => ({
            CaseID: r.CaseID || "",
            CreatedDate: r.CreatedDate || "",
            Status: r.Status || "",
            Engineer: r.Engineer || "Unassigned"
        }));

        process();
    };

    reader.readAsArrayBuffer(file);
}

function process() {
    const today = new Date();

    allData = allData.map(r => {
        let d = parseDate(r.CreatedDate);
        let days = d ? Math.floor((today - d) / (1000 * 60 * 60 * 24)) : 0;
        return { ...r, DaysOpen: days };
    });

    updateKPIs();
    buildAging();
    buildEngineer();
    buildCases();
}

/* KPI */
function updateKPIs() {
    let total = allData.length;
    let open = allData.filter(x => x.Status.toLowerCase() === "open").length;

    document.getElementById("totalCases").innerText = total;
    document.getElementById("openCases").innerText = open;
    document.getElementById("backlogPercent").innerText =
        total ? ((open / total) * 100).toFixed(2) + "%" : "0%";
}

/* ✅ AGING BUCKETS */
function buildAging() {

    let bucket = {
        "0-30": 0,
        "31-50": 0,
        "51-100": 0,
        "100+": 0
    };

    allData.forEach(r => {
        if (r.DaysOpen <= 30) bucket["0-30"]++;
        else if (r.DaysOpen <= 50) bucket["31-50"]++;
        else if (r.DaysOpen <= 100) bucket["51-100"]++;
        else bucket["100+"]++;
    });

    let max = Math.max(...Object.values(bucket));

    let container = document.getElementById("agingContainer");
    container.innerHTML = "";

    Object.keys(bucket).forEach(k => {
        let value = bucket[k];
        let width = max ? (value / max) * 100 : 0;

        container.innerHTML += `
        <div class="bucket">
            <h4>${k}</h4>
            <p>${value}</p>
            <div class="bar" style="width:${width}%"></div>
        </div>`;
    });
}

/* ENGINEER */
function buildEngineer() {

    let map = {};

    allData.forEach(r => {
        let eng = r.Engineer;

        if (!map[eng]) map[eng] = { total: 0, open: 0 };

        map[eng].total++;
        if (r.Status.toLowerCase() === "open") map[eng].open++;
    });

    let tbody = document.querySelector("#engTable tbody");
    tbody.innerHTML = "";

    Object.keys(map).forEach(k => {
        let t = map[k];
        let pct = ((t.open / t.total) * 100).toFixed(1);

        tbody.innerHTML += `
        <tr>
            <td>${k}</td>
            <td>${t.total}</td>
            <td>${t.open}</td>
            <td>${pct}%</td>
        </tr>`;
    });
}

/* CASES */
function buildCases() {

    let tbody = document.querySelector("#caseTable tbody");
    tbody.innerHTML = "";

    let filtered = allData.filter(x => x.DaysOpen > 50);

    filtered.forEach(r => {
        tbody.innerHTML += `
        <tr>
            <td>${r.CaseID}</td>
            <td>${formatDate(r.CreatedDate)}</td>
            <td>${r.DaysOpen}</td>
            <td>${r.Status}</td>
            <td>${r.Engineer}</td>
        </tr>`;
    });
}

/* DATE PARSE */
function parseDate(val) {
    if (!val) return null;

    if (typeof val === "number") {
        return new Date((val - 25569) * 86400 * 1000);
    }

    let d = new Date(val);
    if (!isNaN(d)) return d;

    let parts = val.split(/[-\/]/);
    if (parts.length === 3) {
        return new Date(parts[2], parts[1] - 1, parts[0]);
    }

    return null;
}

/* FORMAT */
function formatDate(val) {
    let d = parseDate(val);
    if (!d) return "";

    let dd = String(d.getDate()).padStart(2, "0");
    let mm = String(d.getMonth() + 1).padStart(2, "0");
    let yy = d.getFullYear();

    return `${dd}-${mm}-${yy}`;
}