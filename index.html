<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>University Dashboard</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">

<style>
    body {
        font-family: Arial, sans-serif;
        background: #fafbfc;
        margin: 20px;
        color: #333;
    }

    h1 {
        text-align: center;
        color: #2c3e50;
        margin-bottom: 20px;
        font-size: 28px;
    }

    #filters {
        text-align: center;
        margin-bottom: 20px;
    }

    #filters label {
        font-weight: bold;
        margin-right: 8px;
    }

    #filters select {
        padding: 8px;
        margin: 0 10px;
        border-radius: 6px;
        border: 1px solid #ccc;
        font-size: 15px;
    }

    .container {
        display: grid;
        grid-template-columns: 260px 1fr;
        gap: 20px;
    }

    #kpis {
        display: flex;
        flex-direction: column;
        gap: 15px;
    }

    .kpi {
        background: #f2f2f3;
        padding: 15px;
        border-radius: 12px;
        text-align: center;
        box-shadow: 0 2px 6px rgba(0,0,0,0.12);
    }

    .kpi h3 {
        margin: 0;
        font-size: 15px;
        color: #666;
    }

    .kpi p {
        font-size: 24px;
        font-weight: bold;
        margin: 10px 0 0 0;
        color: #000;
    }

    #charts {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 20px;
    }

    .chart {
        height: 420px;
        background: #eeeff0;
        border-radius: 12px;
        padding: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
    }

    .chart-title {
        text-align: center;
        font-weight: bold;
        margin-bottom: 8px;
        font-size: 16px;
        color: #333;
    }

    @media (max-width: 1200px) {
        .container {
            grid-template-columns: 1fr;
        }
        #charts {
            grid-template-columns: 1fr;
        }
    }
</style>

<script src="https://cdn.amcharts.com/lib/5/index.js"></script>
<script src="https://cdn.amcharts.com/lib/5/xy.js"></script>
<script src="https://cdn.amcharts.com/lib/5/percent.js"></script>
<script src="https://cdn.amcharts.com/lib/5/themes/Animated.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>

<body>

<h1>University Dashboard</h1>

<div id="filters">
    <label>Year:</label>
    <select id="yearFilter">
        <option value="All">All</option>
    </select>

    <label>Faculty:</label>
    <select id="facultyFilter">
        <option value="All">All</option>
    </select>
</div>

<div class="container">
    <div id="kpis">
        <div class="kpi"><h3>Total Students</h3><p id="totalStudents">0</p></div>
        <div class="kpi"><h3>Total Fees</h3><p id="totalFees">0</p></div>
        <div class="kpi"><h3>Faculties</h3><p id="numFaculties">0</p></div>
        <div class="kpi"><h3>Professors</h3><p id="numDoctors">0</p></div>
    </div>

    <div id="charts">
        <div>
            <div class="chart-title">Students by Faculty</div>
            <div id="facultyBarChart" class="chart"></div>
        </div>

        <div>
            <div class="chart-title">Students Enrollment per Year</div>
            <div id="lineChart" class="chart"></div>
        </div>

        <div>
            <div class="chart-title">Students Distribution by Major</div>
            <div id="majorPieChart" class="chart"></div>
        </div>

        <div>
            <div class="chart-title">Doctor Salary per Major</div>
            <div id="salaryChart" class="chart"></div>
        </div>
    </div>
</div>

<script>
let data = [];
let charts = {};
let totalUniqueDoctors = 0;

const yearFilter = document.getElementById("yearFilter");
const facultyFilter = document.getElementById("facultyFilter");
const totalStudents = document.getElementById("totalStudents");
const totalFees = document.getElementById("totalFees");
const numFaculties = document.getElementById("numFaculties");
const numDoctors = document.getElementById("numDoctors");

async function loadData() {
    try {
        const res = await fetch("university_students_dataset.xlsx");
        const buf = await res.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        data = XLSX.utils.sheet_to_json(sheet);

        const allDoctorNames = data
            .map(d => (d["Major Doctor"] || "").toString().trim())
            .filter(name => name !== "");
        totalUniqueDoctors = [...new Set(allDoctorNames)].length;

        fillFilters();
        updateDashboard();
    } catch (err) {
        console.error(err);
        alert("Could not load the Excel file.");
    }
}

function fillFilters() {
    const years = [...new Set(data.map(d => d.Year))].sort();
    years.forEach(y => yearFilter.innerHTML += `<option value="${y}">${y}</option>`);

    const faculties = [...new Set(data.map(d => d.Faculty))].sort();
    faculties.forEach(f => facultyFilter.innerHTML += `<option value="${f}">${f}</option>`);
}

function updateDashboard() {
    let filtered = data;

    if (yearFilter.value !== "All") filtered = filtered.filter(d => d.Year == yearFilter.value);
    if (facultyFilter.value !== "All") filtered = filtered.filter(d => d.Faculty === facultyFilter.value);

    totalStudents.textContent = filtered.length;
    totalFees.textContent = filtered.reduce((s,d) => s + (d.Fees || 0),0).toLocaleString();
    numFaculties.textContent = [...new Set(filtered.map(d => d.Faculty))].length;
    numDoctors.textContent = totalUniqueDoctors;

    Object.values(charts).forEach(c => c.dispose());
    charts = {};

    createFacultyBarChart(filtered);
    createYearLineChart(filtered);
    createMajorPieChart(filtered);
    createSalaryBarChart(filtered);
}

function createFacultyBarChart(filtered) {
    const count = {};
    filtered.forEach(d => count[d.Faculty] = (count[d.Faculty] || 0) + 1);
    const chartData = Object.entries(count).map(([k,v]) => ({ category: k, value: v }));

    const root = am5.Root.new("facultyBarChart");
    root.setThemes([am5themes_Animated.new(root)]);
    const chart = root.container.children.push(am5xy.XYChart.new(root, {}));

    const xAxis = chart.xAxes.push(am5xy.CategoryAxis.new(root, { categoryField: "category", renderer: am5xy.AxisRendererX.new(root, {}) }));
    const yAxis = chart.yAxes.push(am5xy.ValueAxis.new(root, { renderer: am5xy.AxisRendererY.new(root, {}) }));
    const series = chart.series.push(am5xy.ColumnSeries.new(root, { xAxis, yAxis, categoryXField: "category", valueYField: "value" }));
    series.columns.template.setAll({ cornerRadiusTL: 8, cornerRadiusTR: 8 });

    xAxis.data.setAll(chartData);
    series.data.setAll(chartData);
    charts.faculty = root;
}

function createYearLineChart(filtered) {
    const count = {};
    filtered.forEach(d => count[d.Year] = (count[d.Year] || 0) + 1);
    const chartData = Object.entries(count).sort((a,b) => a[0]-b[0]).map(([y,v]) => ({ year:y, value:v }));

    const root = am5.Root.new("lineChart");
    root.setThemes([am5themes_Animated.new(root)]);
    const chart = root.container.children.push(am5xy.XYChart.new(root, {}));

    const xAxis = chart.xAxes.push(am5xy.CategoryAxis.new(root, { categoryField: "year", renderer: am5xy.AxisRendererX.new(root, {}) }));
    const yAxis = chart.yAxes.push(am5xy.ValueAxis.new(root, { renderer: am5xy.AxisRendererY.new(root, {}) }));
    const series = chart.series.push(am5xy.LineSeries.new(root, { xAxis, yAxis, categoryXField: "year", valueYField: "value", stroke: am5.color("#3498db"), strokeWidth:3 }));

    xAxis.data.setAll(chartData);
    series.data.setAll(chartData);
    charts.line = root;
}

function createMajorPieChart(filtered) {
    const count = {};
    filtered.forEach(d => {
        const major = d["Major"] || "Unknown";
        count[major] = (count[major] || 0) + 1;
    });
    const chartData = Object.entries(count).map(([k,v]) => ({ category: k, value: v }));

    const root = am5.Root.new("majorPieChart");
    root.setThemes([am5themes_Animated.new(root)]);
    const chart = root.container.children.push(am5percent.PieChart.new(root, {}));
    const series = chart.series.push(am5percent.PieSeries.new(root, { valueField:"value", categoryField:"category" }));
    series.slices.template.setAll({ cornerRadius: 6 });
    series.data.setAll(chartData);
    charts.pie = root;
}

function createSalaryBarChart(filtered) {
    const majorStats = {};
    filtered.forEach(d => {
        const major = d["Major"];
        const doctor = d["Major Doctor"];
        if (!major || !doctor) return;
        if (!majorStats[major]) majorStats[major] = { doctor: doctor, totalFees: 0 };
        majorStats[major].totalFees += (d.Fees || 0);
    });

    let chartData = Object.values(majorStats).map(stat => ({ doctor: stat.doctor, salary: Math.round(0.005*stat.totalFees) }));
    chartData = chartData.sort(() => Math.random()-0.5);

    const root = am5.Root.new("salaryChart");
    root.setThemes([am5themes_Animated.new(root)]);
    root.numberFormatter.set("numberFormat","#,###");
    const chart = root.container.children.push(am5xy.XYChart.new(root, { panX:false, panY:false }));

    const yAxis = chart.yAxes.push(am5xy.ValueAxis.new(root, { renderer: am5xy.AxisRendererY.new(root, {}) }));
    const xAxis = chart.xAxes.push(am5xy.CategoryAxis.new(root, { categoryField:"doctor", renderer: am5xy.AxisRendererX.new(root, { cellStartLocation:0.1, cellEndLocation:0.9 }) }));

    xAxis.get("renderer").labels.template.setAll({ rotation:-45, centerY:am5.p100, centerX:am5.p100, paddingRight:15 });
    const series = chart.series.push(am5xy.ColumnSeries.new(root, { xAxis, yAxis, valueYField:"salary", categoryXField:"doctor", fill: am5.color("#9b59b6") }));
    series.columns.template.setAll({ cornerRadiusTL:8, cornerRadiusTR:8, tooltipText:"{categoryX}: {valueY.formatNumber('#,###')} EGP" });

    series.data.setAll(chartData);
    xAxis.data.setAll(chartData);
    charts.salary = root;
}

yearFilter.onchange = updateDashboard;
facultyFilter.onchange = updateDashboard;
loadData();
</script>

</body>
</html>