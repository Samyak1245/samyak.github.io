// ---------- Background starfield ----------
const canvas = document.getElementById('stars');
const ctx = canvas.getContext('2d');
let stars = [];

function resize(){
  canvas.width = window.innerWidth;
  canvas.height = document.documentElement.scrollHeight;
  const count = Math.floor((canvas.width * canvas.height) / 9000);
  stars = Array.from({length: count}, () => ({
    x: Math.random() * canvas.width,
    y: Math.random() * canvas.height,
    r: Math.random() * 1.3 + 0.2,
    a: Math.random() * 0.6 + 0.3,
    tw: Math.random() * 0.02 + 0.005
  }));
}

function draw(){
  ctx.clearRect(0,0,canvas.width,canvas.height);
  ctx.fillStyle = '#e6e9f0';
  for(const s of stars){
    s.a += s.tw;
    const alpha = 0.4 + Math.abs(Math.sin(s.a)) * 0.6;
    ctx.globalAlpha = alpha;
    ctx.beginPath();
    ctx.arc(s.x, s.y, s.r, 0, Math.PI*2);
    ctx.fill();
  }
  ctx.globalAlpha = 1;
  requestAnimationFrame(draw);
}

window.addEventListener('resize', resize);
resize();
draw();

// ---------- Mobile nav toggle ----------
const navToggle = document.getElementById('navToggle');
const navList = document.getElementById('navList');
navToggle.addEventListener('click', () => navList.classList.toggle('open'));
navList.querySelectorAll('a').forEach(a => a.addEventListener('click', () => navList.classList.remove('open')));

// ---------- Data Playground dashboard ----------
document.addEventListener('DOMContentLoaded', function () {
  let activeData = [];
  let customChartInstance = null;

  // Pool of possible columns to pick from — each run selects a random subset,
  // so both the values AND the column set change every click.
  const textPool = [
    { name: "Region", values: ["North", "South", "East", "West", "Central"] },
    { name: "Department", values: ["Sales", "Marketing", "Engineering", "Support", "HR"] },
    { name: "Product", values: ["Alpha", "Nova", "Zenith", "Pulse", "Orbit"] },
    { name: "Category", values: ["Electronics", "Furniture", "Technology", "Apparel", "Groceries"] },
    { name: "Country", values: ["India", "USA", "Germany", "Japan", "Brazil"] },
    { name: "Channel", values: ["Online", "Retail", "Wholesale", "Partner"] }
  ];
  const numberPool = [
    { name: "Sales", min: 2000, max: 32000, decimals: 0 },
    { name: "Profit", min: 500, max: 9000, decimals: 0 },
    { name: "Units", min: 5, max: 500, decimals: 0 },
    { name: "Rating", min: 1, max: 5, decimals: 1 },
    { name: "Cost", min: 1000, max: 20000, decimals: 0 },
    { name: "Headcount", min: 2, max: 80, decimals: 0 }
  ];

  function shuffle(arr) {
    const a = [...arr];
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [a[i], a[j]] = [a[j], a[i]];
    }
    return a;
  }

  function randomInRange(min, max, decimals) {
    const val = min + Math.random() * (max - min);
    return decimals > 0 ? parseFloat(val.toFixed(decimals)) : Math.round(val);
  }

  function generateSampleData() {
    // Pick 1–2 text columns and 2–3 number columns at random each time
    const chosenText = shuffle(textPool).slice(0, 1 + Math.floor(Math.random() * 2));
    const chosenNumbers = shuffle(numberPool).slice(0, 2 + Math.floor(Math.random() * 2));

    const rowCount = 30 + Math.floor(Math.random() * 91);
    const data = [];
    for (let i = 0; i < rowCount; i++) {
      const row = {};
      chosenText.forEach(col => {
        row[col.name] = col.values[Math.floor(Math.random() * col.values.length)];
      });
      chosenNumbers.forEach(col => {
        row[col.name] = randomInRange(col.min, col.max, col.decimals);
      });
      data.push(row);
    }
    return data;
  }

  const fileInput = document.getElementById('excel-file-input');
  const fileNameDisplay = document.getElementById('file-name');
  const xAxisSelect = document.getElementById('x-axis-select');
  const yAxisSelect = document.getElementById('y-axis-select');
  const aggSelect = document.getElementById('agg-select');
  const chartTypeSelect = document.getElementById('chart-type-select');
  const controlsCard = document.getElementById('controls-card');
  const chartsWrapper = document.getElementById('charts-wrapper');
  const exportBtn = document.getElementById('export-chart-btn');

  if (fileInput) {
    fileInput.addEventListener('change', function (e) {
      const file = e.target.files[0];

      if (file && fileNameDisplay) {
        fileNameDisplay.innerText = file.name;
        fileNameDisplay.style.color = 'var(--accent)';
      } else if (fileNameDisplay) {
        fileNameDisplay.innerText = 'No file chosen';
        fileNameDisplay.style.color = 'var(--muted)';
        return;
      }

      if (!file) return;

      const reader = new FileReader();
      reader.onload = function (evt) {
        try {
          const data = new Uint8Array(evt.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet);

          if (!jsonData || jsonData.length === 0) {
            alert("The uploaded file contains no data.");
            return;
          }
          initializeDashboard(jsonData);
        } catch (err) {
          alert("Error parsing file. Make sure it is a valid CSV or Excel file.");
        }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  document.getElementById('load-sample-btn').addEventListener('click', function () {
    if (fileInput) fileInput.value = "";
    if (fileNameDisplay) {
      fileNameDisplay.innerText = 'Sample data loaded';
      fileNameDisplay.style.color = 'var(--accent)';
    }
    initializeDashboard(generateSampleData());
  });

  [xAxisSelect, yAxisSelect, aggSelect, chartTypeSelect].forEach(element => {
    element.addEventListener('change', updateChart);
  });

  exportBtn.addEventListener('click', function () {
    const chartCanvas = document.getElementById('customChart');
    const imageURI = chartCanvas.toDataURL('image/png');
    const link = document.createElement('a');
    link.download = 'data-chart.png';
    link.href = imageURI;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  });

  function initializeDashboard(data) {
    activeData = data;
    const columns = Object.keys(data[0]);

    xAxisSelect.innerHTML = "";
    yAxisSelect.innerHTML = "";

    columns.forEach(col => {
      const optionX = new Option(col, col);
      const optionY = new Option(col, col);
      xAxisSelect.add(optionX);
      yAxisSelect.add(optionY);
    });

    const numericCols = columns.filter(col => typeof data[0][col] === 'number');
    const textCols = columns.filter(col => typeof data[0][col] === 'string');

    if (textCols.length > 0) xAxisSelect.value = textCols[0];
    if (numericCols.length > 0) yAxisSelect.value = numericCols[0];

    controlsCard.style.display = "grid";
    chartsWrapper.style.display = "grid";
    exportBtn.style.display = "inline-flex";

    renderKPIs();
    renderTable();
    updateChart();
  }

  function renderKPIs() {
    const columns = Object.keys(activeData[0]);
    const kpiContainer = document.getElementById('kpi-container');
    kpiContainer.innerHTML = `
      <div class="kpi-card"><label>Total Rows</label><span>${activeData.length}</span></div>
      <div class="kpi-card"><label>Total Columns</label><span>${columns.length}</span></div>
    `;
  }

  function renderTable() {
    const columns = Object.keys(activeData[0]);
    const tableCard = document.getElementById('table-card');
    const tableHead = document.getElementById('table-head');
    const tableBody = document.getElementById('table-body');

    tableHead.innerHTML = columns.map(col => `<th>${col}</th>`).join('');
    tableBody.innerHTML = activeData.slice(0, 10).map(row => {
      return `<tr>${columns.map(col => `<td>${row[col] !== undefined ? row[col] : ''}</td>`).join('')}</tr>`;
    }).join('');

    tableCard.style.display = "block";
  }

  function updateChart() {
    if (!activeData || activeData.length === 0) return;

    const xCol = xAxisSelect.value;
    const yCol = yAxisSelect.value;
    const agg = aggSelect.value;
    const chartType = chartTypeSelect.value;

    const groupedData = {};
    activeData.forEach(row => {
      const key = row[xCol] !== undefined ? row[xCol] : 'Null';
      const val = parseFloat(row[yCol]) || 0;

      if (!groupedData[key]) groupedData[key] = [];
      groupedData[key].push(val);
    });

    const labels = Object.keys(groupedData);
    const aggregatedValues = labels.map(key => {
      const arr = groupedData[key];
      if (agg === 'SUM') return arr.reduce((a, b) => a + b, 0);
      if (agg === 'AVG') return arr.reduce((a, b) => a + b, 0) / arr.length;
      if (agg === 'COUNT') return arr.length;
      if (agg === 'MAX') return Math.max(...arr);
      if (agg === 'MIN') return Math.min(...arr);
    });

    document.getElementById('chart-title').innerText = `${agg} of ${yCol} by ${xCol}`;

    if (customChartInstance) customChartInstance.destroy();

    const chartCanvas = document.getElementById('customChart');
    // Theme-matched palette (cyan / gold + muted pastel variants — no neon greens/purples)
    const colors = ['#7ec8e3', '#d4af6a', '#a9b8e0', '#c9a9c2', '#9bc9c0', '#c9b89a'];

    customChartInstance = new Chart(chartCanvas.getContext('2d'), {
      type: chartType,
      data: {
        labels: labels,
        datasets: [{
          label: `${agg} of ${yCol}`,
          data: aggregatedValues,
          backgroundColor: chartType === 'line' ? 'rgba(110, 231, 249, 0.2)' : colors,
          borderColor: '#7ec8e3',
          borderWidth: 1.5,
          fill: chartType === 'line'
        }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { labels: { color: '#e6e9f0' } }
        },
        scales: (chartType === 'pie' || chartType === 'doughnut') ? {} : {
          x: { ticks: { color: '#8b93a7' }, grid: { color: 'rgba(255,255,255,0.05)' } },
          y: { ticks: { color: '#8b93a7' }, grid: { color: 'rgba(255,255,255,0.05)' } }
        }
      }
    });
  }
});