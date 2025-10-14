// --------------- CONFIGURAZIONE ---------------
const sheetName = "APP";
const excelUrl = "https://raw.githubusercontent.com/felicedario/bilancio/main/bilanciocorrente.xlsm";

// --------------- FUNZIONI DI UTILITÀ ---------------
const formatCurrency = (value) => {
    const number = Number(value) || 0;
    return number.toLocaleString('it-IT', { style: 'currency', currency: 'EUR' });
};

const parseValue = (value) => {
    if (typeof value === 'number') return value;
    if (typeof value !== 'string' || value.trim() === '') return 0;
    const cleanedValue = value.replace(/€/g, '').trim().replace(/\./g, '').replace(/,/g, '.');
    const number = parseFloat(cleanedValue);
    return isNaN(number) ? 0 : number;
};

// --------------- ELEMENTI DELLA PAGINA ---------------
const dashboardTitle = document.getElementById('dashboard-title');
const monthButtonsContainer = document.getElementById('month-buttons-container');
const statusMessage = document.getElementById('status-message');
const dashboardGrid = document.getElementById('dashboard-grid');
const giacenzaValoreEl = document.getElementById('giacenza-valore');
const disponibilitaValoreEl = document.getElementById('disponibilita-valore');
const totalEntrateEl = document.getElementById('total-entrate');
const stipendioValoreEl = document.getElementById('stipendio-valore');
const altroValoreEl = document.getElementById('altro-valore');
const totalSpeseEl = document.getElementById('total-spese');
const necessitaValoreEl = document.getElementById('necessita-valore');
const svagoValoreEl = document.getElementById('svago-valore');
const rimborsareValoreEl = document.getElementById('rimborsare-valore');
const risparmiValoreEl = document.getElementById('risparmi-valore');
const investimentiValoreEl = document.getElementById('investimenti-valore');
const obiettiviSection = document.getElementById('obiettivi-section');
const necessitaPercentEl = document.getElementById('necessita-percent');
const necessitaSpesaCorrenteEl = document.getElementById('necessita-spesa-corrente');
const necessitaMaxEl = document.getElementById('necessita-max');
const necessitaMargineEl = document.getElementById('necessita-margine');
const svagoPercentEl = document.getElementById('svago-percent');
const svagoSpesaCorrenteEl = document.getElementById('svago-spesa-corrente');
const svagoMaxEl = document.getElementById('svago-max');
const svagoMargineEl = document.getElementById('svago-margine');
const risparmiPercentEl = document.getElementById('risparmi-percent');
const risparmiValoreCorrenteEl = document.getElementById('risparmi-valore-corrente');
const risparmiMinEl = document.getElementById('risparmi-min');
const risparmiMargineEl = document.getElementById('risparmi-margine');
const investimentiPercentEl = document.getElementById('investimenti-percent');
const investimentiValoreCorrenteEl = document.getElementById('investimenti-valore-corrente');
const investimentiMinEl = document.getElementById('investimenti-min');
const investimentiMargineEl = document.getElementById('investimenti-margine');
const portafoglioSection = document.getElementById('portafoglio-section');
const portfolioEntrateValoreEl = document.getElementById('portfolio-entrate-valore');

let necessitaChart, svagoChart, risparmiChart, investimentiChart;
let portfolioChart;

// --------------- LOGICA PRINCIPALE ---------------
async function loadExcelData() {
    try {
        statusMessage.textContent = 'Caricamento dati dal tuo repository...';
        const response = await fetch(excelUrl);
        if (!response.ok) throw new Error(`Errore di rete (status: ${response.status})`);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        processWorkbook(workbook);
    } catch (error) {
        console.error("Errore caricamento file Excel:", error);
        statusMessage.textContent = `Errore: Impossibile caricare i dati. Dettagli: ${error.message}`;
    }
}

function processWorkbook(workbook) {
    if (!workbook.SheetNames.includes(sheetName)) {
        throw new Error(`Foglio "${sheetName}" non trovato.`);
    }
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    dashboardTitle.textContent = `Dashboard Finanziaria ${jsonData[0]?.[0] || 'Mensile'}`;

    const mesi = [];
    const datiMensili = {};

    for (let i = 3; i < 15; i++) {
        const row = jsonData[i] || [];
        const mese = row[0];
        if (mese && typeof mese === 'string' && mese.trim() !== '') {
            mesi.push(mese);
            datiMensili[mese] = {
                stipendio: parseValue(row[1]), altro: parseValue(row[2]), entrataTotale: parseValue(row[3]),
                necessita: parseValue(row[5]), svago: parseValue(row[6]), daRimborsare: parseValue(row[7]),
                giacenza: parseValue(row[10]), disponibilita: parseValue(row[11]),
                risparmi: parseValue(row[13]), investimenti: parseValue(row[14]),
                necessitaMax: parseValue(row[16]), necessitaPercent: parseValue(row[17]), necessitaPercentResto: parseValue(row[18]), necessitaMargine: row[19],
                svagoMax: parseValue(row[21]), svagoPercent: parseValue(row[22]), svagoPercentResto: parseValue(row[23]), svagoMargine: row[24],
                risparmiMin: parseValue(row[26]), risparmiPercent: parseValue(row[27]), risparmiPercentResto: parseValue(row[28]), risparmiMargine: row[29],
                investimentiMin: parseValue(row[31]), investimentiPercent: parseValue(row[32]), investimentiPercentResto: parseValue(row[33]), investimentiMargine: row[34],
                necessitaPortafoglioPct: parseValue(row[36]), svagoPortafoglioPct: parseValue(row[37]), rimborsarePortafoglioPct: parseValue(row[38]),
                risparmiPortafoglioPct: parseValue(row[39]), investimentiPortafoglioPct: parseValue(row[40]), nonAllocatoPct: parseValue(row[41])
            };
        }
    }

    statusMessage.style.display = 'none';
    dashboardGrid.classList.remove('hidden');
    obiettiviSection.classList.remove('hidden');
    portafoglioSection.classList.remove('hidden');

    monthButtonsContainer.innerHTML = '';
    mesi.forEach((mese) => {
        const button = document.createElement('button');
        button.className = 'month-button';
        button.textContent = mese.toUpperCase();
        button.onclick = () => updateDashboard(mese, datiMensili);
        monthButtonsContainer.appendChild(button);
    });

    const initialMonth = mesi[new Date().getMonth()] || mesi[0];
    if (initialMonth) updateDashboard(initialMonth, datiMensili);
}

function createOrUpdateChart(chartInstance, context, data, color) {
    const chartData = { datasets: [{ data, backgroundColor: [color, '#f3f4f6'], borderColor: [color, '#f3f4f6'], borderWidth: 1, cutout: '80%' }] };
    if (!chartInstance) {
        return new Chart(context, { type: 'doughnut', data: chartData, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false }, tooltip: { enabled: false } } } });
    }
    chartInstance.data.datasets[0].data = data;
    chartInstance.update();
    return chartInstance;
}

function updateDashboard(mese, datiMensili) {
    const dati = datiMensili[mese];
    if (!dati) return;

    document.querySelectorAll('.month-button').forEach(btn => btn.classList.toggle('active', btn.textContent.toLowerCase() === mese.toLowerCase()));
    
    giacenzaValoreEl.textContent = formatCurrency(dati.giacenza);
    disponibilitaValoreEl.textContent = formatCurrency(dati.disponibilita);
    totalEntrateEl.textContent = formatCurrency(dati.stipendio + dati.altro);
    stipendioValoreEl.textContent = formatCurrency(dati.stipendio);
    altroValoreEl.textContent = formatCurrency(dati.altro);
    totalSpeseEl.textContent = formatCurrency(dati.necessita + dati.svago + dati.daRimborsare);
    necessitaValoreEl.textContent = formatCurrency(dati.necessita);
    svagoValoreEl.textContent = formatCurrency(dati.svago);
    rimborsareValoreEl.textContent = formatCurrency(dati.daRimborsare);
    risparmiValoreEl.textContent = formatCurrency(dati.risparmi);
    investimentiValoreEl.textContent = formatCurrency(dati.investimenti);

    necessitaPercentEl.textContent = `${Math.round(dati.necessitaPercent * 100)}%`;
    necessitaSpesaCorrenteEl.textContent = formatCurrency(dati.necessita);
    necessitaMaxEl.textContent = formatCurrency(dati.necessitaMax);
    necessitaMargineEl.textContent = typeof dati.necessitaMargine === 'number' ? formatCurrency(dati.necessitaMargine) : '--';
    necessitaChart = createOrUpdateChart(necessitaChart, document.getElementById('necessita-chart').getContext('2d'), [dati.necessitaPercent, dati.necessitaPercentResto], '#ef4444');

    svagoPercentEl.textContent = `${Math.round(dati.svagoPercent * 100)}%`;
    svagoSpesaCorrenteEl.textContent = formatCurrency(dati.svago);
    svagoMaxEl.textContent = formatCurrency(dati.svagoMax);
    svagoMargineEl.textContent = typeof dati.svagoMargine === 'number' ? formatCurrency(dati.svagoMargine) : '--';
    svagoChart = createOrUpdateChart(svagoChart, document.getElementById('svago-chart').getContext('2d'), [dati.svagoPercent, dati.svagoPercentResto], '#ef4444'); // Colore cambiato in ROSSO

    risparmiPercentEl.textContent = `${Math.round(dati.risparmiPercent * 100)}%`;
    risparmiValoreCorrenteEl.textContent = formatCurrency(dati.risparmi);
    risparmiMinEl.textContent = formatCurrency(dati.risparmiMin);
    risparmiMargineEl.textContent = typeof dati.risparmiMargine === 'number' ? formatCurrency(dati.risparmiMargine) : '--';
    risparmiChart = createOrUpdateChart(risparmiChart, document.getElementById('risparmi-chart').getContext('2d'), [dati.risparmiPercent, dati.risparmiPercentResto], '#f97316');

    investimentiPercentEl.textContent = `${Math.round(dati.investimentiPercent * 100)}%`;
    investimentiValoreCorrenteEl.textContent = formatCurrency(dati.investimenti);
    investimentiMinEl.textContent = formatCurrency(dati.investimentiMin);
    investimentiMargineEl.textContent = typeof dati.investimentiMargine === 'number' ? formatCurrency(dati.investimentiMargine) : '--';
    investimentiChart = createOrUpdateChart(investimentiChart, document.getElementById('investimenti-chart').getContext('2d'), [dati.investimentiPercent, dati.investimentiPercentResto], '#a855f7'); // Colore cambiato in VIOLA

    portfolioEntrateValoreEl.textContent = formatCurrency(dati.entrataTotale);
    const portfolioCtx = document.getElementById('portfolio-chart').getContext('2d');
    const portfolioData = {
        labels: ['Necessità', 'Svago', 'Da Rimborsare', 'Risparmi', 'Investimenti', 'Non Allocato'],
        datasets: [{
            data: [dati.necessitaPortafoglioPct, dati.svagoPortafoglioPct, dati.
