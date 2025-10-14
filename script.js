// --------------- CONFIGURAZIONE ---------------
const sheetName = "APP";

// URL diretto al file "Raw" sul tuo repository GitHub "bilancio"
const excelUrl = "https://raw.githubusercontent.com/felicedario/bilancio/main/bilanciocorrente.xlsm";

// --------------- FUNZIONI DI UTILITÀ ---------------

// Funzione per formattare i numeri come valuta
const formatCurrency = (value) => {
    const number = Number(value) || 0;
    return number.toLocaleString('it-IT', { style: 'currency', currency: 'EUR' });
};

// Funzione per "pulire" e convertire i valori letti in numeri
const parseValue = (value) => {
    if (typeof value === 'number') return value;
    if (typeof value !== 'string' || value.trim() === '') return 0;
    
    const cleanedValue = value
        .replace(/€/g, '')
        .trim()
        .replace(/\./g, '')
        .replace(/,/g, '.');

    const number = parseFloat(cleanedValue);
    return isNaN(number) ? 0 : number;
};

// --------------- ELEMENTI DELLA PAGINA ---------------
const dashboardTitle = document.getElementById('dashboard-title');
const monthButtonsContainer = document.getElementById('month-buttons-container');
const statusMessage = document.getElementById('status-message');
const dashboardGrid = document.getElementById('dashboard-grid');

// Card Giacenza
const giacenzaValoreEl = document.getElementById('giacenza-valore');
const disponibilitaValoreEl = document.getElementById('disponibilita-valore');

// Card Entrate
const totalEntrateEl = document.getElementById('total-entrate');
const stipendioValoreEl = document.getElementById('stipendio-valore');
const altroValoreEl = document.getElementById('altro-valore');

// Card Spese
const totalSpeseEl = document.getElementById('total-spese');
const necessitaValoreEl = document.getElementById('necessita-valore');
const svagoValoreEl = document.getElementById('svago-valore');
const rimborsareValoreEl = document.getElementById('rimborsare-valore');

// Card Risparmi e Investimenti
const risparmiValoreEl = document.getElementById('risparmi-valore');
const investimentiValoreEl = document.getElementById('investimenti-valore');

// Sezione Obiettivi
const obiettiviSection = document.getElementById('obiettivi-section');
// Obiettivo Necessità
const necessitaPercentEl = document.getElementById('necessita-percent');
const necessitaSpesaCorrenteEl = document.getElementById('necessita-spesa-corrente');
const necessitaMaxEl = document.getElementById('necessita-max');
const necessitaMargineEl = document.getElementById('necessita-margine');
// Obiettivo Svago
const svagoPercentEl = document.getElementById('svago-percent');
const svagoSpesaCorrenteEl = document.getElementById('svago-spesa-corrente');
const svagoMaxEl = document.getElementById('svago-max');
const svagoMargineEl = document.getElementById('svago-margine');
// Obiettivo Risparmi
const risparmiPercentEl = document.getElementById('risparmi-percent');
const risparmiValoreCorrenteEl = document.getElementById('risparmi-valore-corrente');
const risparmiMinEl = document.getElementById('risparmi-min');
const risparmiMargineEl = document.getElementById('risparmi-margine');
// Obiettivo Investimenti
const investimentiPercentEl = document.getElementById('investimenti-percent');
const investimentiValoreCorrenteEl = document.getElementById('investimenti-valore-corrente');
const investimentiMinEl = document.getElementById('investimenti-min');
const investimentiMargineEl = document.getElementById('investimenti-margine');

// Variabili per le istanze dei grafici
let necessitaChart, svagoChart, risparmiChart, investimentiChart;

// --------------- LOGICA PRINCIPALE ---------------

async function loadExcelData() {
    try {
        statusMessage.textContent = 'Caricamento dati dal tuo repository...';
        const response = await fetch(excelUrl);
        if (!response.ok) {
            throw new Error(`Errore di rete nel scaricare il file (status: ${response.status})`);
        }
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        processWorkbook(workbook);
    } catch (error) {
        console.error("Errore nel caricamento o processamento del file Excel:", error);
        statusMessage.textContent = `Errore: Impossibile caricare i dati. Dettagli: ${error.message}`;
    }
}

function processWorkbook(workbook) {
    if (!workbook.SheetNames.includes(sheetName)) {
        throw new Error(`Foglio di lavoro "${sheetName}" non trovato nel file.`);
    }
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    const anno = jsonData[0]?.[0] || 'Mensile';
    dashboardTitle.textContent = `Dashboard Finanziaria ${anno}`;

    const mesi = [];
    const datiMensili = {};

    for (let i = 3; i < 15; i++) {
        const row = jsonData[i] || [];
        const mese = row[0];
        if (mese && typeof mese === 'string' && mese.trim() !== '') {
            mesi.push(mese);
            datiMensili[mese] = {
                // Cards Principali
                stipendio: parseValue(row[1]),
                altro: parseValue(row[2]),
                necessita: parseValue(row[5]),
                svago: parseValue(row[6]),
                daRimborsare: parseValue(row[7]),
                giacenza: parseValue(row[10]),
                disponibilita: parseValue(row[11]),
                risparmi: parseValue(row[13]),
                investimenti: parseValue(row[14]),
                // Obiettivo Necessità (Q-T)
                necessitaMax: parseValue(row[16]),
                necessitaPercent: parseValue(row[17]),
                necessitaPercentResto: parseValue(row[18]),
                necessitaMargine: row[19],
                // Obiettivo Svago (V-Y)
                svagoMax: parseValue(row[21]),
                svagoPercent: parseValue(row[22]),
                svagoPercentResto: parseValue(row[23]),
                svagoMargine: row[24],
                // Obiettivo Risparmi (AA-AD)
                risparmiMin: parseValue(row[26]),
                risparmiPercent: parseValue(row[27]),
                risparmiPercentResto: parseValue(row[28]),
                risparmiMargine: row[29],
                // Obiettivo Investimenti (AF-AI)
                investimentiMin: parseValue(row[31]),
                investimentiPercent: parseValue(row[32]),
                investimentiPercentResto: parseValue(row[33]),
                investimentiMargine: row[34]
            };
        }
    }

    statusMessage.style.display = 'none';
    dashboardGrid.classList.remove('hidden');
    obiettiviSection.classList.remove('hidden');

    monthButtonsContainer.innerHTML = '';
    mesi.forEach((mese) => {
        const button = document.createElement('button');
        button.className = 'month-button';
        button.textContent = mese.toUpperCase();
        button.onclick = () => updateDashboard(mese, datiMensili);
        monthButtonsContainer.appendChild(button);
    });

    const currentMonthIndex = new Date().getMonth();
    const initialMonth = mesi[currentMonthIndex] || mesi[0];
    if (initialMonth) {
        updateDashboard(initialMonth, datiMensili);
    }
}

// Funzione generica per creare o aggiornare un grafico
function createOrUpdateChart(chartInstance, context, data, color) {
    const chartData = {
        datasets: [{
            data: data,
            backgroundColor: [color, '#f3f4f6'],
            borderColor: [color, '#f3f4f6'],
            borderWidth: 1,
            cutout: '80%'
        }]
    };

    if (!chartInstance) {
        return new Chart(context, {
            type: 'doughnut',
            data: chartData,
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { enabled: false }
                }
            }
        });
    } else {
        chartInstance.data.datasets[0].data = data;
        chartInstance.update();
        return chartInstance;
    }
}

function updateDashboard(mese, datiMensili) {
    const datiDelMese = datiMensili[mese];
    if (!datiDelMese) return;

    document.querySelectorAll('.month-button').forEach(btn => {
        btn.classList.toggle('active', btn.textContent.toLowerCase() === mese.toLowerCase());
    });
    
    // Aggiornamento Cards Principali
    giacenzaValoreEl.textContent = formatCurrency(datiDelMese.giacenza);
    disponibilitaValoreEl.textContent = formatCurrency(datiDelMese.disponibilita);
    const totaleEntrate = datiDelMese.stipendio + datiDelMese.altro;
    totalEntrateEl.textContent = formatCurrency(totaleEntrate);
    stipendioValoreEl.textContent = formatCurrency(datiDelMese.stipendio);
    altroValoreEl.textContent = formatCurrency(datiDelMese.altro);
    const totaleSpese = datiDelMese.necessita + datiDelMese.svago + datiDelMese.daRimborsare;
    totalSpeseEl.textContent = formatCurrency(totaleSpese);
    necessitaValoreEl.textContent = formatCurrency(datiDelMese.necessita);
    svagoValoreEl.textContent = formatCurrency(datiDelMese.svago);
    rimborsareValoreEl.textContent = formatCurrency(datiDelMese.daRimborsare);
    risparmiValoreEl.textContent = formatCurrency(datiDelMese.risparmi);
    investimentiValoreEl.textContent = formatCurrency(datiDelMese.investimenti);

    // --- AGGIORNAMENTO SEZIONE OBIETTIVI ---

    // 1. Obiettivo Necessità
    necessitaPercentEl.textContent = `${Math.round(datiDelMese.necessitaPercent * 100)}%`;
    necessitaSpesaCorrenteEl.textContent = formatCurrency(datiDelMese.necessita);
    necessitaMaxEl.textContent = formatCurrency(datiDelMese.necessitaMax);
    necessitaMargineEl.textContent = typeof datiDelMese.necessitaMargine === 'number' ? formatCurrency(datiDelMese.necessitaMargine) : '--';
    const necessitaCtx = document.getElementById('necessita-chart').getContext('2d');
    necessitaChart = createOrUpdateChart(necessitaChart, necessitaCtx, [datiDelMese.necessitaPercent, datiDelMese.necessitaPercentResto], '#d92d20');

    // 2. Obiettivo Svago
    svagoPercentEl.textContent = `${Math.round(datiDelMese.svagoPercent * 100)}%`;
    svagoSpesaCorrenteEl.textContent = formatCurrency(datiDelMese.svago);
    svagoMaxEl.textContent = formatCurrency(datiDelMese.svagoMax);
    svagoMargineEl.textContent = typeof datiDelMese.svagoMargine === 'number' ? formatCurrency(datiDelMese.svagoMargine) : '--';
    const svagoCtx = document.getElementById('svago-chart').getContext('2d');
    svagoChart = createOrUpdateChart(svagoChart, svagoCtx, [datiDelMese.svagoPercent, datiDelMese.svagoPercentResto], '#3b82f6');

    // 3. Obiettivo Risparmi
    risparmiPercentEl.textContent = `${Math.round(datiDelMese.risparmiPercent * 100)}%`;
    risparmiValoreCorrenteEl.textContent = formatCurrency(datiDelMese.risparmi);
    risparmiMinEl.textContent = formatCurrency(datiDelMese.risparmiMin);
    risparmiMargineEl.textContent = typeof datiDelMese.risparmiMargine === 'number' ? formatCurrency(datiDelMese.risparmiMargine) : '--';
    const risparmiCtx = document.getElementById('risparmi-chart').getContext('2d');
    risparmiChart = createOrUpdateChart(risparmiChart, risparmiCtx, [datiDelMese.risparmiPercent, datiDelMese.risparmiPercentResto], '#f97316');

    // 4. Obiettivo Investimenti
    investimentiPercentEl.textContent = `${Math.round(datiDelMese.investimentiPercent * 100)}%`;
    investimentiValoreCorrenteEl.textContent = formatCurrency(datiDelMese.investimenti);
    investimentiMinEl.textContent = formatCurrency(datiDelMese.investimentiMin);
    investimentiMargineEl.textContent = typeof datiDelMese.investimentiMargine === 'number' ? formatCurrency(datiDelMese.investimentiMargine) : '--';
    const investimentiCtx = document.getElementById('investimenti-chart').getContext('2d');
    investimentiChart = createOrUpdateChart(investimentiChart, investimentiCtx, [datiDelMese.investimentiPercent, datiDelMese.investimentiPercentResto], '#f97316');
}

document.addEventListener('DOMContentLoaded', loadExcelData);
