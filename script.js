// --- CONFIGURAZIONE ---
const sheetName = "APP";
const excelUrl = "https://raw.githubusercontent.com/felicedario/bilancio/main/bilanciocorrente.xlsm";

// --- FUNZIONI DI UTILITÀ ---
const formatCurrency = (value) => (Number(value) || 0).toLocaleString('it-IT', { style: 'currency', currency: 'EUR' });
const parseValue = (value) => {
    if (typeof value === 'number') return value;
    if (typeof value !== 'string' || value.trim() === '') return 0;
    const cleaned = value.replace(/€/g, '').trim().replace(/\./g, '').replace(/,/g, '.');
    const number = parseFloat(cleaned);
    return isNaN(number) ? 0 : number;
};

// --- ELEMENTI DELLA PAGINA ---
const el = (id) => document.getElementById(id);
const dashboardTitle = el('dashboard-title'), monthButtonsContainer = el('month-buttons-container'), statusMessage = el('status-message'), dashboardGrid = el('dashboard-grid');
const giacenzaValoreEl = el('giacenza-valore'), disponibilitaValoreEl = el('disponibilita-valore');
const totalEntrateEl = el('total-entrate'), stipendioValoreEl = el('stipendio-valore'), altroValoreEl = el('altro-valore');
const totalSpeseEl = el('total-spese'), necessitaValoreEl = el('necessita-valore'), svagoValoreEl = el('svago-valore'), rimborsareValoreEl = el('rimborsare-valore');
const risparmiValoreEl = el('risparmi-valore'), investimentiValoreEl = el('investimenti-valore');
// Elementi Portafoglio
const portafoglioSection = el('portafoglio-section'), portafoglioEntrateEl = el('portafoglio-entrate');
// Elementi Obiettivi
const obiettiviSection = el('obiettivi-section');
const necessitaPercentEl = el('necessita-percent'), necessitaSpesaCorrenteEl = el('necessita-spesa-corrente'), necessitaMaxEl = el('necessita-max'), necessitaMargineEl = el('necessita-margine');
const svagoPercentEl = el('svago-percent'), svagoSpesaCorrenteEl = el('svago-spesa-corrente'), svagoMaxEl = el('svago-max'), svagoMargineEl = el('svago-margine');
const risparmiPercentEl = el('risparmi-percent'), risparmiValoreCorrenteEl = el('risparmi-valore-corrente'), risparmiMinEl = el('risparmi-min'), risparmiMargineEl = el('risparmi-margine');
const investimentiPercentEl = el('investimenti-percent'), investimentiValoreCorrenteEl = el('investimenti-valore-corrente'), investimentiMinEl = el('investimenti-min'), investimentiMargineEl = el('investimenti-margine');

// --- ISTANZE GRAFICI ---
let charts = {};

// --- LOGICA PRINCIPALE ---
document.addEventListener('DOMContentLoaded', async () => {
    try {
        statusMessage.textContent = 'Caricamento dati dal tuo repository...';
        const response = await fetch(excelUrl);
        if (!response.ok) throw new Error(`Errore di rete (status: ${response.status})`);
        const arrayBuffer = await response.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, { type: 'array' });
        processWorkbook(workbook);
    } catch (error)
        console.error("Errore critico:", error);
        statusMessage.textContent = `Errore: Impossibile caricare i dati. ${error.message}`;
    }
});

function processWorkbook(workbook) {
    if (!workbook.SheetNames.includes(sheetName)) throw new Error(`Foglio "${sheetName}" non trovato.`);
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    dashboardTitle.textContent = `Dashboard Finanziaria ${jsonData[0]?.[0] || ''}`;
    const mesi = [], datiMensili = {};

    for (let i = 3; i < 15; i++) {
        const row = jsonData[i] || [];
        const mese = row[0];
        if (mese && typeof mese === 'string' && mese.trim() !== '') {
            mesi.push(mese);
            datiMensili[mese] = {
                // Dati principali
                stipendio: parseValue(row[1]), altro: parseValue(row[2]), totaleEntrate: parseValue(row[3]),
                necessita: parseValue(row[5]), svago: parseValue(row[6]), daRimborsare: parseValue(row[7]),
                giacenza: parseValue(row[10]), disponibilita: parseValue(row[11]),
                risparmi: parseValue(row[13]), investimenti: parseValue(row[14]),
                // Dati Obiettivi
                necessitaMax: parseValue(row[16]), necessitaPercent: parseValue(row[17]), necessitaPercentResto: parseValue(row[18]), necessitaMargine: row[19],
                svagoMax: parseValue(row[21]), svagoPercent: parseValue(row[22]), svagoPercentResto: parseValue(row[23]), svagoMargine: row[24],
                risparmiMin: parseValue(row[26]), risparmiPercent: parseValue(row[27]), risparmiPercentResto: parseValue(row[28]), risparmiMargine: row[29],
                investimentiMin: parseValue(row[31]), investimentiPercent: parseValue(row[32]), investimentiPercentResto: parseValue(row[33]), investimentiMargine: row[34],
                // Dati Portafoglio (Colonne AK-AP -> indici 36-41)
                portafoglioNecessita: parseValue(row[36]),
                portafoglioSvago: parseValue(row[37]),
                portafoglioRimborsare: parseValue(row[38]),
                portafoglioRisparmi: parseValue(row[39]),
                portafoglioInvestimenti: parseValue(row[40]),
                portafoglioResto: parseValue(row[41]),
            };
        }
    }

    statusMessage.style.display = 'none';
    dashboardGrid.classList.remove('hidden');
    portafoglioSection.classList.remove('hidden');
    obiettiviSection.classList.remove('hidden');

    monthButtonsContainer.innerHTML = '';
    mesi.forEach(mese => {
        const button = document.createElement('button');
        button.className = 'month-button';
        button.textContent = mese.toUpperCase();
        button.onclick = () => updateDashboard(mese, datiMensili);
        monthButtonsContainer.appendChild(button);
    });

    const initialMonth = mesi[new Date().getMonth()] || mesi[0];
    if (initialMonth) updateDashboard(initialMonth, datiMensili);
}

function createOrUpdateChart(chartId, contextId, data, colors, chartType = 'doughnut') {
    const chartData = {
        datasets: [{
            data: data,
            backgroundColor: colors,
            borderColor: chartType === 'doughnut' ? colors : '#fff', // Bordo bianco per la legenda del portafoglio
            borderWidth: chartType === 'doughnut' ? 1 : 2,
            cutout: '70%'
        }],
        labels: ['Necessità', 'Svago', 'Da Rimborsare', 'Risparmi', 'Investimenti', 'Non allocato']
    };

    if (charts[chartId]) {
        charts[chartId].data = chartData;
        charts[chartId].update();
    } else {
        charts[chartId] = new Chart(el(contextId).getContext('2d'), {
            type: chartType,
            data: chartData,
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: chartId === 'portafoglio', // Mostra legenda solo per portafoglio
                        position: 'right',
                    },
                    tooltip: { enabled: true }
                }
            }
        });
    }
}

function updateDashboard(mese, datiMensili) {
    const data = datiMensili[mese];
    if (!data) return;

    document.querySelectorAll('.month-button').forEach(btn => btn.classList.toggle('active', btn.textContent.toLowerCase() === mese.toLowerCase()));
    
    // Cards Principali
    totalEntrateEl.textContent = formatCurrency(data.totaleEntrate);
    giacenzaValoreEl.textContent = formatCurrency(data.giacenza);
    disponibilitaValoreEl.textContent = formatCurrency(data.disponibilita);
    stipendioValoreEl.textContent = formatCurrency(data.stipendio);
    altroValoreEl.textContent = formatCurrency(data.altro);
    totalSpeseEl.textContent = formatCurrency(data.necessita + data.svago + data.daRimborsare);
    necessitaValoreEl.textContent = formatCurrency(data.necessita);
    svagoValoreEl.textContent = formatCurrency(data.svago);
    rimborsareValoreEl.textContent = formatCurrency(data.daRimborsare);
    risparmiValoreEl.textContent = formatCurrency(data.risparmi);
    investimentiValoreEl.textContent = formatCurrency(data.investimenti);

    // Sezione Portafoglio
    portafoglioEntrateEl.textContent = formatCurrency(data.totaleEntrate);
    createOrUpdateChart('portafoglio', 'portafoglio-chart',
        [data.portafoglioNecessita, data.portafoglioSvago, data.portafoglioRimborsare, data.portafoglioRisparmi, data.portafoglioInvestimenti, data.portafoglioResto],
        ['#d92d20', '#3b82f6', '#6D4C41', '#f97316', '#424242', '#e0e0e0'],
        'pie' // Usiamo un tipo 'pie' per vedere meglio la legenda
    );

    // Sezione Obiettivi
    const formatPercent = (val) => `${Math.round(val * 100)}%`;
    const formatMargine = (val) => typeof val === 'number' ? formatCurrency(val) : '--';
    const createDoughnut = (id, data, colors) => createOrUpdateChart(id, `${id}-chart`, data, colors, 'doughnut');

    necessitaPercentEl.textContent = formatPercent(data.necessitaPercent);
    necessitaSpesaCorrenteEl.textContent = formatCurrency(data.necessita);
    necessitaMaxEl.textContent = formatCurrency(data.necessitaMax);
    necessitaMargineEl.textContent = formatMargine(data.necessitaMargine);
    createDoughnut('necessita', [data.necessitaPercent, data.necessitaPercentResto], ['#d92d20', '#f3f4f6']);

    svagoPercentEl.textContent = formatPercent(data.svagoPercent);
    svagoSpesaCorrenteEl.textContent = formatCurrency(data.svago);
    svagoMaxEl.textContent = formatCurrency(data.svagoMax);
    svagoMargineEl.textContent = formatMargine(data.svagoMargine);
    createDoughnut('svago', [data.svagoPercent, data.svagoPercentResto], ['#3b82f6', '#f3f4f6']);

    risparmiPercentEl.textContent = formatPercent(data.risparmiPercent);
    risparmiValoreCorrenteEl.textContent = formatCurrency(data.risparmi);
    risparmiMinEl.textContent = formatCurrency(data.risparmiMin);
    risparmiMargineEl.textContent = formatMargine(data.risparmiMargine);
    createDoughnut('risparmi', [data.risparmiPercent, data.risparmiPercentResto], ['#f97316', '#f3f4f6']);

    investimentiPercentEl.textContent = formatPercent(data.investimentiPercent);
    investimentiValoreCorrenteEl.textContent = formatCurrency(data.investimenti);
    investimentiMinEl.textContent = formatCurrency(data.investimentiMin);
    investimentiMargineEl.textContent = formatMargine(data.investimentiMargine);
    createDoughnut('investimenti', [data.investimentiPercent, data.investimentiPercentResto], ['#f97316', '#f3f4f6']);
}
