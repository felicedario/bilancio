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

// Card Entrate
const totalEntrateEl = document.getElementById('total-entrate');
const stipendioValoreEl = document.getElementById('stipendio-valore');
const altroValoreEl = document.getElementById('altro-valore');

// Card Spese
const totalSpeseEl = document.getElementById('total-spese');
const necessitaValoreEl = document.getElementById('necessita-valore');
const svagoValoreEl = document.getElementById('svago-valore');
const rimborsareValoreEl = document.getElementById('rimborsare-valore');
const giacenzaValoreEl = document.getElementById('giacenza-valore');
const disponibilitaValoreEl = document.getElementById('disponibilita-valore');
// --------------- LOGICA PRINCIPALE ---------------

// Funzione per caricare e processare il file Excel da GitHub
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

// Funzione per elaborare i dati estratti dal file
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
                stipendio: parseValue(row[1]),
                altro: parseValue(row[2]),
                necessita: parseValue(row[5]),
                svago: parseValue(row[6]),
                daRimborsare: parseValue(row[7]),
                giacenza: parseValue(row[10]),      // <-- VALORE DA K4:K15
            disponibilita: parseValue(row[11]) // <-- VALORE DA L4:L15
            };
        }
    }

    statusMessage.style.display = 'none';
    dashboardGrid.classList.remove('hidden');

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

// Funzione per aggiornare i valori nella dashboard
function updateDashboard(mese, datiMensili) {
    const datiDelMese = datiMensili[mese];
    if (!datiDelMese) return;

    document.querySelectorAll('.month-button').forEach(btn => {
        btn.classList.toggle('active', btn.textContent.toLowerCase() === mese.toLowerCase());
    });
    
    const totaleEntrate = datiDelMese.stipendio + datiDelMese.altro;
    const totaleSpese = datiDelMese.necessita + datiDelMese.svago + datiDelMese.daRimborsare;

    totalEntrateEl.textContent = formatCurrency(totaleEntrate);
    stipendioValoreEl.textContent = formatCurrency(datiDelMese.stipendio);
    altroValoreEl.textContent = formatCurrency(datiDelMese.altro);

    totalSpeseEl.textContent = formatCurrency(totaleSpese);
    necessitaValoreEl.textContent = formatCurrency(datiDelMese.necessita);
    svagoValoreEl.textContent = formatCurrency(datiDelMese.svago);
    rimborsareValoreEl.textContent = formatCurrency(datiDelMese.daRimborsare);
    giacenzaValoreEl.textContent = formatCurrency(datiDelMese.giacenza);
    disponibilitaValoreEl.textContent = formatCurrency(datiDelMese.disponibilita);
}

// Avvia il caricamento dei dati quando la pagina è pronta
document.addEventListener('DOMContentLoaded', loadExcelData);


