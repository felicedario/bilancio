// --------------- CONFIGURAZIONE ---------------
// ID del tuo file Excel su Google Drive. Non è più utilizzato per il download diretto, 
// ma lo manteniamo come riferimento se servisse in futuro.
const fileId = "1k7Es36jk9s_mdT4t6RdbuSt460hdmpZA"; 
// Nome del foglio di lavoro da cui leggere i dati
const sheetName = "APP";
// NUOVA CONFIGURAZIONE: Percorso locale del file aggiornato da GitHub Actions
const localExcelPath = "ProvaApp.xlsm"; 

// --------------- FUNZIONI DI UTILITÀ ---------------

// Funzione per formattare i numeri come valuta (es: 1234.56 -> € 1.234,56)
const formatCurrency = (value) => {
    const number = Number(value) || 0;
    return number.toLocaleString('it-IT', { style: 'currency', currency: 'EUR' });
};

// Funzione per "pulire" e convertire i valori letti in numeri
const parseValue = (value) => {
    if (typeof value === 'number') return value;
    if (typeof value !== 'string' || value.trim() === '') return 0;
    
    // Rimuove simbolo euro, punti come separatori migliaia, e sostituisce virgola con punto
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

// Card Saldo/Giacenza (NUOVO)
const saldoMensileEl = document.getElementById('saldo-mensile'); 

// Card Entrate
const totalEntrateEl = document.getElementById('total-entrate');
const stipendioValoreEl = document.getElementById('stipendio-valore');
const altroValoreEl = document.getElementById('altro-valore');

// Card Spese
const totalSpeseEl = document.getElementById('total-spese');
const necessitaValoreEl = document.getElementById('necessita-valore');
const svagoValoreEl = document.getElementById('svago-valore');
const rimborsareValoreEl = document.getElementById('rimborsare-valore');

// Nota: Gli elementi per la card 'Risparmi' non sono dinamici (manca l'ID), 
// ma il resto è pronto.

// --------------- LOGICA PRINCIPALE ---------------

// Funzione per caricare e processare il file Excel locale (ProvaApp.xlsm)
async function loadExcelData() {
    try {
        // Legge il file locale aggiornato da GitHub Actions
        const response = await fetch(localExcelPath);
        
        if (!response.ok) {
            throw new Error(`Errore di rete nel scaricare il file locale (${localExcelPath}, status: ${response.status})`);
        }
        
        const arrayBuffer = await response.arrayBuffer();
        // XLSX è definito in index.html tramite la libreria SheetJS
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
    // Legge il foglio in un array di array, con valori predefiniti per celle vuote
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    // Legge l'anno dalla cella [0][0] (A1)
    const anno = jsonData[0]?.[0] || 'Mensile';
    dashboardTitle.textContent = `Dashboard Finanziaria ${anno}`;

    const mesi = [];
    const datiMensili = {};

    // Cicla sulle righe 3 a 14 per estrarre i dati mensili
    for (let i = 3; i < 15; i++) {
        const row = jsonData[i] || [];
        const mese = row[0]; // Mese (Colonna A, Indice 0)
        
        if (mese && typeof mese === 'string' && mese.trim() !== '') {
            mesi.push(mese);
            datiMensili[mese] = {
                stipendio: parseValue(row[1]), // Colonna B, Indice 1
                altro: parseValue(row[2]),     // Colonna C, Indice 2
                necessita: parseValue(row[5]), // Colonna F, Indice 5
                svago: parseValue(row[6]),     // Colonna G, Indice 6
                daRimborsare: parseValue(row[7]) // Colonna H, Indice 7
            };
        }
    }

    // Nasconde il messaggio di stato e mostra la dashboard
    statusMessage.style.display = 'none';
    dashboardGrid.classList.remove('hidden');

    // Crea i pulsanti dei mesi nel pannello di controllo
    monthButtonsContainer.innerHTML = '';
    mesi.forEach((mese) => {
        const button = document.createElement('button');
        button.className = 'month-button';
        button.textContent = mese.toUpperCase();
        button.onclick = () => updateDashboard(mese, datiMensili);
        monthButtonsContainer.appendChild(button);
    });

    // Seleziona il mese iniziale (il mese corrente o il primo disponibile)
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

    // Aggiorna lo stato "attivo" dei pulsanti
    document.querySelectorAll('.month-button').forEach(btn => {
        btn.classList.toggle('active', btn.textContent.toLowerCase() === mese.toLowerCase());
    });
    
    const totaleEntrate = datiDelMese.stipendio + datiDelMese.altro;
    const totaleSpese = datiDelMese.necessita + datiDelMese.svago + datiDelMese.daRimborsare;
    
    // Calcolo e Aggiornamento Saldo/Giacenza
    const saldoMensile = totaleEntrate - totaleSpese;
    saldoMensileEl.textContent = formatCurrency(saldoMensile);
    
    // Logica per colorare il saldo (verde per positivo, rosso per negativo)
    saldoMensileEl.classList.remove('blue', 'green', 'red'); 
    const saldoClass = saldoMensile >= 0 ? 'green' : 'red';
    saldoMensileEl.classList.add(saldoClass);

    // Aggiornamento Card Entrate
    totalEntrateEl.textContent = formatCurrency(totaleEntrate);
    stipendioValoreEl.textContent = formatCurrency(datiDelMese.stipendio);
    altroValoreEl.textContent = formatCurrency(datiDelMese.altro);

    // Aggiornamento Card Spese
    totalSpeseEl.textContent = formatCurrency(totaleSpese);
    necessitaValoreEl.textContent = formatCurrency(datiDelMese.necessita);
    svagoValoreEl.textContent = formatCurrency(datiDelMese.svago);
    rimborsareValoreEl.textContent = formatCurrency(datiDelMese.daRimborsare);
}

// Avvia il caricamento dei dati quando la pagina è pronta
document.addEventListener('DOMContentLoaded', loadExcelData);
