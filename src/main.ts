import * as XLSX from 'xlsx';
declare const browser: any;

const SESSION_KEY = 'circe-data';

initCirceAddon();
detectPreviousFile();


function initCirceAddon() {
    const circeAddon = document.getElementById('x-circe');
    if (circeAddon) {
        circeAddon.parentElement.removeChild(circeAddon);
    }

    const circeLogo = browser.runtime.getURL('images/circe-96.jpg');

    const template = `
<div id="x-circe">
    <div>
        <p><img src="${circeLogo}" alt="Circ&eacute;" /></p>    
    </div>
    
    <div>
        <p>
            <strong>Circé recherche les matricules du formulaire d'encodage ci-dessous et 
                    injecte les notes en correspondance avec les matricules trouvés dans le fichier Excel donné.
                    
                    <br />Si le matricule n'est pas présent dans le fichier Excel, il sera surligné en <span style="color: orange">orange</span>.
                    <br />Si le matricule a été traité, il sera surligné en <span style="color: green">vert</span>.
            </strong>
        </p>
        
        <p>
            <strong>Instructions :</strong> 
            <ul> 
                <li>Le fichier Excel doit contenir une feuille nommée "Donn&eacute;es"</li>
                <li>Indiquer les colonnes contenant les matricules et les notes (A = 0 ; B = 1 ; ...)</li>
                <li>La colonne contenant les notes ne doit pas contenir de formule mais bien une valeur directe</li>
            </ul>
        </p>
        
        <form>
            <div id="x-circe-choose-file">
                <p>
                    <label>Fichier Excel :</label>
                    <input id="x-circe-excelfile" type="file" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" />
                </p>
                
                <p>
                    <label>N&deg; colonne 'Matricule' :</label>
                    <input id="x-circe-column-matricule" type="text" value="0" size="2" />
                    <small>(colonne A = 0, colonne B = 1, ...)</small>
                </p>
                
                <p>
                    <label>N&deg; colonne 'Note' :</label>
                    <input id="x-circe-column-note" type="text" value="3" size="2" />
                    <small>(colonne A = 0, colonne B = 1, ...)</small>
                </p>
            </div>
                
            <p id="x-circe-session-file" style="display: none;">
                <label>Fichier Excel :</label>
                <span id="x-circe-session-file-name"></span>
                [<a id="x-circe-session-file-clear" href="#">Annuler</a>]
            </p>
            
            <p>
                <label>Arrondir les notes :</label>
                <input id="x-circe-round-note" type="checkbox" />
            </p>
            
            <p>
                <input id="x-circe-btn-fill-notes" type="button" value="Préremplir Notes" />
                <input id="x-circe-btn-fill-e" type="button" value="Préremplir E" />
                <input id="x-circe-btn-clear-notes" type="button" value="Vider les notes" />
            </p>
        </form>
    </div>
</div>
`;


    // Insertion du bloc dans la page
    const div = document.createElement('div');
    div.innerHTML = template;

    const content = document.getElementById('fw16Content');
    content.insertBefore(div.firstElementChild, content.children[0]);

    // Gestion des events
    document.getElementById('x-circe-btn-fill-notes').onclick = fillNotes;
    document.getElementById('x-circe-btn-fill-e').onclick = fillE;
    document.getElementById('x-circe-btn-clear-notes').onclick = clearNotes;
    document.getElementById('x-circe-session-file-clear').onclick = clearPreviousFile;
    document.getElementById('x-circe-round-note').onchange = updateSessionData;

    const inputFile = document.getElementById('x-circe-excelfile') as HTMLInputElement;
    inputFile.onchange = readFile;
}



function getSessionData(): SessionData {
    const sessionData = sessionStorage.getItem(SESSION_KEY);
    if (sessionData) {
        return JSON.parse(sessionData) as SessionData;
    }

    return null;
}

function updateSessionData(): void {
    const sessionData = getSessionData();
    if (sessionData) {
        saveSessionData(sessionData.filename, sessionData.notes);
    }
}

function saveSessionData(filename: string, notes: NoteRow[]): void {
    // Sauver dans le localStorage pour le ré-utiliser plus tard
    const colMat = +(<HTMLInputElement>document.getElementById('x-circe-column-matricule')).value;
    const colNote = +(<HTMLInputElement>document.getElementById('x-circe-column-note')).value;
    const roundNote = (<HTMLInputElement>document.getElementById('x-circe-round-note')).checked;

    const sessionData: SessionData = {
        filename, notes,
        colMat, colNote, roundNote
    };

    sessionStorage.setItem(SESSION_KEY, JSON.stringify(sessionData));
}

function detectPreviousFile() {
    const sessionData = getSessionData();
    if (sessionData) {
        (<HTMLElement>document.getElementById('x-circe-choose-file')).style.display = 'none';
        (<HTMLElement>document.getElementById('x-circe-session-file')).style.display = '';
        (<HTMLElement>document.getElementById('x-circe-session-file-name')).innerText = sessionData.filename;

        (<HTMLInputElement>document.getElementById('x-circe-column-matricule')).value = '' + sessionData.colMat;
        (<HTMLInputElement>document.getElementById('x-circe-column-note')).value = '' + sessionData.colNote;
        (<HTMLInputElement>document.getElementById('x-circe-round-note')).checked = sessionData.roundNote;
    }
    else {
        (<HTMLElement>document.getElementById('x-circe-choose-file')).style.display = '';
        (<HTMLElement>document.getElementById('x-circe-session-file')).style.display = 'none';
        (<HTMLElement>document.getElementById('x-circe-session-file-name')).innerText = '';
    }
}

function clearPreviousFile() {
    sessionStorage.removeItem(SESSION_KEY);

    const inputFile = document.getElementById('x-circe-excelfile') as HTMLInputElement;
    inputFile.value = null;

    detectPreviousFile();
}


function readFile() {
    const input = document.getElementById('x-circe-excelfile') as HTMLInputElement;
    if (!input.files) {
        return;
    }

    const file = input.files[0];
    const reader = new FileReader();
    reader.onload = () => {
        const data = reader.result as string;

        try {
            const rows: NoteRow[] = [];
            const colMat = +(<HTMLInputElement>document.getElementById('x-circe-column-matricule')).value;
            const colNote = +(<HTMLInputElement>document.getElementById('x-circe-column-note')).value;

            const workbook = XLSX.read(data, {type: 'binary'});
            const sheet = workbook.Sheets["Donn\u00E9es"];
            const range = XLSX.utils.decode_range(sheet["!ref"]);

            for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
                const firstCell = sheet[XLSX.utils.encode_cell({r: rowNum, c: colMat})];
                const secondCell = sheet[XLSX.utils.encode_cell({r: rowNum, c: colNote})];

                if (secondCell) {
                    rows.push({
                        matricule: firstCell.v,
                        note: secondCell.v
                    });
                }
            }

            if (!rows || rows.length === 0) {
                throw new Error('invalid_file');
            }

            saveSessionData(file.name, rows);
            detectPreviousFile();
        }
        catch (err) {
            alert('Le fichier choisi semble invalide.\nVeuillez vérifier le contenu du fichier et les numéros de colonnes.');
            console.log(err);
            clearPreviousFile();
        }
    };

    reader.readAsBinaryString(file);
}



function fillNotes() {
    // Reprendre les données dans le sessionStorage
    const data = getSessionData();
    if (!data) {
        return;
    }

    const roundNote = (<HTMLInputElement>document.getElementById('x-circe-round-note')).checked;

    const rows = extractMatriculeRows();
    rows.forEach(row => {
        const matrElem = document.getElementsByClassName(`matricule_${row.id}`)[0] as HTMLElement;
        const coteElem = document.getElementById(`cote_${row.id}`) as HTMLInputElement;

        const noteRow = data.notes.find(n => n.matricule === row.matricule);
        if (noteRow) {
            /* pas de coteElem == note déjà encodée par appariteur */

            if (coteElem == null) {
                matrElem.style.backgroundColor = 'green';
            }
            else if (noteRow.note) {
                let note = noteRow.note;

                if (roundNote) {
                    if (!isNaN(+note)) {
                        note = '' + (Math.round(+note * 2) / 2);
                    }
                }

                coteElem.value = note;
                matrElem.style.backgroundColor = 'green';
            }
        }
        else {
            matrElem.style.backgroundColor = 'orange';
        }
    });
}


function fillE() {
    const rows = extractMatriculeRows();
    rows.forEach(row => {
        const matrElem = document.getElementsByClassName(`matricule_${row.id}`)[0] as HTMLElement;
        const coteElem = document.getElementById(`cote_${row.id}`) as HTMLInputElement;

        /* pas de coteElem == note déjà encodée par appariteur */

        if (coteElem == null) {
            matrElem.style.backgroundColor = 'green';
        }
        else if (coteElem.value === '') {
            coteElem.value = 'E';
            matrElem.style.backgroundColor = 'green';
        }
    });
}


function clearNotes() {
    const rows = extractMatriculeRows();
    rows.forEach(row => {
        const matrElem = document.getElementsByClassName(`matricule_${row.id}`)[0] as HTMLElement;
        const coteElem = document.getElementById(`cote_${row.id}`) as HTMLInputElement;

        matrElem.style.backgroundColor = '';

        if (coteElem) {
            coteElem.value = '';
        }
    });
}


function extractMatriculeRows(): MatriculeRow[] {
    const rows = [];
    const items = document.querySelectorAll("td[class^='matricule_']");

    for (let i = 0; i < items.length; i++) {
        const id = items[i].className.substring('matricule_'.length);
        const matricule = items[i].textContent;

        rows.push({id, matricule});
    }

    return rows;
}


interface MatriculeRow {
    matricule: string;
    id: string;
}

interface NoteRow {
    matricule: string;
    note: string;
}

interface SessionData {
    filename: string;
    notes: NoteRow[];
    colMat: number;
    colNote: number;
    roundNote: boolean;
}
