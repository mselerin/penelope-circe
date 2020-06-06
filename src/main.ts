import * as XLSX from 'xlsx';
declare const browser: any;

const SESSION_KEY = 'circe-data';

if (window.location.pathname === '/portail/CE/resEncode.do') {
    initCirceAddon();
} else {
    removeSessionData();
}


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
                    <br />Si le matricule a été correctement traité, il sera surligné en <span style="color: green">vert</span>.
            </strong>
        </p>
        
        <p>
            <strong>Instructions :</strong> 
            <ol> 
                <li>Le fichier Excel doit contenir une feuille nommée "Donn&eacute;es"</li>
                <li>Indiquer le nom des colonnes contenant les matricules et les notes (par défaut : 'Matricule' et 'Note')</li>
                <li>La colonne contenant les notes ne doit pas contenir de formule mais bien une valeur directe</li>
            </ol>
        </p>
        
        <form id="x-circe-form">
            <p id="x-circe-choose-file">
                <label>Fichier Excel :</label>
                <input id="x-circe-excelfile" type="file" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" />
            </p>
                
            <p id="x-circe-session-file" style="display: none;">
                <label>Fichier Excel :</label>
                <span id="x-circe-session-file-name"></span>
                [<a id="x-circe-session-file-clear" href="#">Annuler</a>]
                <span id="x-circe-session-file-warning"></span>
            </p>
            
            <p>
                <label>Nom colonne 'Matricule' :</label>
                <input id="x-circe-column-matricule" type="text" value="Matricule" size="12" />
            </p>
            
            <p>
                <label>Nom colonne 'Note' :</label>
                <input id="x-circe-column-note" type="text" value="Note" size="12" />
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
    document.getElementById('x-circe-excelfile').onchange = readFile;

    document.getElementById('x-circe-column-matricule').onchange = updateSessionData;
    document.getElementById('x-circe-column-note').onchange = updateSessionData;
    document.getElementById('x-circe-round-note').onchange = updateSessionData;

    detectPreviousFile();

    // Ajout d'un listener sur les 2 boutons qui permettent de quitter la page

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
        saveSessionData(sessionData.filename, sessionData.rows);
    }
}

function saveSessionData(filename: string, rows: any[]): void {
    // Sauver dans le localStorage pour le ré-utiliser plus tard
    const colMat = (<HTMLInputElement>document.getElementById('x-circe-column-matricule')).value;
    const colNote = (<HTMLInputElement>document.getElementById('x-circe-column-note')).value;
    const roundNote = (<HTMLInputElement>document.getElementById('x-circe-round-note')).checked;

    const sessionData: SessionData = {
        filename, rows,
        colMat, colNote, roundNote
    };

    sessionStorage.setItem(SESSION_KEY, JSON.stringify(sessionData));
}

function removeSessionData() {
    sessionStorage.removeItem(SESSION_KEY);
}

function detectPreviousFile() {
    const sessionData = getSessionData();
    if (sessionData) {
        (<HTMLElement>document.getElementById('x-circe-choose-file')).style.display = 'none';
        (<HTMLElement>document.getElementById('x-circe-session-file')).style.display = '';
        (<HTMLElement>document.getElementById('x-circe-session-file-name')).innerText = sessionData.filename;

        (<HTMLInputElement>document.getElementById('x-circe-column-matricule')).value = sessionData.colMat;
        (<HTMLInputElement>document.getElementById('x-circe-column-note')).value = sessionData.colNote;
        (<HTMLInputElement>document.getElementById('x-circe-round-note')).checked = sessionData.roundNote;
    }
    else {
        (<HTMLElement>document.getElementById('x-circe-choose-file')).style.display = '';
        (<HTMLElement>document.getElementById('x-circe-session-file')).style.display = 'none';
        (<HTMLElement>document.getElementById('x-circe-session-file-name')).innerText = '';
    }
}

function clearPreviousFile() {
    removeSessionData();
    (<HTMLFormElement>document.getElementById('x-circe-form')).reset();

    detectPreviousFile();
}


function readFile() {
    const warningElem = document.getElementById('x-circe-session-file-warning') as HTMLElement;
    warningElem.style.display = 'none';

    const input = document.getElementById('x-circe-excelfile') as HTMLInputElement;
    if (!input.files) {
        return;
    }

    const file = input.files[0];
    const reader = new FileReader();
    reader.onload = () => {
        const data = reader.result as string;

        try {
            const workbook = XLSX.read(data, {type: 'binary'});
            const sheet = workbook.Sheets["Donn\u00E9es"];
            const rows = XLSX.utils.sheet_to_json(sheet);

            if (!rows || rows.length === 0) {
                throw new Error('invalid_file');
            }

            // Si on a un 'Code Cours' dans la 1e ligne, on tente de valider qu'il s'agit du bon cours
            const codeCours = rows[0]['Code Cours'];
            if (codeCours) {
                const found = Array.from(document.querySelectorAll('td:first-child'))
                    .some(e => e.textContent.trim() === codeCours);

                if (!found) {
                    warningElem.style.display = '';
                    warningElem.innerHTML = `
                        Attention : le code du cours dans le fichier Excel (${codeCours}) ne semble pas correspondre avec celui de la page.
                    `;
                }
            }

            saveSessionData(file.name, rows);
            detectPreviousFile();
        }
        catch (err) {
            alert('Le fichier choisi semble invalide.\nVeuillez vérifier le contenu du fichier.');
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

    const rows = extractMatriculeRows();
    const notes = data.rows;

    rows.forEach(row => {
        const matrElem = document.getElementsByClassName(`matricule_${row.id}`)[0] as HTMLElement;
        const coteElem = document.getElementById(`cote_${row.id}`) as HTMLInputElement;

        const noteRow = notes.find(n => n[data.colMat] === row.matricule);
        if (noteRow) {
            /* pas de coteElem == note déjà encodée par appariteur */

            if (coteElem == null) {
                matrElem.style.backgroundColor = 'green';
            }
            else if (noteRow[data.colNote]) {
                let note = noteRow[data.colNote];

                if (data.roundNote) {
                    if (!isNaN(+note)) {
                        note = '' + (Math.round(+note * 2) / 2);
                    }
                }

                if (coteElem.value === '') {
                    coteElem.value = note;
                    matrElem.style.backgroundColor = 'green';
                }
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

interface SessionData {
    filename: string;
    rows: any[];
    colMat: string;
    colNote: string;
    roundNote: boolean;
}
