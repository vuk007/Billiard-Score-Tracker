const Igra = {
  _8ball: "8ball",
  _9ball: "9ball",
  _10ball: "10ball"
}; 

const Kolone = {
  "8ball": { VUK: 7, NAUM: 8 },
  "9ball": { VUK: 9, NAUM: 10 },
  "10ball": { VUK: 11, NAUM: 12 }
};


function findDateColumn (tabela) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timezone = ss.getSpreadsheetTimeZone();
    const lastRow = Math.max(tabela.getLastRow(), 1);

    const datumi = tabela.getRange(1, 6, lastRow, 1).getDisplayValues();

    for (let i = 0; i < datumi.length; i++) {
      if (datumi[i][0].trim().toUpperCase() === "DATUM:") {
        headerRow = i + 1;
        break;
      }
    }
    const startRow = headerRow + 1;
    const today = Utilities.formatDate(new Date(),timezone,"dd/MM/yyyy");

    let lastDateRow = headerRow;

    for (let row = startRow; row <= lastRow; row++) {
      const cell = tabela.getRange(row, 6); // 6 je F kolona u kojoj je i datum smesten 
      const value = cell.getValue();
      const display = cell.getDisplayValue().trim();

      if (value instanceof Date) {
        const cellDate = Utilities.formatDate(value,timezone,"dd/MM/yyyy"); //trebao bi ovako da je upisan 
        if (cellDate === today) {
          return row;
        } 
      }
      else if (display === today) { // ako je upisan u drugom formatu datum 
        return row;
      }
      if (display !== "") { // prosli smo sve datume i nema danasnjeg novi treba da se napravi
        lastDateRow = row;
      }
    }
    tabela.getRange(lastDateRow + 1,6).setValue(new Date()).setNumberFormat("dd/MM/yyyy");
    tabela.getRange(lastDateRow + 1, 7, 1, 6).setValues([[0, 0, 0, 0, 0, 0]]);
    return lastDateRow+1;
}

function promeniRezultat(igrac , igra , inkrement){
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000); //10 sek pokusava da preuzme bravu
  try{
    const tabela = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const row = findDateColumn(tabela);
    const kolona = Kolone[igra]?.[igrac];
    if (!kolona) {
      throw new Error("Nepoznata igra ili igrač.");
    }
    var cell = tabela.getRange(row,kolona);
    cell.setValue(cell.getValue() + inkrement);
  }
  finally {
    lock.releaseLock();
  }
}
// VUK
function vuk8Plus()  { promeniRezultat("VUK", Igra._8ball, 1); }
function vuk8Minus() { promeniRezultat("VUK", Igra._8ball, -1); }

function vuk9Plus()  { promeniRezultat("VUK", Igra._9ball, 1); }
function vuk9Minus() { promeniRezultat("VUK", Igra._9ball, -1); }

function vuk10Plus()  { promeniRezultat("VUK", Igra._10ball, 1); }
function vuk10Minus() { promeniRezultat("VUK", Igra._10ball, -1); }

// NAUM
function naum8Plus()  { promeniRezultat("NAUM", Igra._8ball, 1); }
function naum8Minus() { promeniRezultat("NAUM", Igra._8ball, -1); }

function naum9Plus()  { promeniRezultat("NAUM", Igra._9ball, 1); }
function naum9Minus() { promeniRezultat("NAUM", Igra._9ball, -1); }

function naum10Plus()  { promeniRezultat("NAUM", Igra._10ball, 1); }
function naum10Minus() { promeniRezultat("NAUM", Igra._10ball, -1); }



function submitIzbor() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const igraci = sheet.getRange("B2:C2").getValues()[0];
  const operacije = sheet.getRange("D2:E2").getValues()[0];
  const igre = sheet.getRange("B4:D4").getValues()[0];

  function brojIzabranih(opcije) {
    return opcije.filter(x => x === true).length;
  }

  if (
    brojIzabranih(igraci) !== 1 ||
    brojIzabranih(operacije) !== 1 ||
    brojIzabranih(igre) !== 1
  ) {
    SpreadsheetApp.getUi().alert(
      "Izaberi jednog igrača, jednu operaciju i jednu igru!"
    );
    return;
  }

  const igrac = igraci[0] ? "VUK" : "NAUM";
  const inkrement = operacije[0] ? 1 : -1;

  const dostupneIgre = [
    Igra._8ball,
    Igra._9ball,
    Igra._10ball
  ];

  const indeksIgre = igre.findIndex(x => x === true);
  const igra = dostupneIgre[indeksIgre];

  // Mapiranje kombinacija na postojeće funkcije
  const akcije = {
    "VUK|8ball|1": vuk8Plus,
    "VUK|8ball|-1": vuk8Minus,
    "VUK|9ball|1": vuk9Plus,
    "VUK|9ball|-1": vuk9Minus,
    "VUK|10ball|1": vuk10Plus,
    "VUK|10ball|-1": vuk10Minus,

    "NAUM|8ball|1": naum8Plus,
    "NAUM|8ball|-1": naum8Minus,
    "NAUM|9ball|1": naum9Plus,
    "NAUM|9ball|-1": naum9Minus,
    "NAUM|10ball|1": naum10Plus,
    "NAUM|10ball|-1": naum10Minus
  };

  const kljuc = `${igrac}|${igra}|${inkrement}`;

  // Pozivanje odgovarajuće funkcije
  akcije[kljuc]();

  // Resetovanje checkbox-ova
  sheet.getRange("B2:E2").setValue(false);
  sheet.getRange("B4:D4").setValue(false);
}

// provera da li su istovremeno vise checkboxova iste kategorije upaljeni 
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const cell = e.range.getA1Notation();
  
  if (cell == "E4"){ //submit za telefon checkbox 
    sheet.getRange(cell).setValue(false);
    submitIzbor();
  }
  else{
    const grupe = [
      ["B2", "C2"],             // Vuk / Naum
      ["D2", "E2"],             // +1 / -1
      ["B4", "C4", "D4"]  // 8ball / 9ball / 10ball
    ];

    for (const grupa of grupe) {
      if (grupa.includes(cell) && e.value === "TRUE") {
        for (const adresa of grupa) {
          if (adresa !== cell) {
            sheet.getRange(adresa).setValue(false);
          }
        }
        break;
      }
    }
  }
}

