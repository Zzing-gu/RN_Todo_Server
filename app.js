var express = require('express');
var app = express();

const GoogleSpreadsheet = require('google-spreadsheet')
const { promisify } = require('util')

const credentials = require(`./service-account.json`)

const SPREADSHEET_ID = `1C1U5twYkqSaRw705i4tAHFFSIT0cD4HZyrwFWeEqPas`

let mysheet = null;

async function accessSpreadsheet() {
    const doc = new GoogleSpreadsheet(SPREADSHEET_ID)
    await promisify(doc.useServiceAccountAuth)(credentials)
    const info = await promisify(doc.getInfo)()
    console.log(`Loaded doc: ` + info.title + ` by ` + info.author.email)
    const sheet = info.worksheets[0]
    mysheet = info.worksheets[0]
    console.log(
      `sheet 1: ` + sheet.title + ` ` + sheet.rowCount + `x` + sheet.colCount
    )
}


async function getAll() {
    const doc = new GoogleSpreadsheet(SPREADSHEET_ID)
    await promisify(doc.useServiceAccountAuth)(credentials)
    const info = await promisify(doc.getInfo)()
    console.log(`Loaded doc: ` + info.title + ` by ` + info.author.email)
    const sheet = info.worksheets[0]

    console.log(
      `sheet 1: ` + sheet.title + ` ` + sheet.rowCount + `x` + sheet.colCount
    )


    const cells = await promisify(sheet.getCells)({
        'min-row': 1,
        'max-row': 5,
        'min-col': 1,
        'max-col': 3,
        'return-empty': true,
    })
    for (const cell of cells) {
        console.log(`${cell.row},${cell.col}: ${cell.value}`)
    }

    return cells;
}

async function insert(row ,date, task, done) {
    const doc = new GoogleSpreadsheet(SPREADSHEET_ID)
    await promisify(doc.useServiceAccountAuth)(credentials)
    const info = await promisify(doc.getInfo)()
    console.log(`Loaded doc: ` + info.title + ` by ` + info.author.email)
    const sheet = info.worksheets[0]

    const cells = await promisify(sheet.getCells)({
        'min-row': 1,
        'max-row': 15,
        'min-col': 1,
        'max-col': 3,
        'return-empty': true,
    })
    
    var cellDate = cells[3*row-3]
    var cellTask = cells[3*row-2]
    var cellDone = cells[3*row-1]

    cellDate.value = date;
    cellTask.value = task;
    cellDone.value = done;

    await sheet.bulkUpdateCells(cells);

}


async function delrow(row) {
    const doc = new GoogleSpreadsheet(SPREADSHEET_ID)
    await promisify(doc.useServiceAccountAuth)(credentials)
    const info = await promisify(doc.getInfo)()
    console.log(`Loaded doc: ` + info.title + ` by ` + info.author.email)
    const sheet = info.worksheets[0]

    const cells = await promisify(sheet.getCells)({
        'min-row': 1,
        'max-row': 15,
        'min-col': 1,
        'max-col': 3,
        'return-empty': true,
    })
    
    var cellDate = cells[3*row-3]
    var cellTask = cells[3*row-2]
    var cellDone = cells[3*row-1]

    cellDate.value = "";
    cellTask.value = "";
    cellDone.value = false;

    await sheet.bulkUpdateCells(cells);

}


app.get('/', function (req, res) {
  res.send('Hello World!');
});

app.get('/all', function (req, res) {

   
    const cells = async () => {
        const cells = await getAll();
        console.log(cells[0].value);

        

        res.json({ cells : cells })
        
    } 

    cells();
    
});


app.get('/insert', function (req, res) {
    
    console.log(req.query.row);
    insert(req.query.row, req.query.date, req.query.task, req.query.done)
    res.send('insert');
});


app.get('/delete', function (req, res) {
    
    console.log(req.query.row);
    delrow(req.query.row)
    res.send('delete');
});




app.listen(3000, function () {
  console.log('Example app listening on port 3000!');
});


