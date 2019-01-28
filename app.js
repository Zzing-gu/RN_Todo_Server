var express = require('express');
var app = express();

const cors = require('cors');
const bodyParser = require('body-parser');
const morgan = require('morgan');

app.use(morgan('tiny'));
app.use(cors());


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
    // for (const cell of cells) {
    //     console.log(`${cell.row},${cell.col}: ${cell.value}`)
    // }


    var newCells = [];

    

    for (let index = 3; index < cells.length; index++) {
        var tmp = {row: index ,data: '' , task: '', done:false}

        for (const cell of cells) {
            
            if(cell.row == index){
                if(cell.col == 1) {
                    tmp.date = cell.value;
                } else if(cell.col ==2) {
                    tmp.task = cell.value;
                } else if(cell.col ==3) {
                    tmp.done = cell.value;
                }
            }
        }
        
        newCells.push(tmp);
    }

    var jsonCells = {cells : newCells}

    return newCells;
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
    cellDone.value = "";

    await sheet.bulkUpdateCells(cells);

}


app.get('/', function (req, res) {
  res.send('Hello World!');
});

app.get('/all', function (req, res) {

   
    const cells = async () => {
        const cells = await getAll();
        

        

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



const port = process.env.PORT || 1234
app.listen(port, function () {
  console.log(`Example app listening on port ${port}!`);
});


