const express = require('express');
const bodyparser = require('body-parser');
const app = express();
var path = require('path');
var XLSX = require('xlsx');
var fs = require('fs');
URL = require('url');
var FileSave = require('file-saver');

var server;
app.get('/xlsx', (req, res) => {
    var data = [
        {
            "id": 1,
            "name": "Leanne Graham",
            "username": "Bret",
            "email": "Sincere@april.biz",
            "address": {
                "street": "Kulas Light",
                "suite": "Apt. 556",
                "city": "Gwenborough",
                "zipcode": "92998-3874",
                "geo": {
                    "lat": "-37.3159",
                    "lng": "81.1496"
                }
            },
            "phone": "1-770-736-8031 x56442",
            "website": "hildegard.org",
            "company": {
                "name": "Romaguera-Crona",
                "catchPhrase": "Multi-layered client-server neural-net",
                "bs": "harness real-time e-markets"
            }
        },
        {
            "id": 2,
            "name": "Ervin Howell",
            "username": "Antonette",
            "email": "Shanna@melissa.tv",
            "address": {
                "street": "Victor Plains",
                "suite": "Suite 879",
                "city": "Wisokyburgh",
                "zipcode": "90566-7771",
                "geo": {
                    "lat": "-43.9509",
                    "lng": "-34.4618"
                }
            },
            "phone": "010-692-6593 x09125",
            "website": "anastasia.net",
            "company": {
                "name": "Deckow-Crist",
                "catchPhrase": "Proactive didactic contingency",
                "bs": "synergize scalable supply-chains"
            }
        },
        {
            "id": 3,
            "name": "Clementine Bauch",
            "username": "Samantha",
            "email": "Nathan@yesenia.net",
            "address": {
                "street": "Douglas Extension",
                "suite": "Suite 847",
                "city": "McKenziehaven",
                "zipcode": "59590-4157",
                "geo": {
                    "lat": "-68.6102",
                    "lng": "-47.0653"
                }
            },
            "phone": "1-463-123-4447",
            "website": "ramiro.info",
            "company": {
                "name": "Romaguera-Jacobson",
                "catchPhrase": "Face to face bifurcated interface",
                "bs": "e-enable strategic applications"
            }
        }
    ];
    // var wb = xlsx.utils.book_new();
    // var excelconvert = xlsx.utils.json_to_sheet(data);

    var ws = XLSX.utils.json_to_sheet(data);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
    // var workbook = xlsx.utils.book_append_sheet(wb, excelconvert, 'Hello')
    // var dataExcel = xlsx.writeFile(workbook, 'test.xlsx', {
    //     type: "file"
    // })
    // res.send(data)
    var excel = XLSX.writeFile(wb, 'TESTERs.xlsx')
    res.download('TEST.xlsx')
    //  saveAs(new Blob([excel],{type:"application/octet-stream"}), "test.xlsx");
    // res.status(200).download(excel)


    // var stream = XLSX.stream.to_csv(excelconvert);
    // stream.pipe(fs.createWriteStream('output_file_name'));
})
app.get('/', (req, res) => {
    res.send('OK  test')
})

function startService() {
    server = app.listen(3000, 'localhost', () => {
        const port = server.address().port;
        const hostname = server.address().address;
        console.log(`Server running at ${hostname}:${port}`);
    })
}
startService();

function stopSerives() {
    console.log('Close Services');

    server.close();
}

process.on("SIGTERM", () => {
    stopSerives();
});

process.on("SIGINT", () => {
    stopSerives();
});