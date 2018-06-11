const express = require('express');
const bodyparser = require('body-parser');
const app = express();
var path = require('path');
var XLSX = require('xlsx');
var fs = require('fs');

var server;
app.get('/xlsx', (req, res) => {
    var data = [{
            "ticketID": "TID1528277110235",
            "createdAt": "06 06 2018 16:25:10",
            "modifiedAt": "",
            "status": 0,
            "type": "Yoga (All Level) by K.Ning",
            "DateType": "2018-04-02",
            "owner": {
                "employeeID": "8003986",
                "firstName": "Kullapa",
                "lastName": "Makarabhirom"
            }
        },
        {
            "ticketID": "TID1528277222666",
            "createdAt": "06 06 2018 16:27:02",
            "modifiedAt": "",
            "status": 0,
            "type": "bodyBalance (All Level) by K.Ning",
            "DateType": "2018-04-02",
            "owner": {
                "employeeID": "8003986",
                "firstName": "Kullapa",
                "lastName": "Makarabhirom"
            }
        },
        {
            "ticketID": "TID1528277238702",
            "createdAt": "06 06 2018 16:27:18",
            "modifiedAt": "",
            "status": 0,
            "type": "bodyBalance (All Level) by K.Ning",
            "DateType": "2018-04-22",
            "owner": {
                "employeeID": "8003986",
                "firstName": "Kullapa",
                "lastName": "Makarabhirom"
            }
        },
        {
            "ticketID": "TID1528277259875",
            "createdAt": "06 06 2018 16:27:39",
            "modifiedAt": "",
            "status": 0,
            "type": "ZumbaDance (All Level) by K.Ning",
            "DateType": "2018-04-02",
            "owner": {
                "employeeID": "8003986",
                "firstName": "Kullapa",
                "lastName": "Makarabhirom"
            }
        },
        {
            "ticketID": "TID1528341563854",
            "createdAt": "07 06 2018 10:19:23",
            "modifiedAt": "",
            "status": 0,
            "type": "Yoga (All Level) by K.Ning",
            "DateType": "2018-04-22",
            "owner": {
                "employeeID": "6068132",
                "firstName": "Pitsanu",
                "lastName": "Limpanachaiphonkul"
            }
        }
    ]

    var mongoXlsx = require('mongo-xlsx');

    // var data = [{
    //         name: "Peter",
    //         lastName: "Parker",
    //         isSpider: true
    //     },
    //     {
    //         name: "Remy",
    //         lastName: "LeBeau",
    //         powers: ["kinetic cards"]
    //     }
    // ];

    /* Generate automatic model for processing (A static model should be used) */
    models = [
        // {
        //     "displayName": "User Identifier",
        //     "access": "_id",
        //     "type": "string"
        // },
        {
            "displayName": "TicketID",
            "access": "ticketID",
            "type": "string"
        },
        {
            "displayName": "Created",
            "access": "createdAt",
            "type": "string"
        },
        {
            "displayName": "Approved DATE",
            "access": "modifiedAt",
            "type": "string"
        },
        {
            "displayName": "Status",
            "access": "status",
            "type": "number"
        },
        {
            "displayName": "Employee ID",
            "access": "owner[employeeID]",
            "type": "string"
        },
        {
            "displayName": "First Name",
            "access": "owner[firstName]",
            "type": "string"
        },
        {
            "displayName": "Last Name",
            "access": "owner[lastName]",
            "type": "string"
        },
    ]

    var model = mongoXlsx.buildDynamicModel(data);

    /* Generate Excel */
    mongoXlsx.mongoData2Xlsx(data, models, function (err, data) {
        console.log('File saved at:', data.fullPath);
    });
    /* Read Excel */
    mongoXlsx.xlsx2MongoData("./file.xlsx", models, function (err, mongoData) {
        console.log('Mongo data:', mongoData);
    });
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