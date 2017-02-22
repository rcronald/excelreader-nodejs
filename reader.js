var excel = require('exceljs');

var trainDate = new Date("2017-02-01")

var monthTrainPlan = []

var dayOfMonth = 1;

Date.prototype.addDays = function (num) {
    var value = this.valueOf();
    value += 86400000 * num;
    return new Date(value);
}


// read from a file
var workbook = new excel.Workbook();
workbook.xlsx.readFile("Ronal Requena.xlsx")
    .then(function() {

    	var worksheet = workbook.getWorksheet(1);
    	//console.log("worksheet " + worksheet.name)


    	worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
    		if(rowNumber>2 && rowNumber<8){
		    	//console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));

		    	row.eachCell(function(cell, cellNumber){
		    		if(cellNumber>1){
		    			var dayTrainPlan = {}

		    			dayTrainPlan.dayOfMonth = dayOfMonth
		    			dayTrainPlan.date = trainDate
		    			dayTrainPlan.plan = cell.value
		    			//console.log('[' + rowNumber + '-' +  cellNumber + '] = ' + JSON.stringify(cell.value))

		    			trainDate = trainDate.addDays(1)
		    			dayOfMonth++

		    			console.log(dayTrainPlan)
		    			monthTrainPlan.push(dayTrainPlan)
<<<<<<< HEAD
=======
		    			//console.log(monthTrainPlan)
>>>>>>> 80b26f8433f3b92cbb5bd80f70104e9c46a66c54
		    		}
		    	})
    		}
		})

    	//var row = worksheet.getRow(2);
    	//console.log("row2 " + row.getCell(2).value)


    })





