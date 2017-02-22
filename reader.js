var excel = require('exceljs');

var trainDate = new Date("2017-02-01")

var monthTrainPlan = []


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
		    			

		    			dayTrainPlan.date = trainDate
		    			dayTrainPlan.plan = cell.value
		    			//console.log('[' + rowNumber + '-' +  cellNumber + '] = ' + JSON.stringify(cell.value))

		    			trainDate = trainDate.addDays(1)

		    			console.log(dayTrainPlan)
		    			monthTrainPlan.push(dayTrainPlan)
		    			//console.log(monthTrainPlan)
		    		}
		    	})
    		}
		})

    	//var row = worksheet.getRow(2);
    	//console.log("row2 " + row.getCell(2).value)


    })





