<<<<<<< HEAD
const XlsxPopulate = require('xlsx-populate');
 
    XlsxPopulate.fromFileAsync("./2.xlsx")
    .then(workbook => {
        
        var i = 2;
        var j = 10;
        var q = 0;
        var cell = workbook.sheet(0).row(2).cell(8).value();
        while ( workbook.sheet(0).row(i).cell(1).value() !== undefined ) {

            if (cell > workbook.sheet(0).row(i).cell(8).value()) {
                cell = workbook.sheet(0).row(i).cell(8).value();
            }

            if ( cell === workbook.sheet(0).row(i).cell(8).value() &&  q === 0 ) {
                workbook.sheet(0).row(j).cell(15).value("B");
                workbook.sheet(0).row(j).cell(16).value(workbook.sheet(0).row(i).cell(8).value());
                workbook.sheet(0).row(j).cell(17).value(workbook.sheet(0).row(i).cell(4).value());
                q = 1;
                j++;
            }

            if ( cell !== workbook.sheet(0).row(i).cell(8).value() &&  q === 0 ) {
                workbook.sheet(0).row(j).cell(15).value("S");
                workbook.sheet(0).row(j).cell(16).value(workbook.sheet(0).row(i).cell(8).value());
                workbook.sheet(0).row(j).cell(17).value(workbook.sheet(0).row(i).cell(4).value());                
                q = -1;
                j++;
            }

            if ( cell === workbook.sheet(0).row(i).cell(8).value() &&  q === -1 ) {
                workbook.sheet(0).row(j).cell(15).value("B");
                workbook.sheet(0).row(j).cell(16).value(workbook.sheet(0).row(i).cell(8).value());
                workbook.sheet(0).row(j).cell(17).value(workbook.sheet(0).row(i).cell(4).value());                
                q = 1;
                j++;
            }

            if ( cell !== workbook.sheet(0).row(i).cell(8).value() &&  q === 1 ) {
                workbook.sheet(0).row(j).cell(15).value("S");
                workbook.sheet(0).row(j).cell(16).value(workbook.sheet(0).row(i).cell(8).value());
                workbook.sheet(0).row(j).cell(17).value(workbook.sheet(0).row(i).cell(4).value());                
                q = -1;
                j++;
            }

            i++
        }
        workbook.sheet(0).row(j).cell(16).value(workbook.sheet(0).row(i-1).cell(8).value());
        
        return workbook.toFileAsync("./out.xlsx");
    });
=======
MyXlsx = require('xlsx-populate');

MyXlsx.fromBlankAsync()
    .then(workbook => {
var i = 1;
        
    while (i < 20)  {
        
        workbook.sheet("Sheet1").row(i).cell(1).value(i); 

        console.log(i);
        i++
    }
   
    return workbook.toFileAsync("new.xlsx");
})


const XlsxPopulate = require('xlsx-populate');
 
// Load an existing workbook
XlsxPopulate.fromFileAsync("./old.xlsx")
    .then(workbook => {
        // Modify the workbook.
        const value = workbook.sheet("Sheet1").cell("A1").value();
        
        var i = 1;
        
        while (value != undefined)  {
        
            const value = workbook.sheet("Sheet1").row(i).cell(1).value(); 
            
            if (value == undefined) {
                console.log("exit");
                break;
            }

            console.log(value);
        i++
    }
});
>>>>>>> 46984f92268b1ed906095e8c57f2a388cba3fb8c
