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