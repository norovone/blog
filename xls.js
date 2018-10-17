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
