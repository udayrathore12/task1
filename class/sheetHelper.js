class SheetHelper {

  constructor(fileName) {
        this.fileName = fileName;
        const excel = require('exceljs');
        this.workbook = new excel.Workbook();
  }

  async getSheetName() {
    // get worsheet info
    var sheetName = this.workbook.getWorksheet(1)
    return sheetName['name'];
  }
  
  //Read Data From 
  async readData() {
    // await workbook.xlsx.load(objDescExcel.buffer);
    await this.workbook.xlsx.readFile(this.fileName);
    let jsonData = {}; //output block
    //dataset optimized for better search and match
    const matchArray = ['0.00','0.00%','0.00%c45a10','0.00%ed7d31','0.00%a7d08c','0.00c5e0b3','0.00%538136','0.00bdd7ee','0.00ffe59a','ffe59a','0.007030a0','0.00d8d8d8','0.00bfbfbf','0.00a5a5a5','0.00ffffff','ffffff','0.00%ffffff','true'];
    let styleResult = {};
    this.workbook.worksheets.forEach(function(sheet) {
        //get sheet name
        
        //console.log(sheet.getRow(2)._cells[6]);
        // read first row to get column count
        let firstRow = sheet.getRow(1);
        if (!firstRow.cellCount) return;
        let colCount = firstRow.values;
        //console.log(sheet['_rows'][1]['_cells'][2]['_value']['model']['style']['font']['bold']);
        //console.log(sheet['_rows'][1]['_cells'][6]['style']);
        sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            let rowObj = {};  //store rows (lvl 1)
            let listObj = {}; // store cell (lvl 2)
            let values = row.values; //get row | index starts with 1
            let cellData; //store shared formula cell data


            //start array with 1 because of index of values variable
            for (let i = 1; i < colCount.length; i++) {
              let internalObj = {};
              let styleData = {}; //store style data seprately
              if(typeof values[i] === 'undefined')
              {
                internalObj['text'] = '';
              }
              else
              {
                //check if formula is used
                if(typeof values[i]['formula'] == 'undefined')
                {
                  //check if formula is shared from other cell
                  if(typeof values[i]['sharedFormula'] == 'undefined')
                  {
                    //get value
                    internalObj['text'] = values[i];
                  }
                  else
                  {
                    //get sharedFromula
                    cellData = sheet.getCell(values[i]['sharedFormula']); //get cell
                    internalObj['text'] = "=" + cellData['_value']['model']['formula']; //fetch formula from that cell
                  }
                }
                else
                {
                  //get formula
                  let text = "=" + values[i]['formula'];
                  //console.log(values[i]);
                  //console.log(values[i]['formula']);
                  internalObj['text'] = text;
                }
              }

              if(typeof row['_cells'][i-1] == 'undefined')
              {
              //  internalObj['style'] = 'no match';
              }
              else
              {
                let numFmt = ((typeof row['_cells'][i-1]['style']['numFmt'] == 'undefined') ? '' : row['_cells'][i-1]['style']['numFmt']);
                let myBgColor = ((typeof row['_cells'][i-1]['_value']['model']['style']['fill']['bgColor'] == 'undefined') ? '' :row['_cells'][i-1]['_value']['model']['style']['fill']['bgColor'].argb);
                let bold =  ((typeof row['_cells'][i-1]['_value']['model']['style']['font']['bold'] == 'undefined') ? '' : row['_cells'][i-1]['_value']['model']['style']['font']['bold']);
                let textHash = numFmt + (myBgColor.slice(2)).toLowerCase() + bold;
                let matchResult = matchArray.indexOf(textHash);
                if(matchResult == -1)
                {

                }
                else
                {
                 internalObj['style'] = matchResult;                
                }


              }

              listObj[i-1] = internalObj;

            }
            //jsonData.push(obj);
            rowObj['cells'] = listObj;
            jsonData[rowNumber - 1] = rowObj;

        })
    });
    return jsonData;
  }

  //Read Data From 
  async getColWidth() {
    // await workbook.xlsx.load(objDescExcel.buffer);
    await this.workbook.xlsx.readFile(this.fileName);
    let jsonData = {};
    this.workbook.worksheets.forEach(function(sheet) {
    let defaultColWidth = sheet.getRow(1)['_worksheet']['properties'].defaultColWidth;
    let colData = sheet.getRow(1)['_worksheet']['_columns'];
    let listObj = {}; // store cell (lvl 2)
    let colObj = {}; //store col rows (lvl 1)
    for (let i = 0; i < colData.length; i ++) {
      

      let internalObj = {};
    
      if(typeof colData[i].width === 'undefined')
      {
        internalObj['width'] = defaultColWidth;
      }
      else
      {
        internalObj['width'] = colData[i].width;
      }
      listObj[i] = internalObj;

    }
    colObj = listObj;
    jsonData = colObj; 
    
    })
    return jsonData;
  }
}

module.exports = SheetHelper // Export class