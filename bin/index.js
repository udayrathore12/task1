#!/usr/bin/env node

const sheetHelper = require('../class/sheetHelper.js');
//array to print
const styleData = [
    {
      "format": "numberNoDecimal"
    },
    {
      "format": "percentNoDecimal"
    },
    {
      "format": "percentNoDecimal",
      "bgcolor": "#c45a10"
    },
    {
      "format": "percentNoDecimal",
      "bgcolor": "#ed7d31"
    },
    {
      "format": "percentNoDecimal",
      "bgcolor": "#a7d08c"
    },
    {
      "format": "numberNoDecimal",
      "bgcolor": "#c5e0b3"
    },
    {
      "format": "percentNoDecimal",
      "bgcolor": "#538136"
    },
    {
      "format": "numberNoDecimal",
      "bgcolor": "#bdd7ee"
    },
    {
      "format": "numberNoDecimal",
      "bgcolor": "#ffe59a"
    },
    {
      "bgcolor": "#ffe59a"
    },
    {
      "format": "numberNoDecimal",
      "bgcolor": "#7030a0"
    },
    {
      "format": "numberNoDecimal",
      "bgcolor": "#d8d8d8"
    },
    {
      "format": "numberNoDecimal",
      "bgcolor": "#bfbfbf"
    },
    {
      "format": "numberNoDecimal",
      "bgcolor": "#a5a5a5"
    },
    {
      "format": "numberNoDecimal",
      "bgcolor": "#ffffff"
    },
    {
      "bgcolor": "#ffffff"
    },
    {
      "format": "percentNoDecimal",
      "bgcolor": "#ffffff"
    },
    {
      "font": {
        "bold": true
      }
    }
  ];

let finalOutput = {};

let SH = new sheetHelper(process.argv[2]);
async function printData() {
    const rowData = await SH.readData();
    const colData = await SH.getColWidth();
    finalOutput['name'] = await SH.getSheetName();
    finalOutput['freeze'] = "A1";
    finalOutput['merges'] = [];
    finalOutput['style'] = styleData;
    finalOutput['rows'] = rowData;
    finalOutput['cols'] = colData;
    finalOutput['validation'] = [];
    
    console.log(JSON.stringify(finalOutput));
}

printData();

