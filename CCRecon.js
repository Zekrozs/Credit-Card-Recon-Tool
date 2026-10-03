"use strict";

function getExcelData(file) {
  (async () => {
    let response;
    try {
      response = await fetch(file);
    } catch (e) {
      throw new Error(`network error: ${e.message}`);
    }

    if (response.status != 200) {
      throw new Error(`something went wrong. status: ${response.status}`);
    }

    let AB;
    try {
      AB = await response.arrayBuffer();
    } catch (e) {
      throw new Error(`Memory error: ${e.message}`);
    }

    const rawFile = await XLSX.read(AB)
    console.log(rawFile)
    return rawFile
    
  })();
}

// async function logWorkbook(file){

// console.log (await getExcelData(file))

// }
const journal = getExcelData("./excel/Journal.xlsx");
const merchant = getExcelData("./excel/Merchant September.xlsx");
