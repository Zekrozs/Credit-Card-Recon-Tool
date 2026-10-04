"use strict";

class workbook {
  #data;
  #sheet
  constructor(data) {
    this.#data = data;
  }

  static async create(file) {
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

    const data = await XLSX.read(AB);

    return new workbook(data);
    
  }

 

    getReconSheet(){
   const workSheet  = this.#data.Sheets[this.#data.SheetNames[0]]
    this.#sheet = XLSX.utils.sheet_to_json(workSheet);
    console.log(this.#sheet)
  }
}

async function call(file) {
  const journal = await workbook.create(file);
  console.log(journal);
  journal.getReconSheet()

}
call("./excel/Journal.xlsx");
call('./excel/Merchant September.xlsx')

