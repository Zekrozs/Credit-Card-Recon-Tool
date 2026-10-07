"use strict";

class workbook {
  #data;
  #sheet;
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

  async getReconSheet() {
    const workSheet = this.#data.Sheets[this.#data.SheetNames[0]];
    this.#sheet = XLSX.utils.sheet_to_json(workSheet);
    console.log(this.#sheet);

    // console.log(selectedDay);
  }

  static performRecon(journal, merchant) {
    const paymentDateExcel = "46269"
    const excelPaymentDate = merchant.#sheet.filter(
      (requestedDay) => requestedDay["Posting Date DD:MM:YYYY"] == paymentDateExcel,
    );
    const dayToMs = 24 * 60 * 60 * 1000;
    const parsedPaymentDate = (+paymentDateExcel - 25569) * dayToMs;
    const POSDate = new Date(parsedPaymentDate - 3 * dayToMs);
    const MicrosDate = new Date(parsedPaymentDate - 1 * dayToMs);
    (()=>{
      POSDate.setHours(0,0,0,0)
      MicrosDate.setHours(0,0,0,0)
    })()
    const journalDates = {
      POS:journal.#sheet.filter(
        (postingDate) => new Date(postingDate["BUSINESS_DATE"]).toDateString() == POSDate.toDateString(),
      ),
      Micros: journal.#sheet.filter(
        (postingDate) => new Date(postingDate["BUSINESS_DATE"]).toDateString() == MicrosDate.toDateString(),
      ),
    };

    console.log(new Date(parsedPaymentDate));
    console.log(POSDate);
    console.log(MicrosDate);
    console.log(journalDates);
    console.log(excelPaymentDate)
  }
}

async function call() {
  const Journal = await workbook.create("./excel/Journal.xlsx");
  const Merchant = await workbook.create("./excel/Merchant September.xlsx");
  Journal.getReconSheet()
  Merchant.getReconSheet()

workbook.performRecon(Journal,Merchant)
}
call();

