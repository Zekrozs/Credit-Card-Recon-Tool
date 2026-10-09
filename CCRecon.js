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
    const paymentDateExcel = "46269";
    const excelPaymentDate = merchant.#sheet.filter(
      (requestedDay) =>
        requestedDay["Posting Date DD:MM:YYYY"] == paymentDateExcel,
    );

    // need to create an object to oranize the data when refractoring
    const manualPayments = excelPaymentDate.filter(
      (transaction) => transaction["Channel"] !== "ECOM",
    ).filter(card => card['Card Type']?.toUpperCase().includes('VISA'));
    const ECOMPayments = excelPaymentDate.filter(
      (transaction) => transaction["Channel"]?.toUpperCase() === "ECOM",
    ).filter(card => card['Card Type']?.toUpperCase().includes('VISA'));

    const dayToMs = 24 * 60 * 60 * 1000;
    const parsedPaymentDate = (+paymentDateExcel - 25569) * dayToMs;
    const manualDate = new Date(parsedPaymentDate - 3 * dayToMs);
    const MicrosDate = new Date(parsedPaymentDate - 1 * dayToMs);
    (() => {
      manualDate.setHours(0, 0, 0, 0);
      MicrosDate.setHours(0, 0, 0, 0);
    })();
    const journalDates = {
      Manual: journal.#sheet.filter(
        (manualPosting) =>
          new Date(manualPosting["BUSINESS_DATE"]).toDateString() ==
          manualDate.toDateString(),
      ),
      Micros: journal.#sheet.filter(
        (POSPosting) =>
          new Date(POSPosting["BUSINESS_DATE"]).toDateString() ==
          MicrosDate.toDateString(),
      ),
    };

    console.log(new Date(parsedPaymentDate));
    console.log(manualDate);
    console.log(MicrosDate);
    console.log(journalDates);
    console.log(manualPayments);
    console.log(ECOMPayments);
  }
}

async function call() {
  const Journal = await workbook.create("./excel/Journal.xlsx");
  const Merchant = await workbook.create("./excel/Merchant September.xlsx");
  await Journal.getReconSheet();
  await Merchant.getReconSheet();

  workbook.performRecon(Journal, Merchant);
}
call();

// the files
// the payment date
// payment method
