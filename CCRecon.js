"use strict";

class workbook {
  #data;
  #sheet;
  constructor(data) {
    this.#data = data;
  }
  // file is added by user
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
    if (!data) {
      throw new Error("something went wrong, could not read workbook data");
    }

    return new workbook(data);
  }
  async getReconSheet() {
    if (!this.#data) {
      throw new Error(`something went wrong`);
    }
    const workSheet = this.#data?.Sheets[this.#data.SheetNames[0]];
    this.#sheet = await XLSX.utils.sheet_to_json(workSheet);
  }

  static getTransactions(journal, merchant, excelPaymentDate) {
    if (!journal.#sheet) {
      throw new Error(`recon sheet is unavaliable`);
    }
    const selectedPaymentDate = merchant.#sheet.filter(
      (requestedDay) =>
        requestedDay["Posting Date DD:MM:YYYY"] == excelPaymentDate,
    );
    if (!selectedPaymentDate) {
      throw new Error(`could not get date: ${selectedPaymentDate}`);
    }
    // need to create an object to oranize the data when refractoring
    const POSPayments = selectedPaymentDate
      .filter((transaction) => transaction["Channel"]?.toUpperCase() !== "ECOM")
      .filter((card) => card["Card Type"]?.toUpperCase().includes("VISA")); // added by user
    const ECOMPayments = selectedPaymentDate
      .filter((transaction) => transaction["Channel"]?.toUpperCase() === "ECOM")
      .filter((card) => card["Card Type"]?.toUpperCase().includes("VISA")); //added by user

    const dayToMs = 24 * 60 * 60 * 1000;
    const parsedPaymentDate = (+excelPaymentDate - 25569) * dayToMs;
    if (!parsedPaymentDate) return;
    const manualDate = new Date(parsedPaymentDate - 3 * dayToMs);
    const MicrosDate = new Date(parsedPaymentDate - 1 * dayToMs);
    manualDate.setHours(0, 0, 0, 0);
    MicrosDate.setHours(0, 0, 0, 0);
    const Manual = journal.#sheet
      .filter(
        (manualPosting) =>
          new Date(manualPosting["BUSINESS_DATE"])?.toDateString() ==
          manualDate?.toDateString(),
      )
      .filter((transaction) =>
        transaction["TRX_DESC"]?.toUpperCase().includes("MANUAL"),
      )
      .filter(
        (type) =>
          !(
            type["REMARK"]?.toUpperCase().includes("LINK") ||
            type["CREDIT_CARD_SUPPLEMENT"]?.toUpperCase().includes("LINK") ||
            type["REFERENCE"]?.toUpperCase().includes("LINK")
          ),
      );

    const Micros = journal.#sheet
      .filter(
        (POSPosting) =>
          new Date(POSPosting["BUSINESS_DATE"])?.toDateString() ==
          MicrosDate?.toDateString(),
      )
      .filter((transaction) =>
        transaction["TRX_DESC"]?.toUpperCase().includes("POS"),
      );
      const previousDays = {
        previousPayments: [],
        previousPostings: [],
      }

    return { Manual, Micros, POSPayments };
  }

  static performRecon(journal, merchant,) {
    const reconData = workbook.getTransactions(journal, merchant, 46269);
    if (!reconData) {
      throw new Error(`couldnt get Recon data`);
    }
    const journalTransactions = [...reconData.Manual, ...reconData.Micros];
    const result = {
      usedJournalTransactions: new Set(),
      matchingTransactions: [],
      unmatchingPayments: [],
    };
    reconData.POSPayments.forEach((payment) => {
      const revenueIndex = journalTransactions.findIndex(
        (revenue, index) =>
          !result.usedJournalTransactions.has(index) &&
          +revenue["PRINT_CASHIER_CREDIT"] === +payment["Transaction Amount"],
      );

      if (revenueIndex === -1) {
        result.unmatchingPayments.push(payment);
      } else {
        result.usedJournalTransactions.add(revenueIndex);
        result.matchingTransactions.push({
          payment,
          revenue: journalTransactions[revenueIndex],
        });
      }
    });
    result.unmatchingJournal = journalTransactions.filter(
      (_, index) => !result.usedJournalTransactions.has(index),
    );

    console.log(result);

    
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

// allocate any corrections
// i need to carry over unreconciled transactions
  //could be done by inheritance 
