"use strict";

const DOM ={
  pastTransactionsBtn: document.getElementById('previousDataFile')
}
class workbook {
  #data;
  #sheet;
  static previousDays;
  static deletedTransactions = {
    payments: [],
    postings: [],
  };
  static async loadPreviousDays(file) {
    if (!(file instanceof File)) {
      throw new Error("please select a JSON file");
    }
    let data;
    try {
      const text = await file.text();
      data = await JSON.parse(text);
    } catch (e) {
      throw new Error(`couldn't read the file's content: ${e.message}`);
    }

    if (
      !Array.isArray(data.previousPayments) ||
      !Array.isArray(data.previousPostings)
    ) {
      throw new Error("the file must contain previous transactions arrays");
    }
    workbook.previousDays = data;
  }

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

  static getTransactions(journal, merchant, paymentDate) {
    if (!journal.#sheet) {
      throw new Error(`recon sheet is unavaliable`);
    }
    const selectedPaymentDate = merchant.#sheet.filter(
      (requestedDay) => requestedDay["Posting Date DD:MM:YYYY"] == paymentDate,
    );
    if (!selectedPaymentDate) {
      throw new Error(`could not get date: ${selectedPaymentDate}`);
    }
    const POSPayments = selectedPaymentDate
      .filter((transaction) => transaction["Channel"]?.toUpperCase() === "POS")
      .filter((card) => card["Card Type"]?.toUpperCase().includes("VISA")); // added by user
    const ECOMPayments = selectedPaymentDate
      .filter((transaction) => transaction["Channel"]?.toUpperCase() === "ECOM")
      .filter((card) => card["Card Type"]?.toUpperCase().includes("VISA")); //added by user

    const ECOMRefunds = selectedPaymentDate.filter(
      (debitAmount) =>
        debitAmount["Transaction Type"]?.toUpperCase().includes("CREDIT") &&
        debitAmount["Channel"]?.toUpperCase() === "ECOM" &&
        debitAmount["Card Type"]?.toUpperCase().includes("VISA"),
    );

    POSPayments.push(...ECOMRefunds);

    const dayToMs = 24 * 60 * 60 * 1000;
    const parsedPaymentDate = (+paymentDate - 25569) * dayToMs;
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

    return { Manual, Micros, POSPayments };
  }

  static performRecon(journal, merchant, paymentDate) {
    const reconData = workbook.getTransactions(journal, merchant, paymentDate);
    if (!reconData) {
      throw new Error(`couldnt get Recon data`);
    }
    const allPayments = [
      ...this.previousDays.previousPayments,
      ...reconData.POSPayments,
    ];
    const allJournalTransactions = [...reconData.Manual, ...reconData.Micros];
    const result = {
      usedJournalTransactions: new Set(),
      matchingTransactions: [],
      unmatchingPayments: [],
    };

    const correctionIndex = new Set();

    allJournalTransactions.forEach((transaction, index) => {
      for (let i = index + 1; i < allJournalTransactions.length; i++) {
        if (correctionIndex.has(index) || correctionIndex.has(i)) break;
        if (transaction["REFERENCE"]?.toUpperCase().includes("REFUND")) break;
        const valueOne =
          allJournalTransactions[index][["PRINT_CASHIER_CREDIT"]];
        const valueTwo = allJournalTransactions[i][["PRINT_CASHIER_CREDIT"]];

        if (+valueOne + +valueTwo === 0) {
          correctionIndex.add(index);
          correctionIndex.add(i);
          break;
        }
      }
    });

    const journalAfterCorrections = allJournalTransactions.filter(
      (_, index) => !correctionIndex.has(index),
    );

    const reconJournal = [
      ...journalAfterCorrections,
      ...this.previousDays.previousPostings,
    ];
    // console.log({
    //   allJournalCount: allJournalTransactions.length,
    //   removedAsCorrections: correctionIndex.size,
    //   journalAfterCorrectionsCount: journalAfterCorrections.length,
    //   reconJournalCount: reconJournal.length,
    // });
    allPayments.forEach((payment) => {
      const revenueIndex = reconJournal.findIndex(
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
          revenue: reconJournal[revenueIndex],
        });
      }
    });
    result.unmatchingJournal = reconJournal.filter(
      (_, index) => !result.usedJournalTransactions.has(index),
    );
    this.previousDays.previousPayments = [...result.unmatchingPayments];
    this.previousDays.previousPostings = [...result.unmatchingJournal];
    

    console.log(result);
  }

  static removePosting(transactionIndex) {
    const unmatched = this.previousDays.previousPostings;
    this.deletedTransactions.postings.push(unmatched[transactionIndex]);
    unmatched.splice(transactionIndex, 1);
  }

  static removePayment(paymentsIndex) {
    const unmatched = this.previousDays.previousPayments;
    this.deletedTransactions.payments.push(unmatched[paymentsIndex]);
    unmatched.splice(paymentsIndex, 1);
  }
}

async function call() {
  const Journal = await workbook.create("./excel/Journal.xlsx");
  const Merchant = await workbook.create("./excel/Merchant September.xlsx");
  await Journal.getReconSheet();
  await Merchant.getReconSheet();

  workbook.performRecon(Journal, Merchant, 46269);
  workbook.performRecon(Journal, Merchant, 46270);
  workbook.performRecon(Journal, Merchant, 46271);
  workbook.performRecon(Journal, Merchant, 46272);
  workbook.performRecon(Journal, Merchant, 46273);
  workbook.performRecon(Journal, Merchant, 46274);
  workbook.performRecon(Journal, Merchant, 46275);
  workbook.performRecon(Journal, Merchant, 46276);
  workbook.removePosting(9);
  workbook.removePosting(8);
  workbook.removePosting(7);
  workbook.removePosting(6);
  workbook.removePosting(5);
  workbook.removePosting(4);
  workbook.removePosting(3);
  workbook.removePosting(2);
  workbook.removePosting(0);

  workbook.removePayment(8);
  workbook.removePayment(7);

  workbook.performRecon(Journal, Merchant, 46277);
  workbook.performRecon(Journal, Merchant, 46278);
}


document.addEventListener('change', async (e)=>{
  const target = e.target

  const pastTransactions = target.closest('[data-selection="pending-transactions"]')
  if(pastTransactions){
    const file = pastTransactions.files[0]
    if(!file){
      console.error('please select a JSON file first')
      return;
    }

  try{
    await workbook.loadPreviousDays(file)
    console.log(`${file.name} loaded`)
  } catch (error){
    console.log(error.message)
  }
  call();
  }
})

document.addEventListener('click', e =>{
const target = e.target

const pastButton = target.closest('[data-button="past-transactions"]')
if(pastButton){
  DOM.pastTransactionsBtn.click()
}

})