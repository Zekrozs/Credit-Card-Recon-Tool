"use strict";

class workbook {
  #data;
  #sheet;
  static previousDays = {
    previousPayments: [
      {
        "3D Secure": " ",
        "Account Number": "'262320203004",
        "Authorization Code": 141346,
        "Card Number": "444439******8178",
        "Card Type": "VISA Intl.",
        "CashBack Amount": 0,
        Channel: "POS",
        "Discount Amount": -2.5,
        "Merchant ID": "'452308220073",
        "Order Id": "X",
        "Payment ID": 280858890,
        "Posting Date DD:MM:YYYY": 46268,
        RRN: "'624521143602",
        "Reference 1": "'  ",
        "Reference 2": "'  ",
        "Terminal ID": "'6420918345230822",
        "Token Type": "CARD",
        "Total Pay Amount": 1600,
        "Transaction Amount": 100,
        "Transaction Date DD:MM:YYYY": 46267,
        "Transaction Time": 0.8851388888888889,
        "Transaction Type": "Purchase",
        "Txn SubType": "Purchase transactions",
        VAT: -0.38,
      },
      {
        "3D Secure": " ",
        "Account Number": "'262320203004",
        "Authorization Code": 240783,
        "Card Number": "478477******6697",
        "Card Type": "VISA",
        "CashBack Amount": 0,
        Channel: "POS",
        "Discount Amount": -8.84,
        "Merchant ID": "'452308220066",
        "Order Id": "X",
        "Payment ID": 280417458,
        "Posting Date DD:MM:YYYY": 46268,
        RRN: "'624416364771",
        "Reference 1": "'  ",
        "Reference 2": "'  ",
        "Terminal ID": "'6420930745230822",
        "Token Type": "CARD",
        "Total Pay Amount": 2943.48,
        "Transaction Amount": 570,
        "Transaction Date DD:MM:YYYY": 46266,
        "Transaction Time": 0.6922106481481481,
        "Transaction Type": "Purchase",
        "Txn SubType": "Purchase transactions",
        VAT: -1.33,
      },
      {
        "3D Secure": " ",
        "Account Number": "'262320203004",
        "Authorization Code": 59240,
        "Card Number": "445310******1675",
        "Card Type": "VISA",
        "CashBack Amount": 0,
        Channel: "POS",
        "Discount Amount": -13.95,
        "Merchant ID": "'452308220066",
        "Order Id": "X",
        "Payment ID": 280417458,
        "Posting Date DD:MM:YYYY": 46268,
        RRN: "'624416233870",
        "Reference 1": "'  ",
        "Reference 2": "'  ",
        "Terminal ID": "'6420930745230822",
        "Token Type": "CARD",
        "Total Pay Amount": 2943.48,
        "Transaction Amount": 900,
        "Transaction Date DD:MM:YYYY": 46266,
        "Transaction Time": 0.6830787037037037,
        "Transaction Type": "Purchase",
        "Txn SubType": "Purchase transactions",
        VAT: -2.09,
      },
      {
        "3D Secure": " ",
        "Account Number": "'262320203004",
        "Authorization Code": 976296,
        "Card Number": "460901******5050",
        "Card Type": "VISA Intl.",
        "CashBack Amount": 0,
        Channel: "POS",
        "Discount Amount": -36.84,
        "Merchant ID": "'452308220066",
        "Order Id": "X",
        "Payment ID": 280417458,
        "Posting Date DD:MM:YYYY": 46268,
        RRN: "'624414535366",
        "Reference 1": "'  ",
        "Reference 2": "'  ",
        "Terminal ID": "'6420930745230822",
        "Token Type": "CARD",
        "Total Pay Amount": 2943.48,
        "Transaction Amount": 1473.48,
        "Transaction Date DD:MM:YYYY": 46266,
        "Transaction Time": 0.6207523148148149,
        "Transaction Type": "Purchase",
        "Txn SubType": "Purchase transactions",
        VAT: -5.53,
      },
    ],
    previousPostings: [],
  };
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
      .filter((transaction) => transaction["Channel"]?.toUpperCase() !== "ECOM")
      .filter((card) => card["Card Type"]?.toUpperCase().includes("VISA")); // added by user
    const ECOMPayments = selectedPaymentDate
      .filter((transaction) => transaction["Channel"]?.toUpperCase() === "ECOM")
      .filter((card) => card["Card Type"]?.toUpperCase().includes("VISA")); //added by user

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

    allJournalTransactions.forEach((_, index) => {
      for (let i = index + 1; i < allJournalTransactions.length; i++) {
        const valueOne =
          allJournalTransactions[index][["PRINT_CASHIER_CREDIT"]];
        const valueTwo = allJournalTransactions[i][["PRINT_CASHIER_CREDIT"]];

        if (+valueOne + +valueTwo === 0) {
          correctionIndex.add(index);
          correctionIndex.add(i);
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
     console.log({
      allJournalCount: allJournalTransactions.length,
      removedAsCorrections: correctionIndex.size,
      journalAfterCorrectionsCount: journalAfterCorrections.length,
      reconJournalCount: reconJournal.length,
    });
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

    // console.log(result);
    // this.previousDays.previousPostings = [...result.unmatchingJournal];
    // this.previousDays.previousPayments = [...result.unmatchingPayments];
    // console.log(this.previousDays);
    // this.previousDays.previousPayments.forEach(payment => console.log(payment["Transaction Amount"]))
  }
}

async function call() {
  const Journal = await workbook.create("./excel/Journal.xlsx");
  const Merchant = await workbook.create("./excel/Merchant September.xlsx");
  await Journal.getReconSheet();
  await Merchant.getReconSheet();

  workbook.performRecon(Journal, Merchant, 46269);
  workbook.performRecon(Journal, Merchant, 46270);
  // workbook.performRecon(Journal, Merchant, 46271);
}
call();

// the files
// the payment date
// payment method

// allocate any corrections
// i need to carry over unreconciled transactions
//could be done by inheritance
