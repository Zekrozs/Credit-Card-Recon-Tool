"use strict";

const DOM = {
  pastTransactionsBtn: document.getElementById("previousDataFile"),
};
function includesText(value, text) {
  return (value ?? "").toUpperCase().includes(text.toUpperCase());
}
const EXCEL_EPOCH_OFFSET = 25569;
const MILLISECONDS_PER_DAY = 24 * 60 * 60 * 1000;
function excelSerialToDate(serialDate) {
  const serial = +serialDate;

  if (!Number.isFinite(serial)) {
    return null;
  }

  return new Date((serial - EXCEL_EPOCH_OFFSET) * MILLISECONDS_PER_DAY);
}

function toDateKey(value) {
  const date = value instanceof Date ? value : new Date(value);

  if (Number.isNaN(date.getTime())) {
    return null;
  }
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
  ].join("-");
}

const CARD_SCHEMES = {
  VISA: "visa",
  MASTER: "master",
  MADA: "mada",
};

function removeCorrections(postings) {
  const correctionIndex = new Set();
  postings.forEach((posting, index) => {
    if (correctionIndex.has(index)) return;
    for (let nextIndex = index + 1; nextIndex < postings.length; nextIndex++) {
      if (includesText(posting["REFERENCE"], "REFUND")) break;
      if (correctionIndex.has(nextIndex)) continue;
      const valueOne = posting[["PRINT_CASHIER_CREDIT"]];
      const valueTwo = postings[nextIndex][["PRINT_CASHIER_CREDIT"]];

      if (+valueOne + +valueTwo === 0) {
        correctionIndex.add(index);
        correctionIndex.add(nextIndex);
        break;
      }
    }
  });
  return postings.filter((_, index) => !correctionIndex.has(index));
}

function matchTransactions(postings, payments) {
  let usedJournalTransactions = new Set();
  const matchingTransactions = [];
  const unmatchingPayments = [];
  payments.forEach((payment) => {
    const revenueIndex = postings.findIndex(
      (revenue, index) =>
        !usedJournalTransactions.has(index) &&
        +revenue["PRINT_CASHIER_CREDIT"] === +payment["Transaction Amount"],
    );

    if (revenueIndex === -1) {
      unmatchingPayments.push(payment);
    } else {
      usedJournalTransactions.add(revenueIndex);
      matchingTransactions.push({
        payment,
        revenue: postings[revenueIndex],
      });
    }
  });
 const  unmatchingJournal = postings.filter(
    (_, index) => !usedJournalTransactions.has(index),
  );

  return {
    matchingTransactions,
    unmatchingPayments,
    unmatchingJournal,
  };
}

class Workbook {
  #data;
  #sheet;
  static previousDays;
  static deletedTransactions = {
    payments: [],
    postings: [],
  };
  static previousDaysFileHandle;
  static async loadPreviousDays(file) {
    if (!file) {
      throw new Error("please select a JSON file");
    }
    let data;
    try {
      const read = await file.getFile();
      const text = await read.text();
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
    Workbook.previousDaysFileHandle = file;
    Workbook.previousDays = data;
    console.log(`${file.name} loaded`);
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

    return new Workbook(data);
  }
  getReconSheet() {
    if (!this.#data || !this.#data.SheetNames?.length) {
      throw new Error(`Workbook data is unavailable`);
    }
    const firstSheetName = this.#data.SheetNames[0];
    const worksheet = this.#data.Sheets[firstSheetName];
    this.#sheet = XLSX.utils.sheet_to_json(worksheet);
  }
  getRows() {
    if (!this.#sheet) {
      throw new Error("The reconciliation sheet has not been loaded");
    }
    return structuredClone(this.#sheet);
  }
  // paymentdate is added by user
  static getTransactions(
    journalInstance,
    merchantInstance,
    paymentDate,
    cardScheme = CARD_SCHEMES.VISA,
  ) {
    const journalRows = journalInstance.getRows();
    const merchantRows = merchantInstance.getRows();
    if (!journalRows || !merchantRows) {
      throw new Error(`one or more recon sheet is unavaliable`);
    }
    const selectedPaymentDate = merchantInstance.#sheet.filter(
      (requestedDay) => requestedDay["Posting Date DD:MM:YYYY"] == paymentDate,
    );
    if (selectedPaymentDate.length === 0) {
      throw new Error(`could not get date: ${selectedPaymentDate}`);
    }
    const posPaymentsForScheme = selectedPaymentDate.filter((card) =>
      includesText(card["Card Type"], cardScheme),
    ); // added by user
    const ECOMPaymentsForScheme = selectedPaymentDate
      .filter((transaction) => includesText(transaction["Channel"], "ECOM"))
      .filter((card) => includesText(card["Card Type"], cardScheme)); //added by user

    const POSPayments = posPaymentsForScheme.filter(
      (transaction) =>
        includesText(transaction["Channel"], "POS") ||
        (includesText(transaction["Channel"], "ECOM") &&
          includesText(transaction["Transaction Type"], "CREDIT")),
    );

    const parsedPaymentDate = excelSerialToDate(+paymentDate);
    if (!parsedPaymentDate) return;
    const manualSettelmentWindow = 3;
    const MicrosSettelmentWindow = 1;
    const manualPostingDate = new Date(
      parsedPaymentDate.getTime() -
        manualSettelmentWindow * MILLISECONDS_PER_DAY,
    );
    const microsPostingDate = new Date(
      parsedPaymentDate.getTime() -
        MicrosSettelmentWindow * MILLISECONDS_PER_DAY,
    );
    const Manual = journalInstance.#sheet
      .filter(
        (manualPosting) =>
          toDateKey(manualPosting["BUSINESS_DATE"]) ===
          toDateKey(manualPostingDate),
      )
      .filter((transaction) => includesText(transaction["TRX_DESC"], "MANUAL"))
      .filter(
        (type) =>
          !(
            includesText(type["REMARK"], "LINK") ||
            includesText(type["CREDIT_CARD_SUPPLEMENT"], "LINK") ||
            includesText(type["REFERENCE"], "LINK")
          ),
      );

    const Micros = journalInstance.#sheet
      .filter(
        (POSPosting) =>
          toDateKey(POSPosting["BUSINESS_DATE"]) ==
          toDateKey(microsPostingDate),
      )
      .filter((transaction) =>
        transaction["TRX_DESC"]?.toUpperCase().includes("POS"),
      );

    return { Manual, Micros, POSPayments };
  }

  static performRecon(journalInstance, merchantInstance, paymentDate) {
    const reconData = Workbook.getTransactions(
      journalInstance,
      merchantInstance,
      paymentDate,
    );
    if (!reconData) {
      throw new Error(`couldnt get Recon data`);
    }
    const allPayments = [
      ...this.previousDays.previousPayments,
      ...reconData.POSPayments,
    ];
    const allJournalTransactions = [...reconData.Manual, ...reconData.Micros];

    const reconcilableJournal = removeCorrections(allJournalTransactions);

    const reconJournal = [
      ...reconcilableJournal,
      ...this.previousDays.previousPostings,
    ];

    const reconResult = matchTransactions(reconJournal, allPayments);
    this.previousDays.previousPayments = [...reconResult.unmatchingPayments];
    this.previousDays.previousPostings = [...reconResult.unmatchingJournal];
    console.log(reconResult)
    return reconResult
  }

  static async saveData(pending) {
    const fileHandle = Workbook.previousDaysFileHandle;
    if (!fileHandle) {
      throw new Error("No previously loaded file is avaliable");
    }
    const permission = await fileHandle.requestPermission({
      mode: "readwrite",
    });

    if (permission !== "granted") {
      throw new Error("write permission was not granted");
    }

    const pastTransactions = {
      previousPostings: pending.unmatchingJournal,
      previousPayments: pending.unmatchingPayments,
    };

    const writable = await fileHandle.createWritable();
    try {
      await writable.write(JSON.stringify(pastTransactions, null, 2));
    } finally {
      await writable.close();
    }
  }

  static async backupData(pending) {
    const backupData = {
      previousPostings: pending.unmatchingJournal,
      previousPayments: pending.unmatchingPayments,
    };

    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");

    const backupHandle = await window.showSaveFilePicker({
      suggestedName: `pending-transactions-backup-${timestamp}.json`,
      types: [
        {
          description: "JSON file",
          accept: {
            "application/json": [".json"],
          },
        },
      ],
    });

    const writable = await backupHandle.createWritable();

    try {
      await writable.write(JSON.stringify(backupData, null, 2));
    } catch {
      throw new Error(`couldnt create a backup file`);
    } finally {
      await writable.close();
    }
    console.log(`${backupHandle.name} backed up`);
    return true;
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
  const Journal = await Workbook.create("./excel/Journal.xlsx");
  const Merchant = await Workbook.create("./excel/Merchant September.xlsx");
  Journal.getReconSheet();
  Merchant.getReconSheet();

  Workbook.performRecon(Journal, Merchant, 46269);
  Workbook.performRecon(Journal, Merchant, 46270);
  Workbook.performRecon(Journal, Merchant, 46271);
  Workbook.performRecon(Journal, Merchant, 46272);
  Workbook.performRecon(Journal, Merchant, 46273);
  Workbook.performRecon(Journal, Merchant, 46274);
  Workbook.performRecon(Journal, Merchant, 46275);
  Workbook.performRecon(Journal, Merchant, 46276);
  Workbook.removePosting(9);
  Workbook.removePosting(8);
  Workbook.removePosting(7);
  Workbook.removePosting(6);
  Workbook.removePosting(5);
  Workbook.removePosting(4);
  Workbook.removePosting(3);
  Workbook.removePosting(2);
  Workbook.removePosting(0);

  Workbook.removePayment(8);
  Workbook.removePayment(7);

  Workbook.performRecon(Journal, Merchant, 46277);
  const result = Workbook.performRecon(Journal, Merchant, 46278);
  const dataBackedUp = await Workbook.backupData(result);
  if (dataBackedUp) {
    await Workbook.saveData(result);
  }
}

document.addEventListener("click", async (e) => {
  const target = e.target;

  const pastTransactions = target.closest(
    '[data-selection="pending-transactions"]',
  );
  if (pastTransactions) {
    try {
      const [handle] = await window.showOpenFilePicker({
        multiple: false,
        types: [
          {
            description: "JSON file",
            accept: {
              "application/json": [".json"],
            },
          },
        ],
      });
      await Workbook.loadPreviousDays(handle);
      await call();
    } catch (error) {
      console.error(error.message);
    }
  }
});

// document.addEventListener('click', e =>{
// const target = e.target

// const pastButton = target.closest('[data-button="past-transactions"]')
// if(pastButton){
//   DOM.pastTransactionsBtn.click()
// }

// })
