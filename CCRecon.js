"use strict";

const DOM = {
  pastTransactionsBtn: document.getElementById("previousDataFile"),
};

const CARD_SCHEMES = {
  VISA: "visa",
  MASTER: "master",
  MADA: "mada",
};
let appState = {
  workBook: null,
  sessionStore: null,
  results: [],
  sessionType: CARD_SCHEMES.VISA,
  resultsWorkBook: null,
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

function compareMatchingAmounts(a, b) {
  return a.payment["Transaction Amount"] - b.payment["Transaction Amount"];
}

function comparePendingJournal(a, b) {
  const dateOne = toDateKey(a.BUSINESS_DATE);
  const dateTwo = toDateKey(b.BUSINESS_DATE);
  return new Date(dateOne).getTime() - new Date(dateTwo).getTime();
}

function comparePendingPayments(a, b) {
  const dateOne = toDateKey(
    excelSerialToDate(a["Transaction Date DD:MM:YYYY"]),
  );
  const dateTwo = toDateKey(
    excelSerialToDate(b["Transaction Date DD:MM:YYYY"]),
  );
  return new Date(dateOne).getTime() - new Date(dateTwo).getTime();
}

function sortMatching(matched) {
  return matched.sort(compareMatchingAmounts);
}

function sortPendingJournal(unmatchingJournal) {
  return unmatchingJournal.sort(comparePendingJournal);
}
function sortPendingPayments(unmatchingPayments) {
  return unmatchingPayments.sort(comparePendingPayments);
}

// function exportToExcel(result, date, cardScheme) {
//   const workbook = XLSX.utils.book_new();
//   const sortedResult = sortMatching(result);

//   const flattenedMatching = sortedResult.map((row) => ({
//     Date: toDateKey(
//       excelSerialToDate(row.payment["Transaction Date DD:MM:YYYY"]),
//     ),
//     Card_number: row.payment["Card Number"],
//     Authorization_Code: row.payment["Authorization Code"],
//     Channel: row.payment.Channel,
//     Payment_amount: row.payment["Transaction Amount"],
//     Posting_amount: row.revenue["PRINT_CASHIER_CREDIT"],
//     Variance:
//       +row.revenue["PRINT_CASHIER_CREDIT"] - +row.payment["Transaction Amount"],
//   }));

//   const flattenedUnmatchingPayments = result.unmatchingPayments.map((row) => ({
//     paymentDate: toDateKey(
//       excelSerialToDate(row["Transaction Date DD:MM:YYYY"]),
//     ),
//     Card_numberr: row["Card Number"],
//     Authorization_Codee: row["Authorization Code"],
//     Channell: row.Channel,
//     Paymentt_amount: row["Transaction Amount"],
//     "": "",
//   }));

//   const flattenedunmatchingRevenue = result.unmatchingJournal.map((row) => ({
//     posted_on: row[["BUSINESS_DATE"]],
//     posting: row["PRINT_CASHIER_CREDIT"],
//   }));

//   const flattenedData = [
//     ...flattenedMatching,
//     "",
//     "",
//     "",
//     ...flattenedUnmatchingPayments,
//     "",
//     ...flattenedunmatchingRevenue,
//   ];
//   console.log(flattenedData);
//   const worksheet = XLSX.utils.json_to_sheet(flattenedData);
//   XLSX.utils.book_append_sheet(
//     workbook,
//     worksheet,
//     toDateKey(excelSerialToDate(date)),
//   );
//   XLSX.writeFile(workbook, `CC Recon: ${cardScheme}.xlsx`);
// }

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
  const unmatchingJournal = postings.filter(
    (_, index) => !usedJournalTransactions.has(index),
  );

  return {
    matchingTransactions,
    unmatchingPayments,
    unmatchingJournal,
  };
}

class pendingTransactionsStore {
  #fileHandle;
  #data;
  #deletedTransactions = {
    payments: [],
    postings: [],
  };
  #ValidateIndex(index, items, itemType) {
    if (!Number.isInteger(index) || index < 0 || index >= items.length) {
      throw new Error(`invalid ${itemType}, index: ${index}`);
    }
  }

  constructor(fileHandle, data) {
    this.#fileHandle = fileHandle;
    this.#data = {
      previousPayments: [...data.previousPayments],
      previousPostings: [...data.previousPostings],
    };
  }
  static async load(fileHandle) {
    if (!fileHandle) {
      throw new Error("please Select a JSON file");
    }

    try {
      const file = await fileHandle.getFile();
      const text = await file.text();
      const data = await JSON.parse(text);

      if (
        !Array.isArray(data.previousPostings) ||
        !Array.isArray(data.previousPayments)
      ) {
        throw new Error(
          "file must contain previousPayments and previousTransactions arrays",
        );
      }

      console.log(`${file.name} loaded`);

      return new pendingTransactionsStore(fileHandle, data);
    } catch (error) {
      throw new Error(`couldn't load pending transactions: ${error.message}`);
    }
  }

  getPayments() {
    return structuredClone(this.#data.previousPayments);
  }

  getPostings() {
    return structuredClone(this.#data.previousPostings);
  }

  update(result) {
    this.#data.previousPayments = [...result.unmatchingPayments];
    this.#data.previousPostings = [...result.unmatchingJournal];
  }

  removePayment(index) {
    this.#ValidateIndex(index, this.#data.previousPayments, "payment");

    const [payment] = this.#data.previousPayments.splice(index, 1);
    this.#deletedTransactions.payments.push(payment);

    return payment;
  }

  removePosting(index) {
    this.#ValidateIndex(index, this.#data.previousPostings, "posting");

    const [posting] = this.#data.previousPostings.splice(index, 1);
    this.#deletedTransactions.postings.push(posting);

    return posting;
  }

  async save() {
    const permission = await this.#fileHandle.requestPermission({
      mode: "readwrite",
    });
    if (permission !== "granted") {
      throw new Error("write permission was denied");
    }

    const writable = await this.#fileHandle.createWritable();

    try {
      await writable.write(JSON.stringify(this.#data, null, 2));
    } finally {
      await writable.close();
    }
    console.log(`${this.#fileHandle.name} updated`);
  }

  async backUp() {
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");

    const backUpHandle = await window.showSaveFilePicker({
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

    const writable = await backUpHandle.createWritable();

    try {
      await writable.write(JSON.stringify(this.#data, null, 2));
    } finally {
      await writable.close();
    }
    console.log(`${backUpHandle.name} backed up`);
  }
}
class Workbook {
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
      throw new Error("could not read workbook data");
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
    console.log(this.#sheet);
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
    cardScheme,
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
      throw new Error(`could not get date: ${paymentDate}`);
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
      .filter((transaction) => includesText(transaction["TRX_DESC"], "POS"));

    return { Manual, Micros, POSPayments, paymentDate };
  }
}

class ReconciliationSession {
  #journal;
  #merchant;
  #store;
  #cardScheme;

  constructor(journal, merchant, store, cardScheme) {
    this.#journal = journal;
    this.#merchant = merchant;
    this.#store = store;
    this.#cardScheme = cardScheme;
  }

  static async open(journal, merchant, fileHandle, cardScheme) {
    const store = await pendingTransactionsStore.load(fileHandle);

    return new ReconciliationSession(journal, merchant, store, cardScheme);
  }

  perform(paymentDate) {
    const reconData = Workbook.getTransactions(
      this.#journal,
      this.#merchant,
      paymentDate,
      this.#cardScheme,
    );

    if (!reconData) {
      throw new Error("Could not get reconciliation data");
    }

    const allPayments = [
      ...this.#store.getPayments(),
      ...reconData.POSPayments,
    ];

    const currentPostings = [...reconData.Manual, ...reconData.Micros];

    const reconJournal = [
      ...removeCorrections(currentPostings),
      ...this.#store.getPostings(),
    ];

    const result = matchTransactions(reconJournal, allPayments);

    this.#store.update(result);
    console.log(result);
    return result;
  }

  removePosting(index) {
    return this.#store.removePosting(index);
  }

  removePayment(index) {
    return this.#store.removePayment(index);
  }
  save() {
    return this.#store.save();
  }

  backup() {
    return this.#store.backUp();
  }
  getSessionType() {
    return this.#cardScheme;
  }
}

class ExcelExporter {
  #fileHandle;
  #resultWorkbook;
  #cardScheme;

  constructor(fileHandle, resultWorkbook, cardScheme) {
    this.#fileHandle = fileHandle;
    this.#resultWorkbook = resultWorkbook;
    this.#cardScheme = cardScheme;
  }

  static async open(fileHandle, cardScheme) {
    if (!fileHandle) {
      throw new Error("please select a book to save the results");
    }

    try {
      const file = await fileHandle.getFile();
      const arrayBuffer = await file.arrayBuffer(file);
      const workBook = await XLSX.read(arrayBuffer);
      return new ExcelExporter(fileHandle, workBook, cardScheme);
    } catch (error) {
      ("could not read the results workbook");
    }
  }

  appendResult(result, date) {
    const sortedMatching = sortMatching(result.matchingTransactions);
    const sortedPendingJournal = sortPendingJournal(result.unmatchingJournal);
    const sortedPendingPayment = sortPendingPayments(result.unmatchingPayments);

    const matchingRows = sortedMatching.map((row) => ({
      payment_date: toDateKey(
        excelSerialToDate(row.payment["Transaction Date DD:MM:YYYY"]),
      ),
      card: row.payment["Card Number"],
      payment_Channel: row.payment.Channel,
      payment_amount: row.payment["Transaction Amount"],
      posting_amount: row.revenue["PRINT_CASHIER_CREDIT"],
      variance:
        +row.revenue["PRINT_CASHIER_CREDIT"] -
        +row.payment["Transaction Amount"],
    }));

    const unmatchedPaymentRows = sortedPendingPayment.map((row) => ({
      date: toDateKey(excelSerialToDate(row["Transaction Date DD:MM:YYYY"])),
      card_number: row["Card Number"],
      Authorization_Code: row["Authorization Code"],
      Channel: row.Channel,
      Payment_amount: +row["Transaction Amount"],
    }));

    const unmatchedJournalRows = sortedPendingJournal.map((row) => ({
      Room: row.ROOM ?? "No room was assigned",
      Posted_on: row["BUSINESS_DATE"],
      Posting: +row["PRINT_CASHIER_CREDIT"],
      Refrence: row.REFERENCE
    }));

    const maxRowCount = Math.max(
      matchingRows.length,
      unmatchedJournalRows.length,
      unmatchedPaymentRows.length,
    );

    const rows = Array.from({ length: maxRowCount }, (_, index) => ({
      ...(matchingRows[index] ?? {}),
      ...(unmatchedJournalRows[index] ?? {}),
      ...(unmatchedPaymentRows[index] ?? {}),
    }));

    const sheet = XLSX.utils.json_to_sheet(rows);
    const sheetName = toDateKey(excelSerialToDate(date));

    XLSX.utils.book_append_sheet(this.#resultWorkbook, sheet, sheetName);
  }

  async save() {
    if(!this.#resultWorkbook){
      throw new Error('could not read the results workbook')
    }
    const permission = await this.#fileHandle.requestPermission({
      mode: 'readwrite',
    });
    if (!permission) {
      throw new Error("write permission was denied");
    }

    const output = XLSX.write(this.#resultWorkbook, {
      bookType: "xlsx",
      type: "array",
    });

    const writable = await this.#fileHandle.createWritable();

    try {
      await writable.write(output);
    } finally {
      await writable.close();
    }
  }
}

async function getWorkBooks() {
  const Journal = await Workbook.create("./excel/Journal.xlsx");
  const Merchant = await Workbook.create("./excel/Merchant September.xlsx");
  Journal.getReconSheet();
  Merchant.getReconSheet();
  return { Journal, Merchant };
}

async function openSession(fileHandle, data) {
  if (!data) {
    throw new Error("Could not read workbook");
  }
  return await ReconciliationSession.open(
    data.Journal,
    data.Merchant,
    fileHandle,
    CARD_SCHEMES.VISA,
  );
}

document.addEventListener("click", async (e) => {
  const target = e.target;

  const uploadPendingButton = target.closest(
    '[data-selection="pending-transactions"]',
  );
  if (uploadPendingButton) {
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
      appState.sessionStore = await openSession(handle, appState.workBook);
    } catch (error) {
      console.error(error.message);
    }
  }
  const uploadWorkBookButton = target.closest('[data-selection="work-book"]');

  if (uploadWorkBookButton) {
    appState.workBook = await getWorkBooks();
    console.log("workbooks uploaded");
  }

  const reconcile = target.closest('[data-reconcile="start"]');

  if (reconcile) {
    if (!appState.sessionStore || !appState.workBook) {
      throw new Error("please upload all files");
    }
    const dates = [
      46269, 46270, 46271, 46272, 46273, 46274, 46275, 46276, 46277, 46278,
    ];
    dates.forEach((date) => {
      appState.results = appState.sessionStore.perform(date);
      appState.resultsWorkBook.appendResult(appState.results, date);
      console.log("Reconcilliation done. Result:");
      console.log(appState.results);
    });
  }

  const downloadResultButton = target.closest('[data-download="results"]');

  if (downloadResultButton) {
    appState.resultsWorkBook.save()
  }

  const uploadOutputButton = target.closest('[data-upload="output"]');

  if (uploadOutputButton) {
    try {
      const [outputHandle] = await window.showOpenFilePicker({
        multiple: false,
        types: [
          {
            description: "Excel workbook",
            accept: {
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
                [".xlsx"],
              "application/vnd.ms-excel": [".xls"],
              "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"],
            },
          },
        ],
      });

      appState.resultsWorkBook = await ExcelExporter.open(
        outputHandle,
        appState.sessionType,
      );

      console.log("Results workbook uploaded");
    } catch (error) {
      if (error.name !== "AbortError") {
        console.error(error.message);
      }
    }
  }

  // const downloadResult = target.closest('[data-download="results"]')

  // if(downloadResult){

  // }
});
