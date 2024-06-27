function onOpen() {
    // створюємо вкладку з двома кнопками
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('😺 Mono Menu')
        .addItem('💳 Завантажити нові транзакції', 'uploadAllTransactions')
        .addItem('❗️ Створити/перестворити табличку', 'initialCreate')
        .addToUi();
}

// отримуємо токен Монобанку
const MONO_TOKEN = getScriptSecret("MONO_TOKEN")

let columns = [
    "Джерело", "Баланс", "Сума транзакції", "Кешбек", "Опис",
    "Коментар", "Час транзакції", "Категорія"
]

let columnsWidths = [85, 75, 130, 75, 250, 130, 150, 90]

let floatColumns = ["Баланс", "Сума транзакції", "Кешбек"]
let textColumns = ["Опис", "Коментар"]
let datetimeColumns = ["Час транзакції"]

let categories = [
  "🍽️ Кафе і ресторани", "💅 Краса і здоровʼя", "🛒 Магазини", "👕 Одяг", "💃 Відпочинок і розваги",
  "🏠 Платежі і комісії", "🎁 Подарунки", "🚌 Проїзд", "🎗 Благодійність", "Інше"
]

let sources = ["Mono", "Готівка"]

// Функція для створення нової сторінки з дефолтними колонками та форматами
function initialCreate() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetName = "Усі транзакції";
    let oldSheet = ss.getSheetByName(sheetName);
    // видаляємо сторінку якщо вона вже існує
    if (oldSheet) {
        ss.deleteSheet(oldSheet);
    }
    // і створюємо нову сторінку
    let newSheet = ss.insertSheet(sheetName, 0);


    // Додаємо рядок з header-ом.
    let headerRowRange = newSheet.getRange(1, 1, 1, columns.length);
    headerRowRange.setValues([columns]);
    headerRowRange.setFontWeight("bold");
    newSheet.setFrozenRows(1);

    // Додаємо фільтри по колонках
    let dataRange = newSheet.getDataRange();
    dataRange.createFilter();

    // Отримуємо діапазон цієї таблиці для подальшої роботи
    let maxRows = newSheet.getMaxRows();
    let lastColumn = newSheet.getLastColumn();
    let range = newSheet.getRange(1, 1, maxRows, lastColumn);

    // Змінюємо кольорову схему таблиці
    range.applyRowBanding(SpreadsheetApp.BandingTheme.YELLOW);

    // Змінюємо ширину колонок
    for (const [index, width] of columnsWidths.entries()) {
        newSheet.setColumnWidth(index + 1, width);
    }

    // Змінюємо тип перенесення значень в клітинці
    range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

    // Створюємо спадне меню для колонки "Джерело"
    let sourceColumnIndex = columns.indexOf("Джерело") + 1;
    let sourceColumn = newSheet.getRange(2, sourceColumnIndex, maxRows); // start from 2 to ignore header
    let sourceRule = SpreadsheetApp.newDataValidation().requireValueInList(sources).build();
    sourceColumn.setDataValidation(sourceRule);

    // Створюємо спадне меню для колонки "Категорія"
    let catColumnIndex = columns.indexOf("Категорія") + 1;
    let catColumn = newSheet.getRange(2, catColumnIndex, maxRows); // start from 2 to ignore header
    let catRule = SpreadsheetApp.newDataValidation().requireValueInList(categories).build();
    catColumn.setDataValidation(catRule);

    // Приміняємо відповідні типи даних по колонкам
    applyFormating(floatColumns, newSheet, "#,##0.00");
    applyFormating(textColumns, newSheet, "@")
    applyFormating(datetimeColumns, newSheet, "ddd, dd.mm.yyyy, hh:mm");
}

// Допоміжна функція для того, щоб змінювати типи даних по колонках
function applyFormating(columnsToApply, sheet, format) {
    let ranges = columnsToApply.map(column => {
        let columnIndex = columns.indexOf(column) + 1;
        let columnRange = sheet.getRange(1, columnIndex, sheet.getMaxRows(), 1);
        return columnRange;
    });
    ranges.map(range => { range.setNumberFormat(format); });
}

// Змінна яка показує час останнього запиту на Mono API
var lastApiRequest;

// Функція, яка завантажує нові транзакції в табличку
function uploadAllTransactions() {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Усі транзакції");
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log(`Наявні колонки: ${headers}`)

    let to = Date.now();
    let from = getLatestTransactionTs() + 1001;
    let periods = getTimePeriods(from, to)

    periods
        .forEach(
            ([from, to]) => {

                let descriptionColumn = sheet.getRange(2, headers.indexOf("Опис")+1, sheet.getLastRow(), 1).getValues()
                let categoryColumn = sheet.getRange(2, headers.indexOf("Категорія")+1, sheet.getLastRow(), 1).getValues()

                let transactions = getTransactions(from, to)
                let transactionsCnt = transactions.length
                for (let step = transactionsCnt - 1; step >= 0; step--) {
                    var transaction = transactions[step]
                    // Loop through the rows from bottom to top to find the last row with the same description
                    // start from 1, since 0 is header
                    for (var i = 1; i < descriptionColumn.length; i++) {
                        Logger.log(descriptionColumn[i][0]+" "+categoryColumn[i][0])
                        if (descriptionColumn[i][0] == transaction.description) {
                            transaction.category = categoryColumn[i][0];
                            Logger.log(categoryColumn[i][0])
                            Logger.log(transaction)

                            break;
                        }
                    }

                    // записуємо дані в табличку в порядку, що відповідає header-у
                    let entry = headers
                        .map(col => transaction.columnMap().get(col))
                    Logger.log(entry)

                    try {
                        sheet
                            .insertRowBefore(2)
                            .getRange(2, 1, 1, entry.length)
                            .setValues([entry]);
                    } catch (e) {
                        sheet
                            .deleteRow(2)
                        throw e;
                    }
                }
            }
        )
}

function getLatestTransactionTs() {
    Logger.info("Отримуємо час останньої завантаженої транзакції Монобанку")
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Усі транзакції");
    let transactionsTable = sheet.getDataRange().getValues();
    let headers = transactionsTable.shift();

    let sourceIndex = headers.indexOf("Джерело");
    let timestampIndex = headers.indexOf("Час транзакції");

    var from = 0;
    // ітеруємось по рядкам поки не знайдемо транзакцію Монобанку. Беремо час цієї транзакції
    for (let step = 0; step < transactionsTable.length; step++) {
        let transactionTsCell = transactionsTable[step][timestampIndex]
        if (!transactionTsCell) { continue }

        let transactionTs = transactionTsCell.valueOf()
        if (transactionTs > from && transactionsTable[step][sourceIndex] === 'Mono') {
            from = transactionTs
            Logger.info(`Час останньої транзакції - ${new Date(from).toISOString()}`)
            break
        }
    }
    // якщо транзакцій Моно ще не було, то беремо дані за останні 30 днів
    if (from == 0) {
        let lastMonth = new Date(Date.now() - (30 * 24 * 60 * 60 * 1000)).getTime();
        Logger.info(`Останньої транзакції не знайдено, завантажуємо транзакції за останні 30 днів ${lastMonth}`)
        from = lastMonth;
    }
    return from
}

function getTimePeriods(fromRaw, toRaw) {
    // swap if needed
    let [from, to] = fromRaw < toRaw ? [fromRaw, toRaw] : [toRaw, fromRaw];

    Logger.info(`Розбиваємо період (${new Date(from).toISOString()}, ${new Date(to).toISOString()}) на проміжки не більші за 31 добу + 1 годину (2682000 секунд)`)
    // "Максимальний час, за який можливо отримати виписку — 31 доба + 1 година (2682000 секунд)" (c) документація
    const maxPeriodMillis = 2682000 * 1000
    const oneDayMillis = 24 * 60 * 60 * 1000

    var chunks = [];
    // якщо ми намагаємось отримати транзакції за період менший ніж maxPeriodMillis,то просто повертаємо цей період
    if (to - from < maxPeriodMillis) {
        chunks.push([from, to]);
    } else {
        // якщо період більший, то розбиваємо його на проміжки не більші за maxPeriodMillis - 1 день
        for (let chunkFrom = from; chunkFrom < to; chunkFrom += maxPeriodMillis - oneDayMillis) {
            chunkTo = Math.min(chunkFrom + maxPeriodMillis - oneDayMillis, to); // Перевіряємо, що chunkTo не більше ніж сам to
            chunks.push([chunkFrom, chunkTo]);
        }
    }

    let prettyChunks = chunks.map(([from, to]) => [new Date(from).toISOString(), new Date(to).toISOString()]);
    Logger.info(`Отримані проміжки ${prettyChunks}`)
    return chunks;
}

function getTransactions(from, to) {
    var transactions = [];
    var newFrom = from;
    var transactionsCnt;
    var isInitialRun;
    Logger.info(`Отримуємо транзакції за період (${new Date(newFrom).toISOString()}, ${new Date(to).toISOString()})`)

    if (!lastApiRequest) {
        lastApiRequest = Date.now();
        isInitialRun = true;
    }

    do {
        // "засинаємо" бо API має обмеження на кількість запитів в хвилину
        if (!isInitialRun) {
            let sleepTime = 61 * 1000 - (Date.now() - lastApiRequest)
            Logger.info(`Чекаємо ${Math.round(sleepTime / 1000)} секунд перед наступним запитом`)
            Utilities.sleep(sleepTime)
        }
        newTransactions = makeRequest(newFrom, to)
        lastApiRequest = Date.now()
        if (newTransactions.length == 0 ){break}
        newFrom = newTransactions.at(-1).time
        transactionsCnt = newTransactions.length
        transactions.push(newTransactions)
    } while (transactionsCnt == 500)

    return transactions.flat()
}

function makeRequest(from, to) {
    let account = 0
    let URL_STRING = `https://api.monobank.ua/personal/statement/${account}/${from}/${to}`;
    let options = {
        'method': 'get',
        'headers': { 'X-Token': MONO_TOKEN },
        'muteHttpExceptions': true
    };
    Logger.log(`Робимо запит: ${URL_STRING}`)

    let response = UrlFetchApp.fetch(URL_STRING, options);
    let responseCode = response.getResponseCode()
    let json = response.getContentText()

    if (responseCode == 429) {
        throw new Error('Забагато запитів за короткий проміжок часу. Почекайте 1 хвилину і спробуйте ще раз')
    } else if (responseCode >= 300) {
        throw new Error(`${responseCode}: ${json}`)
    }

    let transactions = JSON.parse(json).map(MonoTransaction.fromJSON);

    return transactions
}

function getScriptSecret(key) {
    let secret = PropertiesService.getScriptProperties().getProperty(key)
    if (!secret) throw Error(`Ключ ${key} не знайдено. Будь ласка, додайте його в "Властивості скрипта"`)
    return secret
}

class MonoTransaction {
    constructor({
                    time,
                    description,
                    amount,
                    cashbackAmount,
                    balance,
                    comment
                }
    ) {
        // переводимо epoch seconds в timestamp, а копійки в гривні
        this.time = new Date(time * 1000);
        this.amount = amount / 100;
        this.cashbackAmount = cashbackAmount / 100;
        this.description = description;
        this.comment = comment;
        this.balance = balance / 100;

        this.source = 'Mono';
        this.category = 'Інше';
    }

    columnMap(){
        return new Map([
            ["Джерело", this.source],
            ["Баланс", this.balance],
            ["Сума транзакції", this.amount],
            ["Кешбек", this.cashbackAmount],
            ["Опис", this.description],
            ["Коментар", this.comment],
            ["Час транзакції", this.time],
            ["Категорія", this.category],
        ])
    }

    static fromJSON(json) {
        return new MonoTransaction({
                time: json.time,
                description: json.description,
                amount: json.amount,
                cashbackAmount: json.cashbackAmount,
                balance: json.balance,
                comment: json.comment,
            }
        );
    }
}
