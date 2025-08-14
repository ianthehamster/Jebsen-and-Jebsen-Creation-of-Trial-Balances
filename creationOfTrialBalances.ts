function main(workbook: ExcelScript.Workbook) {
    const excludeSheetNames = ["Steps", "General", "HUB PIC", "MAS and Wise Exchange Rates", "Orbitax Entity Codes", "Sheet1", "Adjusted Trial Balance", "Trial Balance with Negative Tax Expenses"];
    const allSheets = workbook.getWorksheets();

    // Currency Conversion START
    const ratesSheet = workbook.getWorksheet("MAS and Wise Exchange Rates")
    const currencyCodesRange = ratesSheet.getRange("B10:L10")
    const currencyCodes = currencyCodesRange.getValues()[0]

    const sgdRatesRange = ratesSheet.getRange("B11:L11")
    const sgdRates = sgdRatesRange.getValues()[0]

    const currentCurrency: string = "MMK"

    const currencyIndex = currencyCodes.findIndex(code => code === currentCurrency)

    console.log(sgdRates[currencyIndex])
    // Currency Conversion END

    const tbHeader = ["Entity Code", "Entity Name", "Account Code", "Account Name", "Amount"];
    const allTrialBalanceRows: (string | number)[][] = [];

    for (let sheet of allSheets) {
        const sheetName = sheet.getName();

        // Skip excluded sheets and sheets that start with "Trial Balance"
        if (
            excludeSheetNames.includes(sheetName) ||
            sheetName.startsWith("Trial Balance")
        ) {
            continue;
        }

        const usedRange = sheet.getUsedRange();
        if (!usedRange) continue;

        const rowCount = usedRange.getRowCount();
        const dataRange = sheet.getRange(`A9:H${rowCount}`);
        const values = dataRange.getValues();

        if (values.length < 2) continue;

        const entityCode = sheetName.includes("-") ? sheetName.split("-")[0] : sheetName;
        const entityName = sheetName.includes("-") ? sheetName.split("-")[1] : sheetName;

        const filteredRows = values.slice(1).filter(row => row[1]);

        // Currency for that sheet
        const currentCurrency: string | number | boolean = filteredRows[0][3]
        // console.log(currentCurrency)
        const currentCurrencyIndex: number = currencyCodes.findIndex(code => code === currentCurrency)
        const currentCurrencyExchangeRate = sgdRates[currentCurrencyIndex]
        // console.log(typeof currentCurrencyIndex, typeof currentCurrencyExchangeRate)
        // console.log(currentCurrencyIndex, currentCurrencyExchangeRate)


        // if (currentCurrencyExchangeRate === undefined) {
        //     console.log(sheet.getName())
        // }

        const tbData = filteredRows.map(row => {
            const accountCode = row[1].toString();
            const accountName = row[2].toString().replace(/^\d+\s*/, "")
            const amount = Number(row[4]);

            const flipSignAccounts = ["4", "5", "6", "7"];
            const taxExpenseAccounts = ["72000", "73000", "73005", "73010", "73020"]
            const shouldFlip = flipSignAccounts.includes(accountCode[0]) && !taxExpenseAccounts.includes(accountCode);
            const finalAmount = shouldFlip ? -amount : amount;

            let accountCodeNumberType = Number(row[1])

            if (isNaN(accountCodeNumberType)) {
                const randomNumber = 100000 + Math.floor(Math.random() * 1000000)
                accountCodeNumberType = randomNumber
            }

            /** Need to create account codes for accounts that have no account codes */

            return [
                entityCode,
                entityName,
                accountCodeNumberType,
                row[2].toString().replace(/^\d+\s*/, ""),
                finalAmount * Number(currentCurrencyExchangeRate)
            ]
        });

        // Add to overall master list
        allTrialBalanceRows.push(...tbData);
    }

    // Add result to a new worksheet
    const finalResult = [tbHeader, ...allTrialBalanceRows];

    let existingSheet = workbook.getWorksheet("TB with Positive Tax Exp");
    if (existingSheet) {
        existingSheet.delete();
    }
    let consolidatedSheet = workbook.addWorksheet("TB with Positive Tax Exp");

    const numRows = finalResult.length;
    const numCols = tbHeader.length;
    const outputRange = consolidatedSheet.getRangeByIndexes(0, 0, numRows, numCols);
    outputRange.setValues(finalResult);

    // Mapping of Entity Codes with Orbitax 
    const generalSheet = workbook.getWorksheet("General")

    const generalRange = generalSheet.getUsedRange()
    const generalValues = generalRange.getValues()

    const consolidatedTBRange = consolidatedSheet.getUsedRange()
    const consolidatedTBValues = consolidatedTBRange.getValues()

    const entityCodeMap: Map<string, number> = new Map<string, number>()
    const entityNameMap: Map<string, string> = new Map<string, string>()

    for (let i = 1; i < generalValues.length; i++) {
        const orbitaxEntityCode = Number(generalValues[i][1])
        const orbitaxEntityName = generalValues[i][3].toString().trim()
        const jebsenEntityShortName = generalValues[i][6].toString().trim()

        if (orbitaxEntityCode && orbitaxEntityName) {
            entityCodeMap.set(jebsenEntityShortName, orbitaxEntityCode)
            entityNameMap.set(jebsenEntityShortName, orbitaxEntityName)
        }
    }

    for (let i = 1; i < consolidatedTBValues.length; i++) {
        const jebsenEntityCode = consolidatedTBValues[i][0].toString().trim()
        const jebsenEntityShortName = consolidatedTBValues[i][1].toString().trim()
        const mappedCode = entityCodeMap.get(jebsenEntityShortName) ?? `No entity code in Orbitax corresponding to the ${jebsenEntityCode} here`
        const mappedName = entityNameMap.get(jebsenEntityShortName) ?? `No entity code in Orbitax corresponding to the ${jebsenEntityShortName} here`
        consolidatedTBValues[i][0] = mappedCode
        consolidatedTBValues[i][1] = mappedName
    }

    consolidatedTBRange.clear()
    const updatedRangeWithMappedCodesAndNames = consolidatedSheet.getRangeByIndexes(0, 0, consolidatedTBValues.length, consolidatedTBValues[0].length)
    updatedRangeWithMappedCodesAndNames.setValues(consolidatedTBValues)
    // End of Mapping 

    // Creation of second TB with negative Tax expenses for GloBE Income Adjustments QNR
    const existingNegSheet = workbook.getWorksheet("TB with Negative Tax Exp");
    if (existingNegSheet) {
        existingNegSheet.delete();
    }

    const adjustedSheet = consolidatedSheet.copy(ExcelScript.WorksheetPositionType.after, consolidatedSheet);
    adjustedSheet.setName("TB with Negative Tax Exp");

    // Tax Expense Accounts 
    const taxExpenseAccountsToBeFlippedToNegative = ["72000", "73000", "73005", "73010", "73020"];

    const range = adjustedSheet.getUsedRange()
    const values = range.getValues()

    for (let i = 1; i < values.length; i++) {
        const row = values[i]
        const accountCode = row[2]?.toString().padStart(5, "0")
        const amount = Number(row[4])

        if (taxExpenseAccountsToBeFlippedToNegative.includes(accountCode)) {
            row[4] = -Math.abs(amount)
        } else {
            row[4] = amount
        }
    }

    range.setValues(values)
    // End of Creation of second TB
}
