function main(workbook: ExcelScript.Workbook) {
    // Sheets to process
    const sheetNames = ["Tax GL", "PL info", "BS & others"];

    // Create or clear Consolidated Trial Balance sheet
    let ctbSheet = workbook.getWorksheet("Consolidated Trial Balance");
    if (!ctbSheet) {
        ctbSheet = workbook.addWorksheet("Consolidated Trial Balance");
    } else {
        ctbSheet.getUsedRange()?.clear();
    }

    // Set headers
    const headers = ["Entity Code", "Entity Name", "Account Code", "Account Name", "Amount"];
    ctbSheet.getRange("A1:E1").setValues([headers]);

    let outputRow = 2;
    let accountCodeSeed = 1001; // starting arbitrary code
    const safeHarborAccounts = ["Total Revenue", "Net Profit Before Tax", "Tax Expense", "Tax Expense-current"]

    for (const name of sheetNames) {
        const sheet = workbook.getWorksheet(name);
        if (!sheet) continue;

        const values = sheet.getUsedRange().getValues();

        // 1. Locate row with "Country" in column C (index 2)
        let headerRow = -1;
        for (let r = 0; r < values.length; r++) {
            if (values[r][2] && values[r][2].toString().trim() === "Country") {
                headerRow = r;
                break;
            }
        }
        if (headerRow === -1) continue;

        // Extract account names from columns D onward (index 3+)
        const accountNames: string[] = [];
        for (let c = 3; c < values[headerRow].length; c++) {
            if (values[headerRow][c] && values[headerRow][c].toString().trim() !== "") {
                accountNames.push(values[headerRow][c].toString().trim());
            }
        }

        console.log(accountNames)

        // 2. Locate "LGROUP - LEGAL RBU" row in column B
        let lgroupRow = -1;
        for (let r = 0; r < values.length; r++) {
            if (values[r][1] && values[r][1].toString().includes("LGROUP - LEGAL RBU")) {
                lgroupRow = r;
                break;
            }
        }
        if (lgroupRow === -1) continue;

        // Entity rows start after skipping one row post "LGROUP"
        for (let r = lgroupRow + 2; r < values.length; r++) {
            const cellValue = values[r][1]; // Column B
            if (!cellValue || typeof cellValue !== "string" || !cellValue.match(/^E\d+\s*-\s*/)) continue;

            const [entityCode, entityName] = cellValue.split(" - ").map(s => s.trim());

            const entityCodeWithoutE = entityCode.replace(/^E/, "")

            // 3. For each account (data point), record the amount
            for (let i = 0; i < accountNames.length; i++) {
                const accountName = accountNames[i];
                const amount = values[r][3 + i] || 0; // amounts align under headers from column D onward
                let accountCode: number
              

                switch (accountName) {
                    case "Total Revenue":
                        accountCode = 40000
                        break;

                    case "Tax Expense":
                        accountCode = 72000
                        break;

                    case "Tax Expenses-current":
                        accountCode = 72001
                        break;

                    case "Net Profit Before Tax":
                        accountCode = 90000
                        break;

                    default:
                        accountCode = accountCodeSeed++
                        break;
                }

                // Write to Consolidated Trial Balance
                ctbSheet.getCell(outputRow - 1, 0).setValue(entityCodeWithoutE);
                ctbSheet.getCell(outputRow - 1, 1).setValue(entityName);
                ctbSheet.getCell(outputRow - 1, 2).setValue(accountCode);
                ctbSheet.getCell(outputRow - 1, 3).setValue(accountName);
                ctbSheet.getCell(outputRow - 1, 4).setValue(amount);

                outputRow++;
            }
        }
    }

    // Mapping from David's 2024 2 digit codes to the Orbitax entity codes
    const davidChuaExcelFile = workbook.getWorksheet("Entity Codes and Names V2")
    if (davidChuaExcelFile) {
        const jjGeneralTabValues = davidChuaExcelFile.getUsedRange().getValues()

        const jjCodesToCbCR: { "Orbitax Code": string; "J&J Code": string }[] = []

        for (let r = 0; r < jjGeneralTabValues.length; r++) {
            const orbitaxCode = jjGeneralTabValues[r][0]
            const jjCode = jjGeneralTabValues[r][1]
            const entityName = jjGeneralTabValues[r][2]

            if (jjCode && orbitaxCode) {
                jjCodesToCbCR.push({
                    "Orbitax Code": orbitaxCode.toString().trim(),
                    "J&J Code": jjCode.toString().trim()
                })
            }
        }
        console.log(jjCodesToCbCR)

        const ctbValues = ctbSheet.getUsedRange().getValues()



        for (let r = 1; r < ctbValues.length; r++) {
            const jjCode = ctbValues[r][0] ? ctbValues[r][0].toString().trim() : ""

            const mapping = jjCodesToCbCR.find(obj => obj["J&J Code"] === jjCode)
            if (mapping) {
                ctbSheet.getCell(r, 0).setValue(mapping["Orbitax Code"])
            }
        }

    }

}
