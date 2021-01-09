const APP_NAME = "Automate"
const SHEET_NAME_PLACEHOLDERS = "Placeholders"
const SHEET_NAME_SETTINGS = "Settings"

const RANGE_NAME_TEMPLATE_URL = "B1"
const RNAGE_NAME_FOLDER_URL = "B2"

const PDFS_FOLDER_NAME = "PDFs"
const STATUS_SUCCESS = "Created"

function onOpen() {
    const ui = SpreadsheetApp.getUi()
    const menu = ui.createMenu(APP_NAME)
    menu.addItem("Create PDFs", "create")
    menu.addToUi()
}

function create() {
    new App().create()
}

class App {
    constructor() {
        this.ss = SpreadsheetApp.getActive()
        this.ui = SpreadsheetApp.getUi()
        this.wsPlaceholders = this.ss.getSheetByName(SHEET_NAME_PLACEHOLDERS)
        this.wsSettings = this.ss.getSheetByName(SHEET_NAME_SETTINGS)
        this.parentFolder = DriveApp.getFileById(this.ss.getId()).getParents().next()
        this.pdfsFolder = this.getPDFsFolder()
        this.docTemplate = this.getDocTemplate()
    }

    getDocTemplate() {
        const url = this.wsSettings.getRange(RANGE_NAME_TEMPLATE_URL).getDisplayValue()
        try {
            const doc = DocumentApp.openByUrl(url)
            return DriveApp.getFileById(doc.getId())
        } catch (e) {
            return e.message
        }
    }
    getPDFsFolder() {
        const url = this.wsSettings.getRange(RNAGE_NAME_FOLDER_URL).getDisplayValue()
        const id = url.split("/folders/")[1]
        try {
            return DriveApp.getFolderById(id)
        } catch (e) {
            return this.parentFolder.createFolder(PDFS_FOLDER_NAME)
        }
    }

    getPlaceholders() {
        const values = this.wsPlaceholders.getDataRange().getDisplayValues()
        const headers = values[0]
        const placeholders = values.map((v, rowIndex) => {
            const selected = v[0]
            const status = v[1]
            const items = {}
            let valid = false
            if (selected == "TRUE" && status != STATUS_SUCCESS) {
                valid = true
                v.forEach((cell, i) => {
                    const header = headers[i]
                    const match = header.match(/{{.+}}/)
                    if (match) items[match[0]] = cell
                })
            }
            return { items, selected, status, valid, rowIndex }
        })
        return { placeholders, values }
    }

    createPdf(items) {
        const copyFile = this.docTemplate.makeCopy(this.parentFolder).setName(`Doc Template Copy`)
        const copyDoc = DocumentApp.openById(copyFile.getId())
        const body = copyDoc.getBody()
        Object.keys(items).forEach(key => {
            const value = items[key]
            body.replaceText(key, value)
        })
        copyDoc.saveAndClose()
        const blob = DriveApp.getFileById(copyDoc.getId()).getAs("application/pdf").setName(this.docTemplate.getName())
        const pdf = this.pdfsFolder.createFile(blob)
        copyFile.setTrashed(true)
        return pdf
    }

    create() {
        // check if doc template is valid
        if (typeof this.docTemplate === "string") {
            this.ui.alert(
                APP_NAME,
                `Error: ${this.docTemplate}. Please check value in the cell "${RANGE_NAME_TEMPLATE_URL}" of the sheet "${SHEET_NAME_SETTINGS}"`,
                this.ui.ButtonSet.OK
            )
            return
        }

        // get placeholders and check if selected rows are valid
        const { placeholders, values } = this.getPlaceholders()
        const validItems = placeholders.filter(({ valid }) => valid)
        if (validItems.length === 0) {
            this.ui.alert(
                APP_NAME,
                `No valid row selected in the sheet "${SHEET_NAME_PLACEHOLDERS}". Make sure row is selected and status is not "${STATUS_SUCCESS}".`,
                this.ui.ButtonSet.OK
            )
            return
        }

        // create pdfs 
        validItems.forEach(({ rowIndex, items }) => {
            const pdf = this.createPdf(items)
            values[rowIndex][1] = `=HYPERLINK("${pdf.getUrl()}", "${STATUS_SUCCESS}")`
        })
        const statusValues = values.map(v => [v[1]])
        this.wsPlaceholders.getRange(1, 2, statusValues.length, statusValues[0].length).setValues(statusValues)
    }
}
