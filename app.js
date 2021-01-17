const APP_NAME = "Automate"
const SHEET_NAME_PLACEHOLDERS = "Placeholders"
const SHEET_NAME_SETTINGS = "Settings"

const RANGE_NAME_TEMPLATE_URL = "B1"
const RNAGE_NAME_FOLDER_URL = "B2"

const PDFS_FOLDER_NAME = "PDFs"
const STATUS_SUCCESS = "Success"

const COLUMN_INDEX_HEADERS = 1 // column B

function onOpen() {
    const ui = SpreadsheetApp.getUi()
    const menu = ui.createMenu(APP_NAME)
    menu.addItem("Create PDFs", "create")
    menu.addToUi()
}

function create() {
    const ss = SpreadsheetApp.getActive()
    const ui = SpreadsheetApp.getUi()
    try {
        ss.toast("Creating...", APP_NAME, 30)
        new App().create()
    } catch (e) {
        ui.alert(APP_NAME, e.message, ui.ButtonSet.OK)
    }
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

        const transposedValues = []
        const rows = values.length
        const cols = values[0].length
        for (let c = 0; c < cols; c++) {
            transposedValues[c] = []
            for (let r = 0; r < rows; r++) {
                transposedValues[c][r] = values[r][c]
            }
        }

        const headers = transposedValues[COLUMN_INDEX_HEADERS]

        const placeholders = transposedValues.map((v, rowIndex) => {
            const [selected, status, filename, email, subject, body] = v
            const htmlBody = body.split("\n").map(line => `<p>${line}</p>`).join("")
            const items = {}
            let valid = false
            if (selected == "TRUE" && status.indexOf(STATUS_SUCCESS) !== 0 && rowIndex > COLUMN_INDEX_HEADERS) {
                valid = true
                v.forEach((cell, i) => {
                    const header = headers[i]
                    const match = header.match(/{{.+}}/)
                    if (match) items[match[0]] = cell
                })
            }
            return { items, selected, status, valid, filename, email, subject, htmlBody, rowIndex }
        })
        return { placeholders, transposedValues }
    }

    createPdf(items, filename) {
        const copyFile = this.docTemplate.makeCopy(this.parentFolder).setName(`Doc Template Copy`)
        const copyDoc = DocumentApp.openById(copyFile.getId())
        const body = copyDoc.getBody()
        Object.keys(items).forEach(key => {
            const value = items[key]
            body.replaceText(key, value)
        })
        copyDoc.saveAndClose()
        const blob = DriveApp.getFileById(copyDoc.getId()).getAs("application/pdf").setName(filename || this.docTemplate.getName())
        const pdf = this.pdfsFolder.createFile(blob)
        copyFile.setTrashed(true)
        return pdf
    }

    sendEmail({ email, subject, options }) {
        try {
            GmailApp.sendEmail(email, subject, "", options)
            return STATUS_SUCCESS
        } catch (e) {
            return e.message
        }
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
        const { placeholders, transposedValues } = this.getPlaceholders()
        const validItems = placeholders.filter(({ valid }) => valid)
        if (validItems.length === 0) {
            this.ui.alert(
                APP_NAME,
                `No valid row selected in the sheet "${SHEET_NAME_PLACEHOLDERS}". Make sure row is selected and status is not "${STATUS_SUCCESS}"`,
                this.ui.ButtonSet.OK
            )
            return
        }

        // create pdfs 
        validItems.forEach(({ rowIndex, items, filename, email, subject, htmlBody }) => {
            const pdf = this.createPdf(items, filename)
            let isEmailSent = false
            if (email.indexOf("@") !== -1) {
                const options = {
                    htmlBody,
                    attachments: [pdf.getBlob()]
                }
                isEmailSent = this.sendEmail({ email, subject, options })
            }
            let status = ""
            if (isEmailSent === false) {
                status = `${STATUS_SUCCESS}: PDF Created`
            } else if (isEmailSent === STATUS_SUCCESS) {
                status = `${STATUS_SUCCESS}: PDF Created & Email Sent`
            } else {
                status = `${STATUS_SUCCESS}: PDF Created & ${isEmailSent}`
            }
            transposedValues[rowIndex][1] = status
            this.wsPlaceholders.getRange(3, rowIndex + 1).setValue(`=HYPERLINK("${pdf.getUrl()}", "${filename}")`)
        })
        const statusValues = transposedValues.map(v => v[1])
        this.wsPlaceholders.getRange(2, 1, 1, statusValues.length).setValues([statusValues])
    }
}
