function openTab(tabName) {
    var i;
    var tabContents = document.getElementsByClassName("tab-contents");
    var tabLinks = document.getElementsByClassName("tab-links");
    for (i = 0; i < tabContents.length; i++) {
        tabContents[i].classList.remove("active-tab");
    }
    for (i = 0; i < tabLinks.length; i++) {
        tabLinks[i].classList.remove("active-link");
    }
    document.getElementById(tabName).classList.add("active-tab");
    event.currentTarget.classList.add("active-link");
}
    const scriptURL = 'https://script.google.com/macros/s/YOUR_GOOGLE_SCRIPT_URL_HERE/exec';
    const form = document.forms['submit-to-google-sheet'];
    const msg = document.getElementById("msg");

    form.addEventListener('submit', e => {
        e.preventDefault();
        fetch(scriptURL, { method: 'POST', body: new FormData(form)})
            .then(response => {
                msg.innerHTML = "Message sent successfully!";
                setTimeout(() => { msg.innerHTML = ""; }, 5000);
                form.reset();
            })
            .catch(error => {
                msg.innerHTML = "Message sending failed!";
                setTimeout(() => { msg.innerHTML = ""; }, 5000);
                console.error('Error!', error.message)
            });
    });

const sheetName = 'Sheet1';
const scriptProp = PropertiesService.getScriptProperties();

function doPost(e) {
    const lock = LockService.getScriptLock();
    lock.tryLock(10000);

    try {
        const doc = SpreadsheetApp.openById(scriptProp.getProperty('sheet_id'));
        const sheet = doc.getSheetByName(sheetName);

        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const nextRow = sheet.getLastRow() + 1;

        const newRow = headers.map(header => header === 'Timestamp' ? new Date() : e.parameter[header]);
        sheet.getRange(nextRow, 1, 1, newRow.length).setValues([newRow]);

        return ContentService.createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
            .setMimeType(ContentService.MimeType.JSON);
    } catch (e) {
        return ContentService.createTextOutput(JSON.stringify({ 'result': 'error', 'error': e }))
            .setMimeType(ContentService.MimeType.JSON);
    } finally {
        lock.releaseLock();
    }
}

function setup() {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    scriptProp.setProperty('sheet_id', doc.getId());
}
