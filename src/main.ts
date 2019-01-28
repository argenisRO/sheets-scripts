function onOpen(): void {
    SpreadsheetApp.getUi()
        .createMenu('Menu')
        .addItem('Show Menu', 'showMenu')
        .addToUi()
    showSidebar()
}

function showSidebar(): void {
    const html = HtmlService.createTemplateFromFile('src/index')
        .evaluate()
        .setTitle('Deadserious Panel')
        .setWidth(300)
    SpreadsheetApp.getUi().showSidebar(html)
}

/**
 * Gets the content of the provided
 * HTML file. Received as a string.
 *
 * @param {string} input file name in string form
 * @return {string} html output
 * @customfunction
 */
function include(filename: string): string {
    return HtmlService.createHtmlOutputFromFile(filename).getContent()
}

/**
 * Gets the time remaining till the deadline
 * Obtained from the console sheet
 *
 * @return {number} the seconds left between today and the deadline
 * @customfunction
 */
function Deadline(): number {
    const ss = SpreadsheetApp.getActiveSpreadsheet()
    const dayOfWeek: string = ss
        .getSheetByName('Console')
        .getRange(15, 4) // D15
        .getValue()
        .toString()
    const deadline: any = new Date()
    deadline.setDate(deadline.getDate() + ((getDay(dayOfWeek) + 7 - deadline.getDay()) % 7))
    return deadline
}

/**
 * Get the numerical value of a day from 1-7
 *
 * @param {string} input a date in string form (ex.'Monday', 'Thursday')
 * @return numerical value of day
 * @customfunction
 */
function getDay(day: string): number {
    const days = {
        Monday: 1,
        Tuesday: 2,
        Wednesday: 3,
        Thursday: 4,
        Friday: 5,
        Saturday: 6,
        Sunday: 7,
    }
    return days[day]
}
