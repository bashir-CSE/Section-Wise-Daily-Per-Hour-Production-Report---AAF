/**
 * Serves the web application's HTML interface.
 */
function doGet(e) {
	// Create an HTML template from the 'index.html' file.
	const htmlTemplate = HtmlService.createTemplateFromFile('index');

	// Fetch the dynamic data for the dropdowns.
	const data = getDropdownData();
	htmlTemplate.sections = data.sections;
	htmlTemplate.items = data.items;
	htmlTemplate.times = data.times;

	// Build and return the HTML page.
	return htmlTemplate
		.evaluate()
		.setTitle('Hourly Production Report')
		.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Reads the dropdown options from the 'Settings' sheet in your spreadsheet.
 */
function getDropdownData() {
	const ss = SpreadsheetApp.openById(
		'1iReWL_RZgjPHi8OBSguUDjzMF0Qr_Nx6E0ZndJSSRgk'
	);
	const settingsSheet = ss.getSheetByName('Settings');

	// Read data from columns A, B, and C, starting from the second row.
	const sections = settingsSheet
		.getRange('A2:A')
		.getValues()
		.flat()
		.filter(String);
	const items = settingsSheet
		.getRange('B2:B')
		.getValues()
		.flat()
		.filter(String);
	const times = settingsSheet
		.getRange('C2:C')
		.getValues()
		.flat()
		.filter(String);

	return { sections, items, times };
}

/**
 * Saves the submitted form data into the 'Hourly Report' sheet.
 * This function is called from the client-side JavaScript.
 */
function saveData(formData) {
	try {
		const ss = SpreadsheetApp.openById(
			'1iReWL_RZgjPHi8OBSguUDjzMF0Qr_Nx6E0ZndJSSRgk'
		);
		const reportSheet = ss.getSheetByName('Response');

		// Check if the sheet was found. If not, throw an error.
		if (!reportSheet) {
			throw new Error(
				"Sheet 'Response' not found. Please check the sheet name is correct and has no extra spaces."
			);
		}

		const timestamp = Utilities.formatDate(
			new Date(),
			Session.getScriptTimeZone(),
			'dd/MM/yyyy'
		);

		// If the report sheet is empty, add headers first.
		if (reportSheet.getLastRow() === 0) {
			reportSheet.appendRow([
				'Timestamp',
				'Section Name',
				'Time',
				'Item Name',
				'Qty',
			]);
		}

		// Loop through each item submitted and add it as a new row.
		formData.items.forEach((item) => {
			reportSheet.appendRow([
				timestamp,
				formData.sectionName,
				formData.time,
				item.itemName,
				item.qty,
			]);
		});

		return { status: 'success', message: 'Data saved successfully!' };
	} catch (error) {
		Logger.log(error); // Log errors for easier debugging.
		return { status: 'error', message: error.message };
	}
}
