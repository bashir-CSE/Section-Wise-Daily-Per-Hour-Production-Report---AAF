/**
 * Serves the web application's HTML interface.
 */
function doGet(e) {
	// Create an HTML template from the 'index.html' file.
	const htmlTemplate = HtmlService.createTemplateFromFile("index");

	// Build and return the HTML page.
	return htmlTemplate
		.evaluate()
		.setTitle("Hourly Production Report")
		.addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

/**
 * Reads the dropdown options from the 'Settings' sheet in your spreadsheet.
 * This function is used server-side to pre-populate the HTML template.
 */
function getDropdownData() {
	const ss = SpreadsheetApp.openById(
		"1iReWL_RZgjPHi8OBSguUDjzMF0Qr_Nx6E0ZndJSSRgk"
	);
	const settingsSheet = ss.getSheetByName("Settings");

	// Read data from columns A, B, and C, starting from the second row.
	const sections = settingsSheet
		.getRange("A2:A")
		.getValues()
		.flat()
		.filter(String);
	const items = settingsSheet
		.getRange("B2:B")
		.getValues()
		.flat()
		.filter(String);
	const times = settingsSheet
		.getRange("C2:C")
		.getValues()
		.flat()
		.filter(String);

	return { sections, items, times };
}

/**
 * Reads the dropdown options from the 'Settings' sheet in your spreadsheet.
 * This function is called asynchronously from the client-side.
 * @returns {Object} An object containing sections, items, and times arrays.
 */
function getDropdownDataForClient() {
	const ss = SpreadsheetApp.openById(
		"1iReWL_RZgjPHi8OBSguUDjzMF0Qr_Nx6E0ZndJSSRgk"
	);
	const settingsSheet = ss.getSheetByName("Settings");

	// Read data from columns A, B, and C, starting from the second row.
	const sections = settingsSheet
		.getRange("A2:A")
		.getValues()
		.flat()
		.filter(String);
	const items = settingsSheet
		.getRange("B2:B")
		.getValues()
		.flat()
		.filter(String);
	const times = settingsSheet
		.getRange("C2:C")
		.getValues()
		.flat()
		.filter(String);

	// Log the data to help with debugging.
	Logger.log("Sections: " + sections);
	Logger.log("Items: " + items);
	Logger.log("Times: " + times);

	return { sections, items, times };
}

/**
 * Retrieves all past reports from the 'Response' sheet.
 * @returns {Array<Array<any>>} A 2D array of report data.
 */
function getPastReports() {
	const ss = SpreadsheetApp.openById(
		"1iReWL_RZgjPHi8OBSguUDjzMF0Qr_Nx6E0ZndJSSRgk"
	);
	const reportSheet = ss.getSheetByName("Response");

	if (!reportSheet) {
		throw new Error(
			"Sheet 'Response' not found. Please check the sheet name is correct and has no extra spaces."
		);
	}

	const data = reportSheet.getDataRange().getDisplayValues();
	if (data.length <= 1) {
		return [];
	}

	// Return the raw data and let the client handle formatting.
	// This is more robust.
	return data.slice(1);
}

/**
 * Saves the submitted form data into the 'Hourly Report' sheet.
 * This function is called from the client-side JavaScript.
 */
function saveData(formData) {
	try {
		const ss = SpreadsheetApp.openById(
			"1iReWL_RZgjPHi8OBSguUDjzMF0Qr_Nx6E0ZndJSSRgk"
		);
		const reportSheet = ss.getSheetByName("Response");

		// Check if the sheet was found. If not, throw an error.
		if (!reportSheet) {
			throw new Error(
				"Sheet 'Response' not found. Please check the sheet name is correct and has no extra spaces."
			);
		}

		const timezone = ss.getSpreadsheetTimeZone();
		const timestamp = Utilities.formatDate(new Date(), timezone, "dd/MM/yyyy");

		// If the report sheet is empty, add headers first.
		if (reportSheet.getLastRow() === 0) {
			reportSheet.appendRow([
				"Timestamp",
				"Section Name",
				"Time",
				"Item Name",
				"Qty",
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

		return { status: "success", message: "Data saved successfully!" };
	} catch (error) {
		Logger.log(error); // Log errors for easier debugging.
		return { status: "error", message: error.message };
	}
}
