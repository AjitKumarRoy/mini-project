const { google } = require('googleapis');
const User = require('../models/User');
const { oauth2Client } = require('../config/google');
const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

// Helper function to set credential for a given user
const setUserCredentials = async (userId) => {
    const user = await User.findById(userId);
    if (!user) throw new Error('User not found');
    oauth2Client.setCredentials({
        access_token: user.accessToken,
        refresh_token: user.refreshToken
    });

    return user;
};

// Create a new spreadsheet
const createSpreadSheet = async (req, res, next) => {
    try {
        // set user credentials
        await setUserCredentials(req.user.id);

        // create a new spreadsheet
        const resource = {
            properties: { title: req.body.title || 'New Sheet' }
        };

        const response = await sheets.spreadsheets.create({
            resource,
            fields: 'spreadsheetId'
        });
        res.json({ 'spreadsheetId': response.data.spreadsheetId });
    } catch (error) {
        next(error);
    }
};

// Rename an existing spreadsheet
const renameSpreadSheet = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { newTitle } = req.body;

        if (!newTitle) {
            return res.status(400).json({success: false, message: 'newTitle is required.'});
        }
    
        // prepare request to rename spreadsheet
        const requestBody = {
            requests: [{
                updateSpreadsheetProperties: {
                    properties: {title: newTitle},
                    fields: 'title'
                }
            }]
        };

        const response = await sheets.spreadsheets.batchUpdate({
            spreadsheetId: sheetId,
            resource: requestBody
        });

        res.status(200).json({
            success: true,
            message: 'Spreadsheet renamed successfully',
            newTitle
        });
    } catch(error) {
        next(error);
    }
};

// Create a new sheet within an spreadsheet
const createSheet = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { sheetName, options } = req.body;
        if (!sheetName) {
            return res.status(400).json({ 'success': false, 'message': 'sheetName is required.' });
        }

        // prepare request
        const requestBody = {
            requests: [{
                addSheet: {
                    properties: {
                        title: sheetName,
                        ...options // merge optional properties if provided
                    }
                }
            }]
        };

        const response = await sheets.spreadsheets.batchUpdate({
            spreadsheetId: sheetId,
            resource: requestBody
        });

        res.json(response.data);
    } catch (error) {
        next(error);
    }
};

// Rename a sheet 
const renameSheet = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { sheetName, newSheetName } = req.body;

        if (!sheetName || !newSheetName) {
            return res.status(400).json({ success: false, message: 'Both sheetName and newSheetName are required' });
        }

        const request = {
            spreadsheetId: sheetId,
            requests: [
                {
                    updateSheetProperties: {
                        properties: { sheetId: sheetName, title: newSheetName },
                        fields: 'title'
                    }
                }
            ]
        };

        await sheets.spreadsheets.batchUpdate(request);
        res.status(200).json({ success: true, message: 'Sheet renamed successfully' });
    } catch (error) {
        next(error);
    }
};


// Fetch all the data from a sheet, default(Sheet1)
const getSheet = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { sheetName } = req.body || 'Sheet1';
        // Fetch data from the spreadsheet (example: first sheet, all values)
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: sheetId,
            range: sheetName
        });
        res.json(response.data);
    } catch (error) {
        next(error);
    }
};

// Update or enter data in a sheet
const updateSheet = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { sheetName, range, values } = req.body; // expecting range and new values
        const resource = { values };
        const response = await sheets.spreadsheets.values.update({
            spreadsheetId: sheetId,
            range: `${sheetName}!${range}`,
            valueInputOption: 'RAW',
            resource
        });
        res.json({ updatedCells: response.data.updatedCells });
    } catch (error) {
        next(error);
    }
};

// Delete a spreadsheet
const deleteSpreadSheet = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;

        const drive = google.drive({ version: 'v3', auth: oauth2Client });
        await drive.files.delete({ fileId: sheetId });
        res.json({ 'message': 'Spreadsheet deleted successfully' });
    } catch (error) {
        next(error);
    }
};

// Delete a sheet within an spreadsheet
const deleteSheet = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { sheetName } = req.body;

        // Get spreadsheet details to find the sheet ID
        const sheetDetails = await sheets.spreadsheets.get({
            spreadsheetId: sheetId
        });

        const sheet = sheetDetails.data.sheets.find(s => s.properties.title === sheetName);
        if (!sheet) return res.status(404).json({ 'message': "Sheet not found" });

        await sheets.spreadsheets.batchUpdate({
            spreadsheetId: sheetId,
            requestBody: {
                requests: [{ deleteSheet: { sheetId: sheet.properties.sheetId } }]
            }
        });

        res.json({ 'message': `Sheet ${sheetName} deleted successfully.` });
    } catch (error) {
        next(error);
    }
};

// Append data in a sheet
const appendData = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { sheetName, values } = req.body;

        if (!sheetName) {
            return res.status(400).json({success: false, message: 'sheetName is required.'});
        }
        if (!values || !Array.isArray(values)) {
            return res.status(400).json({success: false, message: 'value must be an array of arrays'});
        }

        // Define request
        const request = {
            spreadsheetId: sheetId,
            range: `${sheetName}!A1`, // Appends starting from column A
            valueInputOption: 'RAW', // RAW or USER_ENTERED
            insertDataOption: 'INSERT_ROWS', // Insert new rows
            resource: { values }
        };

        // Append data to the sheet
        const response = await sheets.spreadsheets.values.append(request);

        res.status(200).json({
            success: true,
            message: 'Data appended successfully',
            updatedRange: response.data.updates.updatedRange
        });
    } catch(error) {
        next(error);
    }
};

// Clear Data from a Specified Range
const clearDataFromSheet = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { sheetName, range } = req.body;

        if (!sheetName || !range) {
            return res.status(400).json({success: false, message: 'sheetName and range are required.'});
        }

        const response = await sheets.spreadsheets.values.clear({
            spreadsheetId: sheetId,
            range: `${sheetName}!${range}`
        });

        res.status(200).json({success: true, message: 'Data cleared successfully.'});
    } catch(error) {
        next(error);
    }
};

// Delete rows from a sheet 
const deleteRowsFromSheet = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { sheetName, startRow, endRow  } = req.body;

        if (!sheetName || startRow === undefined || endRow === undefined) {
            return res.status(400).json({ success: false, message: 'sheetName, startRow, and endRow are required' });
        }

        const request = {
            spreadsheetId: sheetId,
            requests: [
                {
                    deleteDimension: {
                        range: {
                            sheetId: sheetName,
                            dimension: 'ROWS',
                            startIndex: startRow - 1,  // 0-based index
                            endIndex: endRow  // 0-based index
                        }
                    }
                }
            ]
        };

        await sheets.spreadsheets.batchUpdate(request);
        res.status(200).json({ success: true, message: 'Rows deleted successfully' });
    } catch (error) {
        next(error);
    }
};

// Delete a Column from a sheet
const deleteColumnFromSheet = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { sheetName, columnIndex } = req.body;

        if (!sheetName || columnIndex === undefined) {
            return res.status(400).json({ success: false, message: 'sheetName and columnIndex are required' });
        }

        const request = {
            spreadsheetId: sheetId,
            requests: [
                {
                    deleteDimension: {
                        range: {
                            sheetId: sheetName,
                            dimension: 'COLUMNS',
                            startIndex: columnIndex - 1,
                            endIndex: columnIndex
                        }
                    }
                }
            ]
        };

        await sheets.spreadsheets.batchUpdate(request);
        res.status(200).json({ success: true, message: 'Column deleted successfully' });
    } catch (error) {
        next(error);
    }
};

// List all sheets with metadata
const listSheetsWithMetadata = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;

        const response = await sheets.spreadsheets.get({
            spreadsheetId: sheetId
        });

        const sheetsMetadata = response.data.sheets.map(sheet => ({
            title: sheet.properties.title,
            sheetId: sheet.properties.sheetId,
            rowCount: sheet.properties.gridProperties.rowCount,
            columnCount: sheet.properties.gridProperties.columnCount
        }));

        res.status(200).json({ success: true, sheets: sheetsMetadata });
    } catch (error) {
        next(error);
    }
};

// List all spreadsheets
const listAllSpreadsheets = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);

        const response = await sheets.spreadsheets.list({
        });

        res.status(200).json({ success: true, spreadsheets: response.data });
    } catch (error) {
        next(error);
    }
};

// Get overall spreadsheets metadata
const getSpreadsheetMetadata = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;

        const response = await sheets.spreadsheets.get({
            spreadsheetId: sheetId
        });

        res.status(200).json({ success: true, metadata: response.data });
    } catch (error) {
        next(error);
    }
};

// Sort a sheet's range
const sortSheetRange = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { sheetName, sortColumnIndex, order } = req.body;

        const sortOrder = order === 'DESCENDING' ? 'DESCENDING' : 'ASCENDING';

        const request = {
            spreadsheetId: sheetId,
            requests: [
                {
                    sortRange: {
                        range: {
                            sheetId: sheetName,
                            startRowIndex: 1,
                            startColumnIndex: 0,
                            endColumnIndex: sortColumnIndex + 1
                        },
                        sortSpecs: [{ dimensionIndex: sortColumnIndex, sortOrder }]
                    }
                }
            ]
        };

        await sheets.spreadsheets.batchUpdate(request);
        res.status(200).json({ success: true, message: 'Sheet sorted successfully' });
    } catch (error) {
        next(error);
    }
};

// Apply data validation and formatting 
const applyDataValidationAndFormatting = async (req, res, next) => {
    try {
        await setUserCredentials(req.user.id);
        const { sheetId } = req.params;
        const { sheetName, range, validationType, criteria } = req.body;

        const request = {
            spreadsheetId: sheetId,
            requests: [
                {
                    setDataValidation: {
                        range: {
                            sheetId: sheetName,
                            startRowIndex: range.startRowIndex,
                            endRowIndex: range.endRowIndex,
                            startColumnIndex: range.startColumnIndex,
                            endColumnIndex: range.endColumnIndex
                        },
                        rule: {
                            condition: {
                                type: validationType,
                                values: criteria
                            },
                            showCustomUi: true
                        }
                    }
                }
            ]
        };

        await sheets.spreadsheets.batchUpdate(request);
        res.status(200).json({ success: true, message: 'Data validation applied successfully' });
    } catch (error) {
        next(error);
    }
};




module.exports = {
    createSpreadSheet,
    renameSpreadSheet,
    deleteSpreadSheet,
    createSheet,
    renameSheet,
    getSheet,
    updateSheet,
    deleteSheet,
    appendData,
    clearDataFromSheet,
    deleteRowsFromSheet,
    deleteColumnFromSheet,
    listSheetsWithMetadata,
    listAllSpreadsheets,
    getSpreadsheetMetadata,
    sortSheetRange,
    applyDataValidationAndFormatting
};