import React, { useEffect, useState, useRef } from "react";
import api from "../api/api";
import { useParams } from "react-router-dom";
import DatePicker from 'react-datepicker';
import 'react-datepicker/dist/react-datepicker.css';
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button"; // Assuming you have a Button component
import { Trash2 } from "lucide-react"; // Assuming you are using lucide-react for icons


// Import a simple X icon (you can replace this with a better icon)
const DeleteIcon = () => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" className="w-5 h-5 text-red-500 hover:text-red-700 cursor-pointer">
        <path strokeLinecap="round" strokeLinejoin="round" d="M6 18L18 6M6 6l12 12" />
    </svg>
);

const SheetDetail = () => {
    const { sheetId } = useParams();
    const [spreadsheetName, setSpreadsheetName] = useState("");
    const [sheets, setSheets] = useState([]);
    const [sheetData, setSheetData] = useState([]);
    const [loading, setLoading] = useState(true);
    const [sheetName, setSheetName] = useState("Sheet1");
    const [formValues, setFormValues] = useState({});
    const [columnCount, setColumnCount] = useState(0);
    const [isSubmitting, setIsSubmitting] = useState(false);
    const [activeTab, setActiveTab] = useState("view");
    const [datePickersVisible, setDatePickersVisible] = useState({});
    const [selectedDates, setSelectedDates] = useState({});
    const sheetContainerRef = useRef(null);
    const [isDeleting, setIsDeleting] = useState(false); // To prevent multiple delete requests
    const [deleteDialogOpen, setDeleteDialogOpen] = useState(false);
    const [rowToDeleteIndex, setRowToDeleteIndex] = useState(null);

    useEffect(() => {
        fetchSpreadsheetDetails();
        fetchSheets();
        fetchSheetData(); // Modified to only fetch sheet data
    }, [sheetName]);

    async function fetchSpreadsheetDetails() {
        try {
            const res = await api.get(`/api/sheets/${sheetId}/metadata`);
            setSpreadsheetName(res.data.metadata.properties.title || "Untitled Spreadsheet");
        } catch (error) {
            console.error("Error fetching spreadsheet details:", error.response ? error.response.data : error.message);
        }
    }

    async function fetchSheets() {
        try {
            const res = await api.get(`/api/sheets/${sheetId}/listSheets`);
            setSheets(res.data.sheets || []);
        } catch (error) {
            console.error("Error fetching sheets:", error.response ? error.response.data : error.message);
        }
    }

    async function fetchSheetData() {
        try {
            setLoading(true);
            const res = await api.post(`/api/sheets/${sheetId}`, { sheetName });
            const data = res.data.data || [];

            const maxColumns = data.reduce((max, row) => Math.max(max, row.length), 0);
            const normalizedData = data.map((row) => [...row, ...Array(maxColumns - row.length).fill("")]);

            setSheetData(normalizedData);
            setColumnCount(maxColumns);
        } catch (error) {
            console.error("Error fetching sheet data:", error.response ? error.response.data : error.message);
        } finally {
            setLoading(false);
        }
    }

    const handleInputChange = (key, value) => {
        setFormValues({ ...formValues, [key]: value });
    };

    const handleDateChange = (date, index) => {
        const updatedSelectedDates = { ...selectedDates };
        updatedSelectedDates[index] = date;
        setSelectedDates(updatedSelectedDates);
        setFormValues({ ...formValues, [`col${index}`]: date.toLocaleDateString() });
        setDatePickersVisible({ ...datePickersVisible, [index]: false });
    };

    const handleDateInputClick = (index) => {
        setDatePickersVisible({ ...datePickersVisible, [index]: !datePickersVisible[index] });
    };

    const handleAddRow = async () => {
        if (isSubmitting) return;

        setIsSubmitting(true);
        try {
            const maxColumns = sheetData.length > 0 ? sheetData[0].length : columnCount;
            const newRow = Array.from({ length: maxColumns }, (_, idx) => formValues[`col${idx}`] || "");

            await api.post(`/api/sheets/${sheetId}/append`, { sheetName, values: [newRow] });

            setFormValues({});
            setSelectedDates({});
            setDatePickersVisible({});
            await fetchSheetData(); // Re-fetch data after adding row
        } catch (error) {
            console.error("Error adding row:", error.response ? error.response.data : error.message);
        } finally {
            setIsSubmitting(false);
        }
    };

    const handleAddSheet = async () => {
        try {
            const newSheetTitle = `Sheet${sheets.length + 1}`;
            await api.post(`/api/sheets/${sheetId}/createSheet`, { sheetName: newSheetTitle });
            fetchSheets();
            setSheetName(newSheetTitle);
        } catch (error) {
            console.error("Error adding sheet:", error.response ? error.response.data : error.message);
        }
    };

    const isDateColumn = (index) => {
        if (sheetData && sheetData[0] && sheetData[0][index]) {
            const placeholder = sheetData[0][index];
            const dateRegex = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
            const lowerPlaceholder = placeholder.toLowerCase();
            return (
                dateRegex.test(placeholder) ||
                lowerPlaceholder === 'date' ||
                lowerPlaceholder === 'dates'
            );
        }
        return false;
    };

    const openDeleteDialog = (rowIndex) => {
        setRowToDeleteIndex(rowIndex);
        setDeleteDialogOpen(true);
    };

    const closeDeleteDialog = () => {
        setRowToDeleteIndex(null);
        setDeleteDialogOpen(false);
    };

    const confirmDeleteRow = async () => {
        if (isDeleting || rowToDeleteIndex === null) return;

        const rowNumberToDelete = rowToDeleteIndex + 2; // +1 for 0-based index, +1 to skip header row

        setIsDeleting(true);
        setDeleteDialogOpen(false); // Close the dialog
        try {
            // Assuming your backend expects 1-based row numbers for deletion
            await api.delete(`/api/sheets/${sheetId}/deleteRows`, { data: { sheetName, startRow: rowNumberToDelete, endRow: rowNumberToDelete } });
            await fetchSheetData(); // Re-fetch data after deletion
        } catch (error) {
            console.error("Error deleting row:", error.response ? error.response.data : error.message);
        } finally {
            setIsDeleting(false);
            setRowToDeleteIndex(null);
        }
    };

    return (
        <div className=" flex flex-col bg-gradient-to-br from-gray-100 to-gray-200">
            {/* Navbar */}
            <nav className="bg-gradient-to-r from-blue-600 to-blue-700 text-white p-4 left-0 w-full z-10 shadow-md">
                <h1 className="text-xl font-semibold">Spreadsheet: {spreadsheetName}</h1>
            </nav>

            {/* Sheet List */}
            <div className="bg-white p-3 w-full z-10 flex gap-2 overflow-x-auto shadow-sm">
                <button onClick={handleAddSheet} className="px-3 py-1 bg-green-500 text-white rounded-md hover:bg-green-600 transition-colors">
                    +
                </button>
                {sheets.map((sheet) => (
                    <button
                        key={sheet.sheetId}
                        className={`px-4 py-2 rounded-md ${sheetName === sheet.title ? "bg-blue-500 text-white" : "bg-gray-200 hover:bg-gray-300 transition-colors"}`}
                        onClick={() => {
                            setSheetName(sheet.title);
                            setColumnCount(0);
                            setSheetData([]);
                        }}
                    >
                        {sheet.title}
                    </button>
                ))}
            </div>

            {/* Tabs */}
            <div className="flex mt-4 py-2 justify-evenly bg-white shadow-sm">
                <button
                    className={`px-4 py-2 rounded-md ${activeTab === "view" ? "bg-blue-500 text-white" : "bg-gray-200 hover:bg-gray-300 transition-colors"}`}
                    onClick={() => setActiveTab("view")}
                >
                    View Sheet
                </button>
                <button
                    className={`px-4 py-2 rounded-md ${activeTab === "add" ? "bg-blue-500 text-white" : "bg-gray-200 hover:bg-gray-300 transition-colors"}`}
                    onClick={() => setActiveTab("add")}
                >
                    Add Row
                </button>
            </div>

            {/* Add Row Section */}
            {activeTab === "add" && sheetData.length > 0 && (
                <div className="sticky top-[260px] left-0 bg-white shadow p-4">
                    <h2 className="font-semibold mb-3">Add New Row</h2>
                    <div className="flex flex-wrap gap-3">
                        {sheetData[0].map((header, idx) => (
                            <div key={idx} className="relative">
                                <input
                                    type="text"
                                    placeholder={header || `Column ${idx + 1}`}
                                    className={`border p-2 rounded-md w-full sm:w-32 text-center ${isDateColumn(idx) ? 'cursor-pointer' : ''}`}
                                    value={formValues[`col${idx}`] || ""}
                                    onChange={(e) => handleInputChange(`col${idx}`, e.target.value)}
                                    onClick={() => isDateColumn(idx) && handleDateInputClick(idx)}
                                    readOnly={isDateColumn(idx)}
                                />
                                {isDateColumn(idx) && datePickersVisible[idx] && (
                                    <div className="absolute z-10 shadow-md rounded-lg bg-white mt-1">
                                        <DatePicker
                                            selected={selectedDates[idx]}
                                            onChange={(date) => handleDateChange(date, idx)}
                                            inline
                                        />
                                    </div>
                                )}
                            </div>
                        ))}
                    </div>
                    <button
                        onClick={handleAddRow}
                        disabled={isSubmitting}
                        className={`py-2 px-4 mt-3 rounded-md ${isSubmitting ? "bg-gray-400 cursor-not-allowed" : "bg-red-500 text-white hover:bg-red-600 transition-colors"}`}
                    >
                        {isSubmitting ? "Adding..." : "Add Row"}
                    </button>
                </div>
            )}

            {/* Scrollable Sheet Data Container */}
            {activeTab === "view" && (
                <div ref={sheetContainerRef} className="mt-4 bg-white p-4 max-h-[400px] overflow-y-auto">
                    {loading ? (
                        <div className="flex justify-center items-center h-full">
                            <div className="animate-spin rounded-full h-16 w-16 border-t-2 border-b-2 border-blue-500"></div>
                        </div>
                    ) : sheetData.length === 0 ? (
                        <div className="p-4 bg-white shadow rounded-lg">
                            <h2 className="font-semibold mb-3">Define Columns</h2>
                            <input
                                type="number"
                                min="1"
                                className="border p-2 mb-3 rounded-md w-full sm:w-40 text-center"
                                placeholder="Number of columns"
                                onChange={(e) => setColumnCount(parseInt(e.target.value, 10) || 0)}
                            />
                            {columnCount > 0 && (
                                <div>
                                    <h3 className="font-semibold mb-3">Enter First Row Data</h3>
                                    <div className="flex flex-wrap gap-3">
                                        {[...Array(columnCount)].map((_, idx) => (
                                            <input
                                                key={idx}
                                                type="text"
                                                placeholder={`Column ${idx + 1}`}
                                                className="border p-2 rounded-md w-full sm:w-32 text-center"
                                                onChange={(e) => handleInputChange(`col${idx}`, e.target.value)}
                                            />
                                        ))}
                                    </div>
                                    <button onClick={handleAddRow} className="bg-green-500 text-white py-2 px-4 mt-3 rounded-md hover:bg-green-600 transition-colors">
                                        Submit
                                    </button>
                                </div>
                            )}
                        </div>
                    ) : (
                        <div className="overflow-auto border border-gray-300 rounded-lg">
                            <table className="w-full border-collapse border">
                                <thead className="bg-gray-100">
                                    <tr>
                                        <th className="border p-3 min-w-[40px] text-center"></th> {/* Delete Icon Header */}
                                        {sheetData[0].map((header, colIdx) => (
                                            <th key={colIdx} className="border p-3 min-w-[100px] sm:min-w-[150px] text-center">
                                                {header}
                                            </th>
                                        ))}
                                    </tr>
                                </thead>
                                <tbody>
                                    {sheetData.slice(1).map((row, rowIdx) => (
                                        <tr key={rowIdx} className="even:bg-gray-50">
                                            <td className="border text-center p-3">
                                                <div onClick={() => openDeleteDialog(rowIdx)}>
                                                <Trash2 className="w-4 h-4 mx-auto text-red-500 hover:text-red-700" />
                                                </div>
                                            </td>
                                            {row.map((cell, colIdx) => (
                                                <td key={colIdx} className="border text-center p-3 min-w-[80px] sm:min-w-[150px] max-w-[200px] overflow-hidden">
                                                    <div className="w-24 sm:w-40 overflow-x-auto whitespace-nowrap">{cell}</div>
                                                </td>
                                            ))}
                                        </tr>
                                    ))}
                                </tbody>
                            </table>
                        </div>
                    )}
                </div>
            )}

            {/* Delete Confirmation Dialog */}
            <Dialog open={deleteDialogOpen} onOpenChange={setDeleteDialogOpen}>
                <DialogContent>
                    <DialogHeader>
                        <DialogTitle>Confirm Delete</DialogTitle>
                    </DialogHeader>
                    <p className="mb-4">Are you sure you want to delete this row?</p>
                    <DialogFooter>
                        <Button type="button" variant="secondary" onClick={closeDeleteDialog}>
                            Cancel
                        </Button>
                        <Button variant="destructive" onClick={confirmDeleteRow} disabled={isDeleting}>
                            {isDeleting ? "Deleting..." : "Delete"}
                        </Button>
                    </DialogFooter>
                </DialogContent>
            </Dialog>
        </div>
    );
};

export default SheetDetail;