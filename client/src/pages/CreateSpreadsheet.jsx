import React, { useState } from 'react';
import api from '../api/api';
import { useNavigate } from 'react-router-dom';
import DatePicker from 'react-datepicker';
import 'react-datepicker/dist/react-datepicker.css';

const CreateSpreadsheet = () => {
  const navigate = useNavigate();
  const [spreadsheetName, setSpreadsheetName] = useState('');
  const [columns, setColumns] = useState([{ name: '', type: 'Text' }]);
  const [selectedDates, setSelectedDates] = useState({});
  const [showCalendar, setShowCalendar] = useState(null);
  const [isSubmitting, setIsSubmitting] = useState(false); // New state

  const handleColumnChange = (index, key, value) => {
    const updatedCols = [...columns];
    updatedCols[index][key] = value;
    setColumns(updatedCols);
    if (key === 'type' && value !== 'Date') {
      setShowCalendar(null);
    }
  };

  const addColumn = () => {
    setColumns([...columns, { name: '', type: 'Text' }]);
  };

  const removeColumn = (index) => {
    setColumns(columns.filter((_, i) => i !== index));
    const updatedSelectedDates = { ...selectedDates };
    delete updatedSelectedDates[index];
    setSelectedDates(updatedSelectedDates);
  };

  const handleDateChange = (date, index) => {
    const updatedSelectedDates = { ...selectedDates };
    updatedSelectedDates[index] = date;
    setSelectedDates(updatedSelectedDates);
    const updatedCols = [...columns];
    updatedCols[index].name = date.toLocaleDateString();
    setColumns(updatedCols);
    setShowCalendar(null);
  };

  const handleColumnNameClick = (index) => {
    if (columns[index].type === 'Date') {
      setShowCalendar(index);
    } else {
      setShowCalendar(null);
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (isSubmitting) {
      return; // Prevent multiple submissions
    }
    setIsSubmitting(true);
    try {
      const res = await api.post('/api/sheets/createSpreadSheet', {
        title: spreadsheetName,
      });
      const { spreadsheetId } = res.data;

      await api.post(`/api/sheets/${spreadsheetId}/writeBoldText`, {
        sheetName: 'Sheet1',
        range: 'A1:' + String.fromCharCode(64 + columns.length) + '1',
        values: [columns.map((col) => col.name)],
      });

      navigate(`/sheets/${spreadsheetId}`);
    } catch (error) {
      console.error(error);
    } finally {
      setIsSubmitting(false); // Re-enable the button
    }
  };

  return (
    <div className="min-h-screen p-4 bg-gradient-to-br from-gray-100 to-gray-200">
      <div className="max-w-md mx-auto bg-white p-6 rounded-lg shadow-md">
        <h1 className="text-2xl font-semibold mb-6 text-center text-blue-600">New Spreadsheet</h1>
        <form onSubmit={handleSubmit} className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700">Spreadsheet Name</label>
            <input
              type="text"
              className="mt-1 block w-full border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 p-2"
              value={spreadsheetName}
              onChange={(e) => setSpreadsheetName(e.target.value)}
              required
            />
          </div>

          {columns.map((col, idx) => (
            <div key={idx} className="flex items-center space-x-2">
              <input
                type="text"
                className={`flex-1 border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 p-2 ${
                  col.type === 'Date' ? 'cursor-pointer' : ''
                }`}
                placeholder={col.type === 'Date' ? 'Select Date' : 'Column Name'}
                value={col.name}
                onChange={(e) => handleColumnChange(idx, 'name', e.target.value)}
                onClick={() => handleColumnNameClick(idx)}
                required
              />
              <select
                className="border-gray-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 p-2"
                value={col.type}
                onChange={(e) => handleColumnChange(idx, 'type', e.target.value)}
              >
                <option>Text</option>
                <option>Number</option>
                <option>Date</option>
              </select>
              {columns.length > 1 && (
                <button type="button" className="text-red-500 hover:text-red-700" onClick={() => removeColumn(idx)}>
                  X
                </button>
              )}
              {showCalendar === idx && (
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

          <button
            type="button"
            onClick={addColumn}
            className="bg-gray-700 hover:bg-gray-800 text-white py-2 px-4 rounded-md w-fit"
          >
            Add Column +
          </button>

          <button
            type="submit"
            className="bg-red-500 hover:bg-red-600 text-white py-3 rounded-md w-full"
            disabled={isSubmitting} // Disable the button
          >
            {isSubmitting ? 'Creating...' : 'Create Spreadsheet'} {/* Show loading text */}
          </button>
        </form>
      </div>
    </div>
  );
};

export default CreateSpreadsheet;