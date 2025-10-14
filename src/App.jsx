import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Download, FileSpreadsheet, AlertCircle } from 'lucide-react';

export default function ExcelProcessor() {
  const [file, setFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState('');

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      setFile(selectedFile);
      setError('');
      setSuccess('');
    }
  };

  const processExcel = async () => {
    if (!file) {
      setError('Please select an Excel file first');
      return;
    }

    setProcessing(true);
    setError('');
    setSuccess('');

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });

      // Create new workbook for output
      const outputWorkbook = XLSX.utils.book_new();

      // Process each sheet
      workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        // Find STATION column (case-insensitive)
        const stationKey = Object.keys(jsonData[0] || {}).find(
          key => key.toUpperCase() === 'STATION'
        );

        if (!stationKey) {
          throw new Error(`Sheet "${sheetName}" does not have a STATION column`);
        }

        // Count occurrences of each station
        const stationCounts = {};
        jsonData.forEach((row) => {
          const station = row[stationKey];
          if (station) {
            const stationStr = String(station).trim();
            if (stationStr) {
              stationCounts[stationStr] = (stationCounts[stationStr] || 0) + 1;
            }
          }
        });

        // Convert to array and sort alphabetically
        const sortedStations = Object.entries(stationCounts)
          .map(([office, count]) => ({
            'OFFICE': office,
            'NO. OF TOOLKITS': count
          }))
          .sort((a, b) => a.OFFICE.localeCompare(b.OFFICE));

        // Create new worksheet
        const newWorksheet = XLSX.utils.json_to_sheet(sortedStations);
        XLSX.utils.book_append_sheet(outputWorkbook, newWorksheet, sheetName);
      });

      // Generate and download the file
      const outputData = XLSX.write(outputWorkbook, {
        bookType: 'xlsx',
        type: 'array'
      });

      const blob = new Blob([outputData], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });

      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `processed_${file.name}`;
      link.click();
      URL.revokeObjectURL(url);

      setSuccess(`Successfully processed ${workbook.SheetNames.length} sheet(s)!`);
    } catch (err) {
      setError(`Error processing file: ${err.message}`);
    } finally {
      setProcessing(false);
    }
  };

  return (
    <div className="min-h-screen min-w-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-8">
      <div className="max-w-2xl mx-auto">
        <div className="bg-white rounded-2xl shadow-xl p-8">
          <div className="flex items-center justify-center mb-6">
            <FileSpreadsheet className="w-12 h-12 text-indigo-600 mr-3" />
            <h1 className="text-3xl font-bold text-gray-800">
              Excel Station Processor
            </h1>
          </div>

          <p className="text-gray-600 text-center mb-8">
            Upload an Excel file with a STATION column to generate a summary report
          </p>

          <div className="space-y-6">
            <div className="border-2 border-dashed border-gray-300 rounded-lg p-8 text-center hover:border-indigo-400 transition-colors">
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={handleFileChange}
                className="hidden"
                id="file-upload"
              />
              <label
                htmlFor="file-upload"
                className="cursor-pointer flex flex-col items-center"
              >
                <Upload className="w-12 h-12 text-gray-400 mb-3" />
                <span className="text-sm font-medium text-gray-700">
                  {file ? file.name : 'Click to upload Excel file'}
                </span>
                <span className="text-xs text-gray-500 mt-1">
                  Supports .xlsx and .xls files
                </span>
              </label>
            </div>

            <button
              onClick={processExcel}
              disabled={!file || processing}
              className="w-full bg-indigo-600 text-white py-3 px-6 rounded-lg font-semibold hover:bg-indigo-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors flex items-center justify-center"
            >
              {processing ? (
                <>
                  <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white mr-2"></div>
                  Processing...
                </>
              ) : (
                <>
                  <Download className="w-5 h-5 mr-2" />
                  Process & Download
                </>
              )}
            </button>

            {error && (
              <div className="bg-red-50 border border-red-200 rounded-lg p-4 flex items-start">
                <AlertCircle className="w-5 h-5 text-red-600 mr-3 flex-shrink-0 mt-0.5" />
                <p className="text-red-800 text-sm">{error}</p>
              </div>
            )}

            {success && (
              <div className="bg-green-50 border border-green-200 rounded-lg p-4">
                <p className="text-green-800 text-sm font-medium">{success}</p>
              </div>
            )}
          </div>

          <div className="mt-8 pt-6 border-t border-gray-200">
            <h2 className="text-lg font-semibold text-gray-800 mb-3">
              How it works:
            </h2>
            <ul className="space-y-2 text-sm text-gray-600">
              <li className="flex items-start">
                <span className="font-bold text-indigo-600 mr-2">1.</span>
                Upload an Excel file with one or more sheets
              </li>
              <li className="flex items-start">
                <span className="font-bold text-indigo-600 mr-2">2.</span>
                Each sheet must have a STATION column
              </li>
              <li className="flex items-start">
                <span className="font-bold text-indigo-600 mr-2">3.</span>
                The app counts unique stations per sheet
              </li>
              <li className="flex items-start">
                <span className="font-bold text-indigo-600 mr-2">4.</span>
                Download a new Excel file with OFFICE and NO. OF TOOLKITS columns
              </li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
}