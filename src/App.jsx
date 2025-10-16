import { useEffect, useState } from "react";
import * as XLSX from "xlsx";
import {
  Upload,
  Download,
  FileSpreadsheet,
  AlertCircle,
  Settings,
} from "lucide-react";
import divisionMappingData from "./assets/mapping.json";

export default function ExcelProcessor() {
  const [file, setFile] = useState(null);
  const [processing, setProcessing] = useState(false);
  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [showMapping, setShowMapping] = useState(false);
  const [divisionMapping, setDivisionMapping] = useState(null);
  const [mappingText, setMappingText] = useState("");

  useEffect(() => {
    if (divisionMapping) {
      localStorage.setItem("divisionMapping", JSON.stringify(divisionMapping));
    }
  }, [divisionMapping]);

  useEffect(() => {
    const loadMapping = async () => {
      try {
        // First check if there's a saved mapping in localStorage
        const savedMapping = localStorage.getItem("divisionMapping");
        if (savedMapping) {
          const parsed = JSON.parse(savedMapping);
          setDivisionMapping(parsed);
          setMappingText(JSON.stringify(parsed, null, 2));
          return;
        }

        // If no saved mapping, load from JSON file
        setDivisionMapping(divisionMappingData);
        setMappingText(JSON.stringify(divisionMappingData, null, 2));
      } catch (err) {
        setError("Failed to load division mapping file", err);
      }
    };
    loadMapping();
  }, []);

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0];
    if (selectedFile) {
      setFile(selectedFile);
      setError("");
      setSuccess("");
    }
  };

  const handleMappingUpdate = () => {
    try {
      const parsed = JSON.parse(mappingText);
      setDivisionMapping(parsed);
      setShowMapping(false);
      setError("");
      setSuccess("Division mapping updated successfully!");
    } catch (err) {
      setError("Invalid JSON format for division mapping", err);
    }
  };

  const processExcel = async () => {
    if (!file) {
      setError("Please select an Excel file first");
      return;
    }

    if (!divisionMapping) {
      setError("Division mapping not loaded yet");
      return;
    }

    setProcessing(true);
    setError("");
    setSuccess("");

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });

      const outputWorkbook = XLSX.utils.book_new();

      workbook.SheetNames.forEach((sheetName) => {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        const stationKey = Object.keys(jsonData[0] || {}).find(
          (key) => key.toUpperCase().trim() === "STATION"
        );

        if (!stationKey) {
          throw new Error(
            `Sheet "${sheetName}" does not have a STATION column`
          );
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

        // Build division data with all offices from mapping
        const sortedDivisions = Object.keys(divisionMapping).sort();
        const divisionData = {};

        sortedDivisions.forEach((division) => {
          const offices = divisionMapping[division];
          const officeData = [];

          // Add all offices from mapping, sorted alphabetically
          const sortedOffices = [...offices].sort((a, b) => a.localeCompare(b));
          sortedOffices.forEach((office) => {
            // Find matching station (case-insensitive)
            const matchingStation = Object.keys(stationCounts).find(
              (station) => station.toLowerCase() === office.toLowerCase()
            );
            const count = matchingStation ? stationCounts[matchingStation] : 0;
            officeData.push({ office, count });
          });

          divisionData[division] = officeData;
        });

        // Calculate max rows needed
        const maxRows = Math.max(
          ...Object.values(divisionData).map((d) => d.length)
        );

        // Create output data in horizontal format
        const outputData = [];

        // Header row 1: Division names
        const divisionHeaderRow = {};
        sortedDivisions.forEach((division, index) => {
          const colOffset = index * 2;
          divisionHeaderRow[`col${colOffset}`] = division;
          divisionHeaderRow[`col${colOffset + 1}`] = "";
        });
        outputData.push(divisionHeaderRow);

        // Header row 2: Office and No. of Toolkits
        const columnHeaderRow = {};
        sortedDivisions.forEach((division, index) => {
          const colOffset = index * 2;
          columnHeaderRow[`col${colOffset}`] = "Office";
          columnHeaderRow[`col${colOffset + 1}`] = "No. of Toolkits";
        });
        outputData.push(columnHeaderRow);

        // Data rows
        for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
          const dataRow = {};
          sortedDivisions.forEach((division, divIndex) => {
            const colOffset = divIndex * 2;
            const offices = divisionData[division];
            if (rowIndex < offices.length) {
              dataRow[`col${colOffset}`] = offices[rowIndex].office;
              dataRow[`col${colOffset + 1}`] = offices[rowIndex].count;
            } else {
              dataRow[`col${colOffset}`] = "";
              dataRow[`col${colOffset + 1}`] = "";
            }
          });
          outputData.push(dataRow);
        }

        // Total row
        const totalRow = {};
        sortedDivisions.forEach((division, divIndex) => {
          const colOffset = divIndex * 2;
          const total = divisionData[division].reduce(
            (sum, item) => sum + item.count,
            0
          );
          totalRow[`col${colOffset}`] = "Total";
          totalRow[`col${colOffset + 1}`] = total;
        });
        outputData.push(totalRow);

        // Create worksheet from array of arrays for better control
        const ws_data = outputData.map((row) => {
          const arr = [];
          let colIndex = 0;
          while (
            row[`col${colIndex}`] !== undefined ||
            row[`col${colIndex + 1}`] !== undefined
          ) {
            arr.push(row[`col${colIndex}`]);
            arr.push(row[`col${colIndex + 1}`]);
            colIndex += 2;
          }
          return arr;
        });

        const newWorksheet = XLSX.utils.aoa_to_sheet(ws_data);

        // Merge cells for division headers
        const merges = [];
        sortedDivisions.forEach((division, index) => {
          const colOffset = index * 2;
          merges.push({
            s: { r: 0, c: colOffset },
            e: { r: 0, c: colOffset + 1 },
          });
        });
        newWorksheet["!merges"] = merges;

        // Set column widths
        const colWidths = [];
        sortedDivisions.forEach(() => {
          colWidths.push({ wch: 20 }, { wch: 15 });
        });
        newWorksheet["!cols"] = colWidths;

        XLSX.utils.book_append_sheet(outputWorkbook, newWorksheet, sheetName);
      });

      // Generate and download the file
      const outputData = XLSX.write(outputWorkbook, {
        bookType: "xlsx",
        type: "array",
      });

      const blob = new Blob([outputData], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `processed_${file.name}`;
      link.click();
      URL.revokeObjectURL(url);

      setSuccess(
        `Successfully processed ${workbook.SheetNames.length} sheet(s)!`
      );
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
            <h2 className="text-3xl font-bold text-gray-800">
              Excel Station Processor
            </h2>
          </div>

          <p className="text-gray-600 text-center mb-8">
            Upload an Excel file with a STATION column to generate a
            division-based summary report
          </p>

          <div className="space-y-6">
            <button
              onClick={() => setShowMapping(!showMapping)}
              className="w-full bg-gray-100 text-gray-100 py-3 px-6 rounded-lg font-semibold hover:bg-gray-200 transition-colors flex items-center justify-center"
            >
              <Settings className="w-5 h-5 mr-2" />
              {showMapping ? "Hide" : "Configure"} Division Mapping
            </button>

            {showMapping && (
              <div className="border border-gray-200 rounded-lg p-4 bg-gray-50">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Division Mapping (JSON format)
                </label>
                <textarea
                  value={mappingText}
                  onChange={(e) => setMappingText(e.target.value)}
                  className="w-full h-48 p-3 border border-gray-300 rounded-lg font-mono text-sm"
                  placeholder='{"Division 1": ["Station1", "Station2"]}'
                  style={{color: 'black'}}
                />
                <button
                  onClick={handleMappingUpdate}
                  className="mt-3 bg-indigo-600 text-white py-2 px-4 rounded-lg font-semibold hover:bg-indigo-700 transition-colors"
                >
                  Update Mapping
                </button>
              </div>
            )}

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
                  {file ? file.name : "Click to upload Excel file"}
                </span>
                <span className="text-xs text-gray-500 mt-1">
                  Supports .xlsx and .xls files
                </span>
              </label>
            </div>

            <button
              onClick={processExcel}
              disabled={!file || !divisionMapping || processing}
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
                Division mapping loads automatically from division-mapping.json
              </li>
              <li className="flex items-start">
                <span className="font-bold text-indigo-600 mr-2">2.</span>
                (Optional) Configure mapping through the web interface
              </li>
              <li className="flex items-start">
                <span className="font-bold text-indigo-600 mr-2">3.</span>
                Upload Excel file with STATION column and process
              </li>
              <li className="flex items-start">
                <span className="font-bold text-indigo-600 mr-2">4.</span>
                Download Excel with divisions in columns and totals at the
                bottom
              </li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
}
