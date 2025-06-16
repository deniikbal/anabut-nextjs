'use client';

import { useState } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';

export default function Home() {
  const [inputFile, setInputFile] = useState<File | null>(null);
  const [processing, setProcessing] = useState(false);

  const createDynamicMapping = (values: any[], keyValue?: string) => {
    const allLetters = ['A', 'B', 'C', 'D', 'E'];
    const uniqueValues = [...new Set(values.filter(v => v !== null && v !== undefined))];
    
    if (keyValue) {
      const keyLetter = allLetters[Math.floor(Math.random() * allLetters.length)];
      const remainingLetters = allLetters.filter(l => l !== keyLetter);
      const mapping: { [key: string]: string } = { [keyValue]: keyLetter };
      
      uniqueValues.forEach((val, i) => {
        if (String(val) !== String(keyValue)) {
          mapping[String(val)] = remainingLetters[i % 4];
        }
      });
      
      return mapping;
    }
    
    return uniqueValues.reduce((acc, val, i) => {
      acc[String(val)] = allLetters[i % 5];
      return acc;
    }, {} as { [key: string]: string });
  };

  const getColumnWidth = (data: any[][], columnIndex: number): number => {
    // Get all values in the column
    const values = data.map(row => String(row[columnIndex] || ''));
    // Find the longest value
    const maxLength = Math.max(...values.map(value => value.length));
    // Add some padding (2 characters)
    return maxLength + 2;
  };

  const processExcel = async (file: File) => {
    setProcessing(true);
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

      // First pass: Create mappings for each column
      const columnMappings: { [key: number]: { [key: string]: string } } = {};
      const firstRow = jsonData[0] || [];
      
      for (let colIndex = 2; colIndex < firstRow.length; colIndex++) {
        const keyValue = String(firstRow[colIndex]);
        const colValues = jsonData.slice(1).map((r: any[]) => r[colIndex]);
        columnMappings[colIndex] = createDynamicMapping(colValues, keyValue);
      }

      // Process header row first
      const headerRow = [...firstRow];
      for (let colIndex = 2; colIndex < firstRow.length; colIndex++) {
        const keyValue = String(firstRow[colIndex]);
        headerRow[colIndex] = columnMappings[colIndex][keyValue];
      }

      // Second pass: Process the data rows
      const processedData = jsonData.map((row: any[], rowIndex: number) => {
        if (rowIndex === 0) {
          return headerRow;
        }

        const processedRow = [...row];

        // Process data rows
        for (let colIndex = 2; colIndex < row.length; colIndex++) {
          const value = String(row[colIndex]);
          processedRow[colIndex] = columnMappings[colIndex][value] || '';
        }

        // Add combined column
        const combined = processedRow.slice(2)
          .filter((val: string) => ['A', 'B', 'C', 'D', 'E'].includes(val))
          .join('');
        processedRow.push(combined);

        // Add correct answers count
        const correct = combined.split('').filter((val: string, i: number) => {
          return val === headerRow[i + 2];
        }).length;
        processedRow.push(`${correct}/${combined.length}`);

        return processedRow;
      });

      // Create new workbook
      const newWorkbook = XLSX.utils.book_new();
      const newWorksheet = XLSX.utils.aoa_to_sheet(processedData);

      // Set column widths
      const maxColumns = Math.max(...processedData.map(row => row.length));
      const columnWidths: { [key: string]: number } = {};
      
      for (let i = 0; i < maxColumns; i++) {
        const colLetter = XLSX.utils.encode_col(i);
        columnWidths[colLetter] = getColumnWidth(processedData, i);
      }

      // Apply column widths
      newWorksheet['!cols'] = Object.keys(columnWidths).map(col => ({
        wch: columnWidths[col]
      }));

      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Processed Data');

      // Generate filename with timestamp
      const timestamp = new Date().toISOString().replace(/[:.]/g, '').slice(0, 15);
      const filename = `ANABUT_${timestamp}.xlsx`;

      // Save file
      const excelBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      saveAs(blob, filename);

    } catch (error) {
      console.error('Error processing file:', error);
      alert('Error processing file. Please try again.');
    } finally {
      setProcessing(false);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      setInputFile(file);
    }
  };

  const handleProcess = () => {
    if (inputFile) {
      processExcel(inputFile);
    } else {
      alert('Please select a file first');
    }
  };

  return (
    <main className="min-h-screen bg-gray-50 py-12 px-4 sm:px-6 lg:px-8">
      <div className="max-w-md mx-auto bg-white rounded-lg shadow-lg p-6">
        <h1 className="text-2xl font-bold text-gray-900 mb-6 text-center">
          Excel Mapper
        </h1>
        
        <div className="space-y-4">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">
              Input File
            </label>
            <div className="flex items-center space-x-4">
              <input
                type="file"
                accept=".xlsx"
                onChange={handleFileChange}
                className="block w-full text-sm text-gray-500
                  file:mr-4 file:py-2 file:px-4
                  file:rounded-md file:border-0
                  file:text-sm file:font-semibold
                  file:bg-blue-50 file:text-blue-700
                  hover:file:bg-blue-100"
              />
            </div>
          </div>

          <button
            onClick={handleProcess}
            disabled={!inputFile || processing}
            className={`w-full py-2 px-4 rounded-md text-white font-medium
              ${!inputFile || processing
                ? 'bg-gray-400 cursor-not-allowed'
                : 'bg-blue-600 hover:bg-blue-700'
              }`}
          >
            {processing ? 'Processing...' : 'Process'}
          </button>
        </div>

        <div className="mt-6 text-center text-sm text-gray-500">
          © 2024 ANABUT - Excel Mapper
        </div>
      </div>
    </main>
  );
}
