import { useState } from 'react';
import { read, writeFile } from 'xlsx';
import './App.css';
import reactLogo from './assets/react.svg';
import viteLogo from '/vite.svg';

const updateAGraphIncludedExcelFile = async (excelFilePath: string) => {
  try {
    const response = await fetch(excelFilePath);
    const blob = await response.blob();
    // Read the file as an ArrayBuffer
    const arrayBuffer = await blob.arrayBuffer();

    // Read the workbook from the ArrayBuffer
    const workbook = read(arrayBuffer, { type: 'array' });

    // Step 2: Access the desired sheet by name or index
    // For example, first sheet
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Step 3: Modify the sheet (e.g., update cell A1)
    worksheet.B4 = { t: 'n', v: 25 };
    worksheet.C4 = { t: 'n', v: 25 };
    worksheet.D4 = { t: 'n', v: 25 };
    worksheet.E4 = { t: 'n', v: 25 };

    // Step 4: Write the updated workbook back to a file
    writeFile(workbook, 'updated_by_sheet_js.xlsx');
  } catch (error) {
    console.error('Error updating graph:', error);
  }
};

updateAGraphIncludedExcelFile('/Radar-Chart.xlsx');

function App() {
  const [count, setCount] = useState(0);

  return (
    <>
      <div>
        <a href="https://vitejs.dev" target="_blank" rel="noreferrer">
          <img src={viteLogo} className="logo" alt="Vite logo" />
        </a>
        <a href="https://react.dev" target="_blank" rel="noreferrer">
          <img src={reactLogo} className="logo react" alt="React logo" />
        </a>
      </div>
      <h1>Vite + React</h1>
      <div className="card">
        <button type="button" onClick={() => setCount((count) => count + 1)}>
          count is {count}
        </button>
        <p>
          Edit <code>src/App.tsx</code> and save to test HMR
        </p>
      </div>
      <p className="read-the-docs">
        Click on the Vite and React logos to learn more
      </p>
    </>
  );
}

export default App;
