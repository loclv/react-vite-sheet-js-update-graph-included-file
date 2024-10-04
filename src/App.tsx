import ExcelJS from 'exceljs';
import { useState } from 'react';
import './App.css';
import reactLogo from './assets/react.svg';
import viteLogo from '/vite.svg';

const updateAGraphIncludedExcelFile = async (excelFilePath: string) => {
  try {
    const response = await fetch(excelFilePath);
    const blob = await response.blob();
    // Read the file as an ArrayBuffer
    const arrayBuffer = await blob.arrayBuffer();

    const workbook = new ExcelJS.Workbook();
    // Read the workbook from the ArrayBuffer
    await workbook.xlsx.load(arrayBuffer);

    const buffer = await workbook.xlsx.writeBuffer();

    const workbookBlob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const url = URL.createObjectURL(workbookBlob);

    const link = document.createElement('a');
    link.href = url;
    link.download = `updated-at-${new Date().getTime()}.xlsx`;
    link.click();

    URL.revokeObjectURL(url);
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
        <button onClick={() => setCount((count) => count + 1)}>
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
