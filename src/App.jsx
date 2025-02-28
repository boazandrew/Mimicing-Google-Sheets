import React from 'react';
import Spreadsheet from './components/Spreadsheet';

function App() {
  return (
    <div className="min-h-screen bg-gray-50">
      <header className="bg-white shadow-sm p-4 border-b">
        <h1 className="text-xl font-semibold text-gray-800">Google Sheets Clone</h1>
      </header>
      <main>
        <Spreadsheet />
      </main>
    </div>
  );
}

export default App;
