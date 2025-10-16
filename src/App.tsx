import React, { useState } from 'react'

declare var XLSX: any; // Declares that XLSX is a global variable

// Define the structure for the parsed data row
interface MisReportRow {
  no: string;
  particular: string;
  ref: string;
  units: string;
  previous_month: number;
  increased: number;
  decreased: number;
  this_month: number;
}

// Define the structure for the output of each sheet
interface SheetSqlOutput {
  sheetName: string;
  sql: string;
}

// All CSS is now embedded within the component
const AppStyles = `
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    background-color: #f4f7f9;
    color: #333;
    padding: 20px;
    margin: 0;
  }

  .container {
    max-width: 900px;
    margin: 0 auto;
    background-color: #fff;
    padding: 30px;
    border-radius: 8px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
  }

  h1 {
    color: #2c3e50;
    border-bottom: 2px solid #e0e0e0;
    padding-bottom: 10px;
  }

  .input-group {
      display: flex;
      gap: 20px;
      margin: 20px 0;
  }

  .input-group div {
      display: flex;
      flex-direction: column;
  }

  .input-group label {
      margin-bottom: 5px;
      font-weight: bold;
      color: #555;
  }

  .input-group input {
      padding: 8px;
      border: 1px solid #ccc;
      border-radius: 4px;
  }


  input[type="file"] {
    display: block;
    margin: 20px 0;
    padding: 10px;
    border: 1px solid #ccc;
    border-radius: 4px;
  }

  .error {
    color: #e74c3c;
    background-color: #fbeae5;
    padding: 10px;
    border-radius: 4px;
  }

  .results {
    margin-top: 30px;
  }

  .sql-block {
    margin-bottom: 25px;
  }

  .sql-block h3 {
    color: #3498db;
    margin-bottom: 10px;
  }

  .sql-header {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 5px;
  }

  .sql-header button {
    padding: 5px 10px;
    border: 1px solid #ccc;
    background-color: #f0f0f0;
    border-radius: 4px;
    cursor: pointer;
    font-weight: bold;
  }
  .sql-header button:hover {
    background-color: #e0e0e0;
  }

  pre {
    background-color: #2d2d2d;
    color: #f8f8f2;
    padding: 15px;
    border-radius: 5px;
    white-space: pre-wrap;
    word-wrap: break-word;
    font-family: 'Courier New', Courier, monospace;
    font-size: 14px;
  }
`;

function App() {
  const [sqlOutputs, setSqlOutputs] = useState<SheetSqlOutput[]>([]);
  const [error, setError] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);

  // State for user-provided values
  const [officeId, setOfficeId] = useState<string>('1'); // Default to 1

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!officeId) {
        setError('Please provide an Office ID value before uploading.');
        return;
    }
    
    if (typeof XLSX === 'undefined') {
        setError('Error: The XLSX library is not loaded. Please ensure the CDN script tag is in your HTML file.');
        return;
    }

    setIsLoading(true);
    setError('');
    setSqlOutputs([]);

    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const generatedSqls: SheetSqlOutput[] = [];

        workbook.SheetNames.forEach((sheetName: string) => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

          let branchName = '';
          let reportMonth = '';
          jsonData.forEach(row => {
            const rowString = (row || []).join(',').toLowerCase();
            if (rowString.includes('township and branch name')) {
              const cellValue = (row || []).find(cell => typeof cell === 'string' && cell.toLowerCase().includes('township and branch name'));
              if (cellValue) {
                branchName = cellValue.split(':')[1]?.trim() || 'N/A';
              }
            }
            if (rowString.includes('report month')) {
              const dateIndex = (row || []).findIndex(cell => typeof cell === 'string' && cell.toLowerCase().includes('report month'));
              if (dateIndex !== -1 && row[dateIndex + 1]) {
                const dateValue = row[dateIndex + 1];
                if (typeof dateValue === 'number') {
                    const date = XLSX.SSF.parse_date_code(dateValue);
                    reportMonth = `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`;
                } else {
                    reportMonth = new Date(dateValue).toISOString().split('T')[0];
                }
              }
            }
          });
          
          if (!branchName || !reportMonth) return;

          const headerRowIndex = jsonData.findIndex(row =>
            row && String(row[0]).toLowerCase().trim() === 'no' && String(row[1]).toLowerCase().trim().includes('particular')
          );

          if (headerRowIndex === -1) return;

          const dataRows: MisReportRow[] = [];
          let lastNo = ''; 
          let lastParticular = '';

          for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i] || [];
            
            // A main category is defined by having a value in the 'No' column (col 0).
            // Only update the main 'lastParticular' when we find a new main category.
            if (row[0]) {
              lastNo = String(row[0]);
              if (row[1]) {
                lastParticular = String(row[1]);
              }
            }
            
            const ref = row[5];
            // A valid data row must have a 'Ref:' (column F) and not be a 'Total' summary row
            if (ref && String(ref).toLowerCase().trim() !== 'total') {
              
              let finalParticular = lastParticular;
              const subParticular = row[3]; // e.g., 'Compulsory', 'Voluntary' from column D

              // If a sub-category exists in column D, append it for more detail
              if (subParticular) {
                finalParticular = `${lastParticular} - ${subParticular}`;
              }

              dataRows.push({
                no: lastNo,
                particular: finalParticular,
                ref: String(ref),
                units: String(row[6] || ''), // Unit is in column G
                previous_month: Number(row[7]) || 0,
                increased: Number(row[8]) || 0,
                decreased: Number(row[9]) || 0,
                this_month: Number(row[10]) || 0, // Column K is index 10
              });
            }
          }

          if (dataRows.length > 0) {
            const tableName = 'ct_mis_report';
            const columns = 'no, particular, ref, units, report_month, previous_month, increased, decreased, this_month, office_id, created_by';
            
            const values = dataRows.map(row => {
              const escapedParticular = row.particular.replace(/'/g, "''");
              
              return `('${row.no}', '${escapedParticular}', '${row.ref}', '${row.units}', '${reportMonth}', ${row.previous_month}, ${row.increased}, ${row.decreased}, ${row.this_month}, ${officeId}, 'mifos')`;
            }).join(',\n');
            
            const sql = `INSERT INTO ${tableName} (${columns})\nVALUES\n${values};`;
            generatedSqls.push({ sheetName, sql });
          }
        });
        
        setSqlOutputs(generatedSqls);
      } catch (err) {
        console.error(err);
        setError('Failed to process the file. Please ensure it is a valid Excel file.');
      } finally {
        setIsLoading(false);
      }
    };
    
    reader.onerror = () => {
        setError('Failed to read the file.');
        setIsLoading(false);
    };

    reader.readAsBinaryString(file);
  };
  
  const copyToClipboard = (text: string) => {
    navigator.clipboard.writeText(text);
  };

  return (
    <>
      <style>{AppStyles}</style>
      <div className="container">
        <h1>📈 Excel to SQL Generator (MIS Reports)</h1>
        <p>Upload your MIS report. The tool will generate an SQL `INSERT` query for the `ct_mis_report` table for each sheet found.</p>
        
        <div className="input-group">
            <div>
                <label htmlFor="officeId">Office ID:</label>
                <input 
                    type="number" 
                    id="officeId"
                    value={officeId} 
                    onChange={(e) => setOfficeId(e.target.value)} 
                    placeholder="e.g., 1"
                />
            </div>
        </div>

        <input type="file" accept=".xlsx, .xls" onChange={handleFileUpload} />
        
        {isLoading && <p>Processing your file...</p>}
        {error && <p className="error">{error}</p>}
        
        <div className="results">
          {sqlOutputs.map((output, index) => (
            <div key={index} className="sql-block">
              <h3>Sheet: {output.sheetName}</h3>
              <div className="sql-header">
                <span>Generated SQL Query</span>
                <button onClick={() => copyToClipboard(output.sql)}>Copy</button>
              </div>
              <pre>
                <code>{output.sql}</code>
              </pre>
            </div>
          ))}
        </div>
      </div>
    </>
  );
}

export default App;

