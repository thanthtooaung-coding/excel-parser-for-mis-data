import * as XLSX from 'xlsx'
import './App.css'
import { useState } from 'react'
interface MisReportRow {
  no: string
  particular: string
  ref: string
  units: string
  previous_month: number
  during_month_increased: number
  during_month_decreased: number
}
interface SheetSqlOutput {
  sheetName: string
  sql: string
}

function App() {
  const [sqlOutputs, setSqlOutputs] = useState<SheetSqlOutput[]>([])
  const [error, setError] = useState<string>('')
  const [isLoading, setIsLoading] = useState<boolean>(false)

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    setIsLoading(true)
    setError('')
    setSqlOutputs([])

    const reader = new FileReader()

    reader.onload = (e) => {
      try {
        const data = e.target?.result
        const workbook = XLSX.read(data, { type: 'binary' })

        const generatedSqls: SheetSqlOutput[] = []

        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName]
          const jsonData: any[][] = XLSX.utils.sheet_to_json(worksheet, { header: 1 })

          let branchName = ''
          let reportMonth = ''
          jsonData.forEach(row => {
            const rowString = row.join(',').toLowerCase()
            if (rowString.includes('township and branch name')) {
              const cellValue = row.find(cell => typeof cell === 'string' && cell.toLowerCase().includes('township and branch name'))
              if (cellValue) {
                branchName = cellValue.split(':')[1]?.trim() || 'N/A'
              }
            }
            if (rowString.includes('report month')) {
              const dateIndex = row.findIndex(cell => typeof cell === 'string' && cell.toLowerCase().includes('report month'))
              if (dateIndex !== -1 && row[dateIndex + 1]) {
                const dateValue = row[dateIndex + 1]
                if (typeof dateValue === 'number') {
                    const date = XLSX.SSF.parse_date_code(dateValue)
                    reportMonth = `${date.y}-${String(date.m).padStart(2, '0')}-${String(date.d).padStart(2, '0')}`
                } else {
                    reportMonth = new Date(dateValue).toISOString().split('T')[0]
                }
              }
            }
          })
          
          if (!branchName || !reportMonth) return;

          const headerRowIndex = jsonData.findIndex(row =>
            String(row[0]).toLowerCase().trim() === 'no' && String(row[1]).toLowerCase().trim().includes('particular')
          )

          if (headerRowIndex === -1) return;

          const dataRows: MisReportRow[] = []
          
          let lastNo = '' 
          let lastParticular = ''

          for (let i = headerRowIndex + 1; i < jsonData.length; i++) {
            const row = jsonData[i]
            
            if (row[0]) {
              lastNo = String(row[0])
            }

            if (row[1]) {
              lastParticular = String(row[1])
            }
            
            const ref = row[5]

            if (ref && String(ref).toLowerCase().trim() !== 'total') {
              dataRows.push({
                no: lastNo,
                particular: lastParticular,
                ref: String(ref),
                units: String(row[6] || ''),
                previous_month: Number(row[7]) || 0,
                during_month_increased: Number(row[8]) || 0,
                during_month_decreased: Number(row[9]) || 0,
              })
            }
          }

          if (dataRows.length > 0) {
            const tableName = 'mis_reports'
            const columns = 'no, particular, ref, units, previous_month, during_month_increased, during_month_decreased, report_month, branch_name'
            
            const values = dataRows.map(row => {
              const escapedParticular = row.particular.replace(/'/g, "''")
              const escapedBranchName = branchName.replace(/'/g, "''")
              
              return `('${row.no}', '${escapedParticular}', '${row.ref}', '${row.units}', ${row.previous_month}, ${row.during_month_increased}, ${row.during_month_decreased}, '${reportMonth}', '${escapedBranchName}')`
            }).join(',\n');
            
            const sql = `INSERT INTO ${tableName} (${columns})\nVALUES\n${values};`
            generatedSqls.push({ sheetName, sql })
          }
        })
        
        setSqlOutputs(generatedSqls)

      } catch (err) {
        console.error(err)
        setError('Failed to process the file. Please ensure it is a valid Excel file.')
      } finally {
        setIsLoading(false)
      }
    }
    
    reader.onerror = () => {
        setError('Failed to read the file.');
        setIsLoading(false);
    };

    reader.readAsBinaryString(file)
  }
  
  const copyToClipboard = (text: string) => {
    navigator.clipboard.writeText(text);
  }

  return (
    <div className="container">
      <h1>📈 Excel to SQL Generator for MIS Reports</h1>
      <p>Upload your `xxx MIS Reports.xlsx` file. The tool will generate an SQL `INSERT` query for each sheet found.</p>
      
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
  )
}

export default App
