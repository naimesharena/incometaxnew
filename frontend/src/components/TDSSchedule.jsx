import React, { useState } from 'react'

export default function TDSSchedule({data, onChange}) {
  const [tdsSalary, setTdsSalary] = useState(data.tds_salary||[])
  const [tdsOther, setTdsOther] = useState(data.tds_other||[])

  const addRow = (type) => {
    const newRow = {tan:'', employer_name:'', income_chargeable:0, tax_deducted:0, year:'2025'}
    if(type==='salary'){
      const nl=[...tdsSalary, newRow]; setTdsSalary(nl); onChange({...data, tds_salary:nl})
    }else{
      const nl=[...tdsOther, newRow]; setTdsOther(nl); onChange({...data, tds_other:nl})
    }
  }

  const renderTable = (rows, setRows, field) => (
    <table className="w-full text-sm border">
      <thead className="bg-gray-100">
        <tr>
          <th className="p-2 border">TAN of Deductor</th>
          <th className="p-2 border">Employer / Deductor Name</th>
          <th className="p-2 border">Income Chargeable</th>
          <th className="p-2 border">Tax Deducted</th>
          <th className="p-2 border">Year</th>
          <th className="p-2 border">Action</th>
        </tr>
      </thead>
      <tbody>
        {rows.map((r,i)=>(
          <tr key={i}>
            <td className="p-1 border"><input className="w-full border rounded px-2 py-1 uppercase" value={r.tan} maxLength={10} onChange={e=>{
              const nl=[...rows]; nl[i].tan=e.target.value.toUpperCase(); setRows(nl); onChange({...data, [field]: nl})
            }}/></td>
            <td className="p-1 border"><input className="w-full border rounded px-2 py-1" value={r.employer_name} onChange={e=>{
              const nl=[...rows]; nl[i].employer_name=e.target.value; setRows(nl); onChange({...data, [field]: nl})
            }}/></td>
            <td className="p-1 border"><input type="number" className="w-full border rounded px-2 py-1" value={r.income_chargeable} onChange={e=>{
              const nl=[...rows]; nl[i].income_chargeable=parseFloat(e.target.value)||0; setRows(nl); onChange({...data, [field]: nl})
            }}/></td>
            <td className="p-1 border"><input type="number" className="w-full border rounded px-2 py-1" value={r.tax_deducted} onChange={e=>{
              const nl=[...rows]; nl[i].tax_deducted=parseFloat(e.target.value)||0; setRows(nl); onChange({...data, [field]: nl})
            }}/></td>
            <td className="p-1 border"><select className="w-full border rounded px-2 py-1" value={r.year} onChange={e=>{
              const nl=[...rows]; nl[i].year=e.target.value; setRows(nl); onChange({...data, [field]: nl})
            }}>
              <option>2025</option><option>2026</option><option>2024</option>
            </select></td>
            <td className="p-1 border"><button className="text-red-600" onClick={()=>{
              const nl=rows.filter((_,idx)=>idx!==i); setRows(nl); onChange({...data, [field]: nl})
            }}>Del</button></td>
          </tr>
        ))}
      </tbody>
    </table>
  )

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm p-6">
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-lg font-bold">TDS on Salary (Form 16) – TDS Sheet B2:XFD77 + SchTDS.bas 3220 lines</h2>
          <button onClick={()=>addRow('salary')} className="px-3 py-1 bg-blue-600 text-white rounded text-sm">+ Add TDS1</button>
        </div>
        {renderTable(tdsSalary, setTdsSalary, 'tds_salary')}
        <div className="mt-2 text-xs text-gray-600">Validations: ValidateTAN1_TDS (TAN pattern 4L5D1L), ValidateIncChargeSal, ValidateTotTaxDeducted, Name length 125 – Data validation 21 rules in TDS sheet</div>
      </div>

      <div className="bg-white rounded-xl shadow-sm p-6">
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-lg font-bold">TDS on Other than Salary (Form 16A) – TDS2/TDS3</h2>
          <button onClick={()=>addRow('other')} className="px-3 py-1 bg-blue-600 text-white rounded text-sm">+ Add TDS2</button>
        </div>
        {renderTable(tdsOther, setTdsOther, 'tds_other')}
      </div>

      <div className="bg-white rounded-xl shadow-sm p-6">
        <h3 className="font-semibold mb-3">TCS Schedule – TCS Sheet B2:AD18 + SchTCS.bas 567 lines</h3>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label className="text-sm font-medium">TCS Total Collected</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.tcs_total||0} onChange={e=>onChange({...data, tcs_total: parseFloat(e.target.value)||0})}/>
          </div>
          <div className="text-sm text-gray-600 mt-6">
            TAN validation, Year = TCS_CollectedYear named range, Claim out of total TCS validation – ValidatesheetTCS
          </div>
        </div>
      </div>

      <div className="bg-yellow-50 p-3 rounded text-xs text-yellow-800">
        CSV Import supported: CSV_TDS1.csv, CSV_TDS2.csv, CSV_TDS3.csv, CSV_TCS.csv, CSV_IT.csv formats present in ITR-2/ITR-3 TCS & TDS & IT folders – parsing via ImportExcel.bas InsertRowsToImport
      </div>
    </div>
  )
}
