import React, { useState } from 'react'

export default function BusinessIncome({data, onChange}) {
  const [nature, setNature] = useState(data.business_nature||[])
  const addNature = ()=>{
    const nn=[...nature, {code:'', name:'', is_manufacturing:false}]
    setNature(nn); onChange({...data, business_nature: nn})
  }
  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm p-6">
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-xl font-bold">Business Income – ITR-3 (Part A – P&L, BS, OI, Manufacturing, Trading) + ITR-4 (44AD,44ADA,44AE)</h2>
          <button onClick={addNature} className="px-3 py-1 bg-blue-600 text-white rounded text-sm">+ Add Nature of Business</button>
        </div>
        <div className="text-xs text-gray-600 p-2 bg-yellow-50 rounded mb-4">
          Maps to: ITR-2/3 sheets – Nature Of Business hidden, Part A-BS, Manufacturing Account, Trading Account, Profit and Loss, Part A-OI, Quantitative Details, BP, DPM-DOA, DEP_DCG, ESR, ICDS, 10AA, GST hidden. VBA: Part_A_General 1317 lines, PARTA_BS, BP, GST validations. Presumptive: ITR-4 44AE sheet hidden + BP visible.
        </div>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label className="text-sm font-medium">Business Type (ITR-4)</label>
            <select className="w-full border rounded-lg px-3 py-2 mt-1" value={data.presumptive_type||''} onChange={e=>onChange({...data, presumptive_type: e.target.value})}>
              <option value="">Select</option>
              <option value="44AD">44AD – Small Business (6%/8% presumptive)</option>
              <option value="44ADA">44ADA – Professionals (50%)</option>
              <option value="44AE">44AE – Goods Carriage</option>
            </select>
          </div>
          <div>
            <label className="text-sm font-medium">Gross Receipts / Turnover</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.gross_receipts||0} onChange={e=>onChange({...data, gross_receipts: parseFloat(e.target.value)||0})}/>
          </div>
          <div>
            <label className="text-sm font-medium">Presumptive Income (auto: 6% digital, 8% cash, 50% profession)</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1 bg-gray-50" value={data.presumptive_income||0} onChange={e=>onChange({...data, presumptive_income: parseFloat(e.target.value)||0})}/>
            <button className="text-xs text-blue-600 mt-1" onClick={()=>{
              const gross=data.gross_receipts||0
              const type=data.presumptive_type
              let income=0
              if(type==='44AD') income=gross*0.08
              else if(type==='44ADA') income=gross*0.5
              else if(type==='44AE') income=gross*0.75 // simplified per vehicle
              onChange({...data, presumptive_income: Math.round(income)})
            }}>Auto Calculate</button>
          </div>
        </div>

        <div className="mt-6">
          <h3 className="font-semibold mb-2">Nature of Business (NIC Code – from DataBase dropdown)</h3>
          <table className="w-full text-sm border">
            <thead className="bg-gray-100"><tr><th className="p-2 border">Code</th><th className="p-2 border">Business Name</th><th className="p-2 border">Manufacturing?</th><th className="p-2 border">Action</th></tr></thead>
            <tbody>
              {nature.map((n,i)=>(
                <tr key={i}>
                  <td className="p-1 border"><input className="w-full border rounded px-2 py-1" value={n.code} onChange={e=>{const nn=[...nature]; nn[i].code=e.target.value; setNature(nn); onChange({...data, business_nature: nn})}}/></td>
                  <td className="p-1 border"><input className="w-full border rounded px-2 py-1" value={n.name} onChange={e=>{const nn=[...nature]; nn[i].name=e.target.value; setNature(nn); onChange({...data, business_nature: nn})}}/></td>
                  <td className="p-1 border"><input type="checkbox" checked={n.is_manufacturing} onChange={e=>{const nn=[...nature]; nn[i].is_manufacturing=e.target.checked; setNature(nn); onChange({...data, business_nature: nn})}}/></td>
                  <td className="p-1 border"><button className="text-red-600" onClick={()=>{const nn=nature.filter((_,idx)=>idx!==i); setNature(nn); onChange({...data, business_nature: nn})}}>Del</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm p-6">
        <h3 className="font-semibold mb-3">Profit & Loss – Simplified (P&L hidden sheet + BP)</h3>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div><label className="text-sm font-medium">Gross Profit</label><input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.gross_profit||0} onChange={e=>onChange({...data, gross_profit: parseFloat(e.target.value)||0})}/></div>
          <div><label className="text-sm font-medium">Expenses (OI – Other Information)</label><input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.expenses||0} onChange={e=>onChange({...data, expenses: parseFloat(e.target.value)||0})}/></div>
          <div><label className="text-sm font-medium">Net Profit (BP)</label><input type="number" className="w-full border rounded-lg px-3 py-2 mt-1 bg-gray-50" value={(data.gross_profit||0)-(data.expenses||0)} readOnly/><span className="text-xs text-gray-500">BP sheet validation – SchBP.bas</span></div>
        </div>
      </div>
    </div>
  )
}
