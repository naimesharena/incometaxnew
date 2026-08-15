import React, { useState } from 'react'

export default function CapitalGains({data, onChange}) {
  const [entries, setEntries] = useState(data.capital_gains_entries||[])
  const addRow = () => {
    const ne=[...entries, {type:'112A', isin:'', buy_date:'', buy_value:0, sell_date:'', sell_value:0, gain:0}]
    setEntries(ne); onChange({...data, capital_gains_entries: ne})
  }
  return (
    <div className="bg-white rounded-xl shadow-sm p-6 space-y-4">
      <div className="flex justify-between items-center">
        <h2 className="text-xl font-bold">Capital Gains – CG Schedule + 112A + 115AD (ITR-2/3) – CSV_112A.csv, CSV_115AD.csv</h2>
        <button onClick={addRow} className="px-3 py-1 bg-blue-600 text-white rounded text-sm">+ Add CG Entry</button>
      </div>
      <div className="text-xs text-gray-600 p-2 bg-blue-50 rounded">
        Maps to hidden sheets: CG, Schedule 112A, Schedule 115AD(1)(iii) proviso, DPM-DOA, DEP_DCG – Excel has 5730 formulas for CG computation with indexation, cost inflation index, 112A grandfathering 31Jan2018 FMV, loss carry forward CYLA/BFLA/CFL.
      </div>
      <div className="overflow-x-auto">
        <table className="w-full text-xs border">
          <thead className="bg-gray-100">
            <tr>
              <th className="p-2 border">Type</th>
              <th className="p-2 border">ISIN / Asset</th>
              <th className="p-2 border">Buy Date</th>
              <th className="p-2 border">Buy Value</th>
              <th className="p-2 border">Sell Date</th>
              <th className="p-2 border">Sell Value</th>
              <th className="p-2 border">Gain (auto)</th>
              <th className="p-2 border">Action</th>
            </tr>
          </thead>
          <tbody>
            {entries.map((e,i)=>(
              <tr key={i}>
                <td className="p-1 border"><select className="border rounded px-1 py-1" value={e.type} onChange={ev=>{
                  const ne=[...entries]; ne[i].type=ev.target.value; setEntries(ne); onChange({...data, capital_gains_entries: ne})
                }}><option>112A</option><option>115AD</option><option>Short Term</option><option>Long Term</option></select></td>
                <td className="p-1 border"><input className="w-full border rounded px-1 py-1" value={e.isin} onChange={ev=>{const ne=[...entries]; ne[i].isin=ev.target.value; setEntries(ne); onChange({...data, capital_gains_entries: ne})}}/></td>
                <td className="p-1 border"><input type="date" className="border rounded px-1 py-1" value={e.buy_date} onChange={ev=>{const ne=[...entries]; ne[i].buy_date=ev.target.value; setEntries(ne); onChange({...data, capital_gains_entries: ne})}}/></td>
                <td className="p-1 border"><input type="number" className="w-full border rounded px-1 py-1" value={e.buy_value} onChange={ev=>{const ne=[...entries]; ne[i].buy_value=parseFloat(ev.target.value)||0; ne[i].gain=(ne[i].sell_value||0)-(ne[i].buy_value||0); setEntries(ne); onChange({...data, capital_gains_entries: ne})}}/></td>
                <td className="p-1 border"><input type="date" className="border rounded px-1 py-1" value={e.sell_date} onChange={ev=>{const ne=[...entries]; ne[i].sell_date=ev.target.value; setEntries(ne); onChange({...data, capital_gains_entries: ne})}}/></td>
                <td className="p-1 border"><input type="number" className="w-full border rounded px-1 py-1" value={e.sell_value} onChange={ev=>{const ne=[...entries]; ne[i].sell_value=parseFloat(ev.target.value)||0; ne[i].gain=(ne[i].sell_value||0)-(ne[i].buy_value||0); setEntries(ne); onChange({...data, capital_gains_entries: ne})}}/></td>
                <td className="p-1 border font-mono">₹{(e.gain||0).toLocaleString()}</td>
                <td className="p-1 border"><button className="text-red-600" onClick={()=>{const ne=entries.filter((_,idx)=>idx!==i); setEntries(ne); onChange({...data, capital_gains_entries: ne})}}>Del</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="text-sm">Total CG Gain: <strong>₹{entries.reduce((s,e)=>s+(e.gain||0),0).toLocaleString()}</strong></div>
    </div>
  )
}
