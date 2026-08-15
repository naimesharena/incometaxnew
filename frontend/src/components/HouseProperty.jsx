import React, { useState } from 'react'

export default function HouseProperty({data, onChange}) {
  const [coOwners, setCoOwners] = useState(data.co_owners||[])
  const update = (k,v)=>onChange({...data, [k]: v})

  const addCoOwner = () => {
    const newList = [...coOwners, {name:'', pan:'', aadhaar:'', share:0}]
    setCoOwners(newList)
    update('co_owners', newList)
  }

  return (
    <div className="bg-white rounded-xl shadow-sm p-6 space-y-6">
      <h2 className="text-xl font-bold text-gray-800">House Property (HP Sheet + 24(b) Hidden)</h2>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div>
          <label className="text-sm font-medium">Property Type</label>
          <select className="w-full border rounded-lg px-3 py-2 mt-1" value={data.hp_type||'self_occupied'} onChange={e=>update('hp_type', e.target.value)}>
            <option value="self_occupied">Self Occupied (u/s 23(2))</option>
            <option value="let_out">Let Out</option>
            <option value="deemed_let_out">Deemed Let Out</option>
          </select>
        </div>
        <div>
          <label className="text-sm font-medium">Annual Letable Value / Rent Received</label>
          <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.annual_value||0} onChange={e=>update('annual_value', parseFloat(e.target.value)||0)}/>
        </div>
        <div>
          <label className="text-sm font-medium">House Property Income (computed, but override for MVP)</label>
          <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1 bg-gray-50" value={data.hp_income||0} onChange={e=>update('hp_income', parseFloat(e.target.value)||0)}/>
        </div>

        <div>
          <label className="text-sm font-medium">Municipal Taxes Paid</label>
          <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.municipal_tax||0} onChange={e=>update('municipal_tax', parseFloat(e.target.value)||0)}/>
        </div>
        <div>
          <label className="text-sm font-medium">Interest on Borrowed Capital (24b)</label>
          <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.interest_24b||0} onChange={e=>{
            const v=parseFloat(e.target.value)||0
            update('interest_24b', v)
            // Auto compute HP income: ALV - tax - 30% - interest + arrears
            // Simplified:Rent - municipal - 30% - interest
            const rent = data.annual_value||0
            const municipal = data.municipal_tax||0
            const balance = Math.max(0, rent - municipal)
            const thirty = balance*0.3
            const income = balance - thirty - v
            update('hp_income', income)
          }}/>
          <span className="text-xs text-gray-500">Cap 2L for self-occupied, Schedule 24(b) hidden W5 concatenation logic</span>
        </div>
        <div>
          <label className="text-sm font-medium">Arrears</label>
          <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.arrears||0} onChange={e=>update('arrears', parseFloat(e.target.value)||0)}/>
        </div>
      </div>

      <div className="border-t pt-4">
        <div className="flex justify-between items-center mb-3">
          <h3 className="font-semibold">Co-Owners (AddPropertyCoOWners – mdHouseProperty.bas 3205 lines)</h3>
          <button onClick={addCoOwner} className="px-3 py-1 bg-blue-600 text-white rounded text-sm">+ Add Co-Owner</button>
        </div>
        {coOwners.map((co,i)=>(
          <div key={i} className="grid grid-cols-1 md:grid-cols-4 gap-2 mb-2 p-2 bg-gray-50 rounded">
            <input placeholder="Name" className="border rounded px-2 py-1" value={co.name} onChange={e=>{
              const nl=[...coOwners]; nl[i].name=e.target.value; setCoOwners(nl); update('co_owners', nl)
            }}/>
            <input placeholder="PAN (ABCDE1234F)" className="border rounded px-2 py-1 uppercase" value={co.pan} maxLength={10} onChange={e=>{
              const nl=[...coOwners]; nl[i].pan=e.target.value.toUpperCase(); setCoOwners(nl); update('co_owners', nl)
            }}/>
            <input placeholder="Share %" type="number" className="border rounded px-2 py-1" value={co.share} onChange={e=>{
              const nl=[...coOwners]; nl[i].share=parseFloat(e.target.value)||0; setCoOwners(nl); update('co_owners', nl)
            }}/>
            <button className="text-red-600 text-sm" onClick={()=>{
              const nl=coOwners.filter((_,idx)=>idx!==i); setCoOwners(nl); update('co_owners', nl)
            }}>Remove</button>
          </div>
        ))}
      </div>

      <div className="p-3 bg-blue-50 rounded text-xs text-blue-700">
        <strong>Mapping:</strong> HP sheet B2:W82 + 24(b) hidden A1:W12 + HP.Co.Pan named ranges. VBA AddPropertyCoOWners creates dynamic rows G9:J15. Computation: BalanceALV=ALV – Unrealized – Tax, 30% of Balance, TotalDeduct=30%+Interest, Income=Balance – Deductions + Arrears.
      </div>
    </div>
  )
}
