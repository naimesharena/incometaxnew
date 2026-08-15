import React, { useState } from 'react'

export default function Deductions({data, onChange}) {
  const update = (k,v)=>onChange({...data, [k]: v})

  const [donations, setDonations] = useState(data.donations_80G||[])

  const addDonation = ()=>{
    const nd=[...donations, {donee_name:'', pan:'', amount:0, eligible_pct:'100%'}]
    setDonations(nd)
    update('donations_80G', nd)
    const tot=nd.reduce((s,d)=>s+ (d.amount||0),0)
    update('deduction_80G', tot)
  }

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm p-6">
        <h2 className="text-xl font-bold text-gray-800 mb-4">Deductions Chapter VI-A – 80C to 80U (Hidden Sheets)</h2>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label className="text-sm font-medium">80C (LIC, PPF, etc) – Cap 1.5L</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80C||0} onChange={e=>update('deduction_80C', parseFloat(e.target.value)||0)}/>
            <span className="text-xs text-gray-500">80C hidden B2:G25, Amount.80C E5:E8 + 80CCC F19:F22</span>
          </div>
          <div>
            <label className="text-sm font-medium">80CCD(1) – NPS Employee</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80CCD1||0} onChange={e=>update('deduction_80CCD1', parseFloat(e.target.value)||0)}/>
          </div>
          <div>
            <label className="text-sm font-medium">80CCD(1B) – Additional NPS 50k</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80CCD1B||0} onChange={e=>update('deduction_80CCD1B', parseFloat(e.target.value)||0)}/>
          </div>

          <div>
            <label className="text-sm font-medium">80D – Medical Insurance (Self/Family)</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80D||0} onChange={e=>update('deduction_80D', parseFloat(e.target.value)||0)}/>
            <span className="text-xs text-gray-500">Selection80D cascade: (Self, Family Non Senior etc) – 25000/50000/75000/1L – BK18 logic</span>
          </div>
          <div>
            <label className="text-sm font-medium">80D Preventive Health Checkup (cap 5k)</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80D_preventive||0} onChange={e=>update('deduction_80D_preventive', parseFloat(e.target.value)||0)}/>
          </div>
          <div>
            <label className="text-sm font-medium">80DD – Dependent Disability</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80DD||0} onChange={e=>update('deduction_80DD', parseFloat(e.target.value)||0)}/>
            <span className="text-xs text-gray-500">80U-80DD hidden H20 etc, Amtdeduction_80DD E20</span>
          </div>

          <div>
            <label className="text-sm font-medium">80E – Education Loan Interest</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80E||0} onChange={e=>update('deduction_80E', parseFloat(e.target.value)||0)}/>
          </div>
          <div>
            <label className="text-sm font-medium">80EE – Home Loan Interest (first house)</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80EE||0} onChange={e=>{
              const v=parseFloat(e.target.value)||0
              if(v>0 && (data.deduction_80EEA||0)>0){alert('Deduction u/s 80EE and 80EEA cannot be claimed together – lock_80EE_flag VBA'); return}
              update('deduction_80EE', v)
            }}/>
            <span className="text-xs text-red-500">Mutual exclusivity 80EE vs 80EEA – md80EE.bas</span>
          </div>
          <div>
            <label className="text-sm font-medium">80EEA – Affordable Housing</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80EEA||0} onChange={e=>{
              const v=parseFloat(e.target.value)||0
              if(v>0 && (data.deduction_80EE||0)>0){alert('80EE and 80EEA cannot together'); return}
              update('deduction_80EEA', v)
            }}/>
          </div>

          <div>
            <label className="text-sm font-medium">80EEB – Electric Vehicle Loan</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80EEB||0} onChange={e=>update('deduction_80EEB', parseFloat(e.target.value)||0)}/>
          </div>
          <div>
            <label className="text-sm font-medium">80U – Self Disability</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80U||0} onChange={e=>update('deduction_80U', parseFloat(e.target.value)||0)}/>
          </div>
          <div>
            <label className="text-sm font-medium">80GGC – Political Donations (need PAN/Bank)</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.deduction_80GGC||0} onChange={e=>update('deduction_80GGC', parseFloat(e.target.value)||0)}/>
          </div>
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm p-6">
        <div className="flex justify-between items-center mb-4">
          <h3 className="font-semibold">80G Donations – Categories A/B/C/D (80G hidden 173 rows AZ173)</h3>
          <button onClick={addDonation} className="px-3 py-1 bg-green-600 text-white rounded text-sm">+ Add Donation</button>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm border">
            <thead className="bg-gray-100">
              <tr>
                <th className="p-2 border">Donee Name</th>
                <th className="p-2 border">PAN of Donee</th>
                <th className="p-2 border">Amount</th>
                <th className="p-2 border">Eligible %</th>
                <th className="p-2 border">Action</th>
              </tr>
            </thead>
            <tbody>
              {donations.map((d,i)=>(
                <tr key={i}>
                  <td className="p-1 border"><input className="w-full border rounded px-2 py-1" value={d.donee_name} onChange={e=>{
                    const nd=[...donations]; nd[i].donee_name=e.target.value; setDonations(nd); update('donations_80G', nd)
                  }}/></td>
                  <td className="p-1 border"><input className="w-full border rounded px-2 py-1 uppercase" value={d.pan} maxLength={10} onChange={e=>{
                    const nd=[...donations]; nd[i].pan=e.target.value.toUpperCase(); setDonations(nd); update('donations_80G', nd)
                  }}/></td>
                  <td className="p-1 border"><input type="number" className="w-full border rounded px-2 py-1" value={d.amount} onChange={e=>{
                    const nd=[...donations]; nd[i].amount=parseFloat(e.target.value)||0; setDonations(nd); update('donations_80G', nd); update('deduction_80G', nd.reduce((s,x)=>s+(x.amount||0),0))
                  }}/></td>
                  <td className="p-1 border"><select className="w-full border rounded px-2 py-1" value={d.eligible_pct} onChange={e=>{
                    const nd=[...donations]; nd[i].eligible_pct=e.target.value; setDonations(nd); update('donations_80G', nd)
                  }}>
                    <option>100%</option><option>50%</option>
                  </select></td>
                  <td className="p-1 border"><button className="text-red-600" onClick={()=>{
                    const nd=donations.filter((_,idx)=>idx!==i); setDonations(nd); update('donations_80G', nd); update('deduction_80G', nd.reduce((s,x)=>s+(x.amount||0),0))
                  }}>Del</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="mt-2 text-xs text-gray-600">Total 80G: ₹{(data.deduction_80G||0).toLocaleString('en-IN')} – comb_80G_A Y8:Y11 named ranges, C_Eligible, CD_EligibleAmount AI3, etc. 10% of ATI cap calculated backend.</div>
      </div>
    </div>
  )
}
