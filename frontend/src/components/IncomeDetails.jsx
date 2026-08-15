import React from 'react'

export default function IncomeDetails({data, onChange}) {
  const update = (k,v)=>onChange({...data, [k]: v})

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm p-6">
        <h2 className="text-xl font-bold text-gray-800 mb-4">Salary Income (Sec 17) – Income Details Sheet</h2>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label className="text-sm font-medium text-gray-700">Gross Salary (17(1))</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.salary||0} onChange={e=>update('salary', parseFloat(e.target.value)||0)}/>
          </div>
          <div>
            <label className="text-sm font-medium text-gray-700">Basic Salary</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.basic||0} onChange={e=>update('basic', parseFloat(e.target.value)||0)}/>
          </div>
          <div>
            <label className="text-sm font-medium text-gray-700">Dearness Allowance (DA)</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.da||0} onChange={e=>update('da', parseFloat(e.target.value)||0)}/>
          </div>
          <div>
            <label className="text-sm font-medium text-gray-700">HRA Received</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.hra_received||0} onChange={e=>update('hra_received', parseFloat(e.target.value)||0)}/>
          </div>
          <div>
            <label className="text-sm font-medium text-gray-700">Rent Paid (for HRA)</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.rent_paid||0} onChange={e=>update('rent_paid', parseFloat(e.target.value)||0)}/>
          </div>
          <div className="flex items-center gap-2 mt-6">
            <input type="checkbox" checked={data.is_metro||false} onChange={e=>update('is_metro', e.target.checked)} className="rounded"/>
            <label className="text-sm">Metro City? (50% vs 40% – EA 10(13A) G7 logic)</label>
          </div>
        </div>

        <div className="mt-4 p-3 bg-yellow-50 rounded text-xs text-yellow-800">
          <strong>HRA Exemption Formula (hidden sheet EA 10(13A) G12):</strong> Min( Actual HRA, Rent – 10% Salary, 50%/40% Salary ) = calculated in backend mdTaxCalc equivalent.
        </div>

        <div className="mt-6 grid grid-cols-1 md:grid-cols-2 gap-4">
          <div>
            <label className="text-sm font-medium text-gray-700">Other Sources – Interest</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.other_sources_interest||0} onChange={e=>update('other_sources_interest', parseFloat(e.target.value)||0)}/>
          </div>
          <div>
            <label className="text-sm font-medium text-gray-700">Other Sources – Dividend / Others</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={data.other_sources_dividend||0} onChange={e=>{
              const val=parseFloat(e.target.value)||0
              update('other_sources_dividend', val)
              update('other_sources', (data.other_sources_interest||0)+val)
            }}/>
          </div>
        </div>

        <div className="mt-3 text-sm text-gray-600">
          Total Other Sources: <strong>₹{( (data.other_sources_interest||0)+(data.other_sources_dividend||0) ).toLocaleString('en-IN')}</strong> → stored as other_sources for computation
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm p-6">
        <h3 className="font-semibold text-gray-800 mb-3">CSV Import (112A & 115AD – Capital Gains for ITR-2/3)</h3>
        <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center">
          <p className="text-sm text-gray-600">For ITR-2/3, upload CSV_112A.csv / CSV_115AD.csv format (ISIN, Buy Value, Sell Value) – parsing mirrors ImportExcel.bas</p>
          <input type="file" accept=".csv" className="mt-3 text-sm" onChange={(e)=>{
            const file=e.target.files[0]
            if(file){
              const reader=new FileReader()
              reader.onload=(ev)=>{
                try{
                  const text=ev.target.result
                  const lines=text.split('\n').slice(0,5)
                  alert(`CSV preview (first 5 lines):\n${lines.join('\n')}\n\nParsing stub implemented – would map to Capital Gains schedule`)
                }catch(err){alert(err)}
              }
              reader.readAsText(file)
            }
          }}/>
        </div>
      </div>
    </div>
  )
}
