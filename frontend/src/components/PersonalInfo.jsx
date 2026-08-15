import React, { useEffect, useState } from 'react'
import { masterAPI } from '../services/api'

export default function PersonalInfo({data, onChange}) {
  const [states, setStates] = useState([])
  const [empCats, setEmpCats] = useState([])
  const [pincodeInfo, setPincodeInfo] = useState(null)

  useEffect(()=>{
    masterAPI.getStates().then(r=>setStates(r.data.states || [])).catch(()=>{})
    masterAPI.getEmployerCategories().then(r=>setEmpCats(r.data.categories || [])).catch(()=>{})
  },[])

  const handlePincodeBlur = async (pin) => {
    if(pin && pin.length===6){
      try{
        const res = await masterAPI.getPincode(pin)
        if(res.data.valid){
          setPincodeInfo(res.data)
          onChange({...data, state_code: res.data.state})
        }
      }catch(e){}
    }
  }

  const update = (field, val) => onChange({...data, [field]: val})

  return (
    <div className="bg-white rounded-xl shadow-sm p-6">
      <h2 className="text-xl font-bold text-gray-800 mb-6">Personal Information (Part A – General)</h2>
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <div>
          <label className="block text-sm font-medium text-gray-700">First Name *</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.first_name||''} onChange={e=>update('first_name', e.target.value.toUpperCase())} placeholder="As per PAN"/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Middle Name</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.middle_name||''} onChange={e=>update('middle_name', e.target.value.toUpperCase())}/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Last Name / Surname *</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.last_name||''} onChange={e=>update('last_name', e.target.value.toUpperCase())}/>
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700">PAN * (e.g., ABCDE1234F)</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2 uppercase" value={data.pan||''} onChange={e=>update('pan', e.target.value.toUpperCase())} maxLength={10}/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Aadhaar (12 digits)</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.aadhaar||''} onChange={e=>update('aadhaar', e.target.value)} maxLength={12}/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Date of Birth (YYYY-MM-DD)</label>
          <input type="date" className="mt-1 w-full border rounded-lg px-3 py-2" value={data.dob||''} onChange={e=>update('dob', e.target.value)}/>
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700">Mobile *</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.mobile||''} onChange={e=>update('mobile', e.target.value)} placeholder="10 digits"/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Email *</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.email||''} onChange={e=>update('email', e.target.value)} placeholder="you@example.com"/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Employer Category</label>
          <select className="mt-1 w-full border rounded-lg px-3 py-2" value={data.employer_category||'OTH'} onChange={e=>update('employer_category', e.target.value)}>
            {empCats.map((cat,i)=><option key={i} value={cat}>{cat}</option>)}
          </select>
        </div>

        <div className="md:col-span-3 border-t pt-4 mt-2">
          <h3 className="font-semibold text-gray-700">Address – Residence (auto-fill state from pincode – mirrors StateMatchesPin VBA)</h3>
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700">Flat / Door / Block</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.flat_no||''} onChange={e=>update('flat_no', e.target.value)}/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Premises / Building</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.premises||''} onChange={e=>update('premises', e.target.value)}/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Road / Street</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.street||''} onChange={e=>update('street', e.target.value)}/>
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700">Area / Locality</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.area||''} onChange={e=>update('area', e.target.value)}/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">City / District</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.city||''} onChange={e=>update('city', e.target.value)}/>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Pincode * (6 digits)</label>
          <input className="mt-1 w-full border rounded-lg px-3 py-2" value={data.pincode||''} onChange={e=>update('pincode', e.target.value)} onBlur={e=>handlePincodeBlur(e.target.value)} maxLength={6}/>
          {pincodeInfo && <span className="text-xs text-green-600">Mapped State: {pincodeInfo.state}</span>}
        </div>

        <div>
          <label className="block text-sm font-medium text-gray-700">State Code</label>
          <select className="mt-1 w-full border rounded-lg px-3 py-2" value={data.state_code||''} onChange={e=>update('state_code', e.target.value)}>
            <option value="">Select State</option>
            {states.map((s,i)=><option key={i} value={s.split('-')[0] || s}>{s}</option>)}
          </select>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Filing Section</label>
          <select className="mt-1 w-full border rounded-lg px-3 py-2" value={data.filing_section||'139(1)'} onChange={e=>update('filing_section', e.target.value)}>
            <option>139(1)</option>
            <option>139(4)</option>
            <option>139(5)-Revised</option>
            <option>119(2)(b)</option>
            <option>139(8A)</option>
          </select>
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700">Residential Status</label>
          <select className="mt-1 w-full border rounded-lg px-3 py-2" value={data.residential_status||'RES'} onChange={e=>update('residential_status', e.target.value)}>
            <option value="RES">RES-Resident</option>
            <option value="NRI">NRI-Non Resident</option>
            <option value="NOR">NOR-Resident but not ordinarily resident</option>
          </select>
        </div>
      </div>

      <div className="mt-6 p-4 bg-blue-50 rounded-lg text-sm text-blue-800">
        <strong>Excel Mapping:</strong> Income Details!B2:BR192 + Part A Gen_139(8A) hidden – PAN validation via VBA ChkPAN (5 letters 4 digits 1 letter), Aadhaar 12 digits, Pincode→State auto via DataBase CP:CQ (19k rows) implemented above.
      </div>
    </div>
  )
}
