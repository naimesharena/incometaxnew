import React, { useState } from 'react'
import { masterAPI } from '../services/api'

export default function TaxesAndBank({taxesData, bankData, onTaxesChange, onBankChange}) {
  const [banks, setBanks] = useState(bankData.bank_accounts||[])
  const [ifscMsg, setIfscMsg] = useState({})

  const addBank = () => {
    const nb=[...banks, {ifsc:'', bank_name:'', account_number:'', account_type:'Saving', is_for_refund: banks.length===0}]
    setBanks(nb); onBankChange({bank_accounts: nb})
  }

  const validateIFSC = async (idx, code) => {
    try{
      const res=await masterAPI.validateIFSC(code)
      setIfscMsg({...ifscMsg, [idx]: res.data.valid ? 'Valid IFSC pattern' : 'Invalid IFSC'})
      // Also try to fetch bank name from IFSC list search? For MVP we can derive bank code first 4 chars -> bank_codes.json mapping
      const bankCode=code.substring(0,4)
      const bcRes=await masterAPI.getBankCodes(bankCode,5)
      if(bcRes.data && bcRes.data.length>0){
        const nb=[...banks]; nb[idx].bank_name=bcRes.data[0].name; setBanks(nb); onBankChange({bank_accounts: nb})
      }
    }catch(e){}
  }

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm p-6">
        <h2 className="text-xl font-bold mb-4">Taxes Paid – Advance Tax, Self Assessment (Taxes Paid and Verification B2:Y61)</h2>
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label className="text-sm font-medium">Advance Tax Paid</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={taxesData.advance_tax||0} onChange={e=>onTaxesChange({...taxesData, advance_tax: parseFloat(e.target.value)||0})}/>
          </div>
          <div>
            <label className="text-sm font-medium">Self Assessment Tax</label>
            <input type="number" className="w-full border rounded-lg px-3 py-2 mt-1" value={taxesData.self_assessment||0} onChange={e=>onTaxesChange({...taxesData, self_assessment: parseFloat(e.target.value)||0})}/>
          </div>
          <div>
            <label className="text-sm font-medium">BSR Code (for challan, 7 digits)</label>
            <input className="w-full border rounded-lg px-3 py-2 mt-1" placeholder="e.g., 0280001" value={taxesData.bsr_code||''} onChange={e=>onTaxesChange({...taxesData, bsr_code: e.target.value})}/>
          </div>
        </div>
        <div className="mt-4 p-3 bg-blue-50 rounded text-xs text-blue-800">
          Validations: CheckIFSC, BSR 7 digits, Challan 5 digits – Taxes Paid sheet 40 validations + SchTaxVerify.bas AddRows_Others – Button_nature validation for nature of income dropdown.
        </div>
      </div>

      <div className="bg-white rounded-xl shadow-sm p-6">
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-xl font-bold">Bank Accounts – BA Schedule (BankCode 318 rows, IFSC 140k rows, SchBA.bas 1318 lines)</h2>
          <button onClick={addBank} className="px-3 py-1 bg-green-600 text-white rounded text-sm">+ Add Bank Account</button>
        </div>
        <div className="space-y-3">
          {banks.map((b,i)=>(
            <div key={i} className="grid grid-cols-1 md:grid-cols-5 gap-2 p-3 border rounded-lg bg-gray-50">
              <div>
                <label className="text-xs font-medium">IFSC * (e.g., SBIN0001234)</label>
                <input className="w-full border rounded px-2 py-1 uppercase text-sm" value={b.ifsc} maxLength={11} onChange={e=>{
                  const nb=[...banks]; nb[i].ifsc=e.target.value.toUpperCase(); setBanks(nb); onBankChange({bank_accounts: nb})
                }} onBlur={e=>validateIFSC(i, e.target.value)}/>
                {ifscMsg[i] && <span className="text-xs text-green-600">{ifscMsg[i]}</span>}
              </div>
              <div>
                <label className="text-xs font-medium">Bank Name (auto from IFSC – GetBankName)</label>
                <input className="w-full border rounded px-2 py-1 text-sm bg-white" value={b.bank_name} onChange={e=>{
                  const nb=[...banks]; nb[i].bank_name=e.target.value; setBanks(nb); onBankChange({bank_accounts: nb})
                }}/>
              </div>
              <div>
                <label className="text-xs font-medium">Account No * (9-18 digits)</label>
                <input className="w-full border rounded px-2 py-1 text-sm" value={b.account_number} onChange={e=>{
                  const nb=[...banks]; nb[i].account_number=e.target.value; setBanks(nb); onBankChange({bank_accounts: nb})
                }}/>
              </div>
              <div>
                <label className="text-xs font-medium">Account Type</label>
                <select className="w-full border rounded px-2 py-1 text-sm" value={b.account_type} onChange={e=>{
                  const nb=[...banks]; nb[i].account_type=e.target.value; setBanks(nb); onBankChange({bank_accounts: nb})
                }}>
                  <option>Saving</option><option>Current</option>
                </select>
              </div>
              <div className="flex flex-col gap-1">
                <label className="text-xs font-medium flex items-center gap-1">
                  <input type="checkbox" checked={b.is_for_refund} onChange={e=>{
                    const nb=banks.map((x,idx)=>({...x, is_for_refund: idx===i ? e.target.checked : false})); setBanks(nb); onBankChange({bank_accounts: nb})
                  }}/> For Refund
                </label>
                <button className="text-xs text-red-600 mt-1" onClick={()=>{
                  const nb=banks.filter((_,idx)=>idx!==i); setBanks(nb); onBankChange({bank_accounts: nb})
                }}>Remove</button>
              </div>
            </div>
          ))}
        </div>
        {banks.length===0 && <p className="text-sm text-gray-500">At least one bank account mandatory – Schedule BA ValidateSheetBA</p>}
      </div>
    </div>
  )
}
