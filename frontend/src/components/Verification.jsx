import React, { useState } from 'react'
import { itrAPI } from '../services/api'

export default function Verification({fullPayload, computation}) {
  const [jsonResult, setJsonResult] = useState(null)
  const [loading, setLoading] = useState(false)

  const handleGenerate = async () => {
    setLoading(true)
    try{
      const res=await itrAPI.generateJSON(fullPayload)
      setJsonResult(res.data)
    }catch(e){
      alert('Error generating JSON: '+(e.response?.data?.detail || e.message))
    }finally{setLoading(false)}
  }

  const downloadJSON = () => {
    if(!jsonResult) return
    const blob=new Blob([jsonResult.json_string], {type:'application/json'})
    const url=URL.createObjectURL(blob)
    const a=document.createElement('a')
    a.href=url
    a.download=`${fullPayload.itr_form}_AY26-27_${fullPayload.personal_info?.pan||'PAN'}.json`
    a.click()
  }

  const copyDigest = () => {
    if(jsonResult?.digest){
      navigator.clipboard.writeText(jsonResult.digest)
      alert('Digest copied: '+jsonResult.digest)
    }
  }

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm p-6">
        <h2 className="text-xl font-bold mb-4">Verification & JSON Generation – GenerateJson.bas 7120 lines + mdHashing + HS256.cls</h2>
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div>
            <h3 className="font-semibold mb-2">Declaration</h3>
            <div className="space-y-3 text-sm">
              <div>
                <label className="block text-xs font-medium">Place</label>
                <input className="w-full border rounded px-3 py-2" defaultValue={fullPayload.personal_info?.city||'Delhi'} />
              </div>
              <div>
                <label className="block text-xs font-medium">Capacity</label>
                <select className="w-full border rounded px-3 py-2">
                  <option>Self</option><option>Representative</option>
                </select>
              </div>
              <div className="text-xs text-gray-600 mt-4 p-2 bg-gray-50 rounded">
                I solemnly declare that to the best of my knowledge the information given in this return is correct and complete and is in accordance with provisions of Income-tax Act, 1961.
                <br/>Maps to Verification sheet Taxes Paid and Verification!B2:Y61 + Part A Gen_139(8A) if applicable.
              </div>
            </div>
          </div>
          <div>
            <h3 className="font-semibold mb-2">Hash Meta (from DataBase sheet B3:B4)</h3>
            <div className="text-sm bg-yellow-50 p-3 rounded">
              <div>Hash Key: <code>7Z3mxclnABiXtYG</code> (DataBase!B3)</div>
              <div>Hash Iteration: <code>1849</code> (DataBase!B4)</div>
              <div>Date Processing: 2026-08-02</div>
              <div className="mt-2 text-xs">Original VBA: Base64_HMACSHA256_JSON + EncodeBase64json + ToJsonFormat – Digest = Base64( iterated HMAC-SHA256(JSON, key) ) – HS256.cls HMAC impl – matches dept e-filing acceptance</div>
            </div>
            <button onClick={handleGenerate} disabled={loading} className="w-full mt-4 px-6 py-3 bg-indigo-600 text-white rounded-lg font-semibold hover:bg-indigo-700 disabled:opacity-50">
              {loading ? 'Generating...' : 'Generate Official JSON + Digest'}
            </button>
          </div>
        </div>
      </div>

      {jsonResult && (
        <div className="bg-white rounded-xl shadow-sm p-6">
          <div className="flex justify-between items-center mb-4">
            <h3 className="text-lg font-bold">Generated JSON – {jsonResult.itr_form} AY 2026-27</h3>
            <div className="flex gap-2">
              <button onClick={copyDigest} className="px-3 py-1 bg-gray-200 rounded text-sm">Copy Digest</button>
              <button onClick={downloadJSON} className="px-4 py-2 bg-green-600 text-white rounded text-sm">Download .json</button>
            </div>
          </div>
          <div className="mb-4 p-3 bg-gray-100 rounded text-sm font-mono break-all">
            <strong>Digest (CreationInfo.Digest):</strong> {jsonResult.digest}
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <h4 className="font-semibold text-sm mb-2">Tax Computation Overview</h4>
              <pre className="text-xs bg-gray-50 p-3 rounded overflow-auto max-h-[400px]">{JSON.stringify(jsonResult.tax_computation, null, 2)}</pre>
            </div>
            <div>
              <h4 className="font-semibold text-sm mb-2">Official JSON (truncated preview, full download available) – Structure from ITR-1_2026_Main_V1.1.json 64 definitions</h4>
              <pre className="text-xs bg-gray-900 text-green-400 p-3 rounded overflow-auto max-h-[400px]">{jsonResult.json_string.substring(0,5000)}{jsonResult.json_string.length>5000 ? '\n... truncated ...' : ''}</pre>
            </div>
          </div>
          <div className="mt-4 p-3 bg-blue-50 rounded text-xs text-blue-800">
            <strong>Compliance:</strong> JSON must validate against official schema {jsonResult.itr_form}_2026_Main_V1.1.json (64-287 definitions). This builder uses CreationInfo (SWVersionNo, SWCreatedBy, JSONCreatedBy, JSONCreationDate, IntermediaryCity, Digest) – mirrors EfilingCommon.bas getSWVersionNo etc. Final JSON ready for upload to incometax.gov.in (after adding TRP if applicable).
          </div>
        </div>
      )}

      <div className="bg-white rounded-xl shadow-sm p-6">
        <h3 className="font-semibold mb-2">Import Existing JSON / Prefill (PreFillJson.bas 4430 lines, ImportJson.bas 6763 lines)</h3>
        <input type="file" accept=".json" className="text-sm" onChange={async (e)=>{
          const file=e.target.files[0]
          if(!file) return
          const text=await file.text()
          try{
            const jsonData=JSON.parse(text)
            const res=await itrAPI.importJSON(jsonData)
            alert(`Imported ${res.data.itr_form} – preview logged`)
            console.log(res.data)
          }catch(err){alert('Invalid JSON: '+err.message)}
        }}/>
        <p className="text-xs text-gray-500 mt-2">Supports prefill JSON from dept (Base64 decode via DecodeBase64), XML import (SalaryXMLImport, ReliefXMLImport), and previous year Excel ImportPreviousVersion (InsertRowsToImport).</p>
      </div>
    </div>
  )
}
