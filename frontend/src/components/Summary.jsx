import React from 'react'

export default function Summary({computation, validation, onCalculate}) {
  if(!computation){
    return (
      <div className="bg-white rounded-xl shadow-sm p-6 text-center">
        <p className="text-gray-600">Click Calculate to compute tax (mirrors mdTaxCalc + SUMMARY hidden sheet)</p>
        <button onClick={onCalculate} className="mt-4 px-6 py-2 bg-blue-600 text-white rounded-lg">Calculate Tax – New & Old Regime</button>
      </div>
    )
  }

  const inc=computation.income_breakdown
  const oldTax=computation.tax_old_regime
  const newTax=computation.tax_new_regime
  const chosen=computation.chosen_tax
  const interest=computation.interest
  const final=computation.final

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-xl shadow-sm p-6">
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-xl font-bold">Tax Computation Summary – SUMMARY Hidden B2:I31 + Part B ATI 45 rows</h2>
          <button onClick={onCalculate} className="px-4 py-2 bg-blue-600 text-white rounded text-sm">Recalculate</button>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="border rounded-lg p-4">
            <h3 className="font-semibold text-gray-700 mb-2">Income Breakdown</h3>
            <div className="space-y-1 text-sm">
              <div className="flex justify-between"><span>Gross Total Income</span><span className="font-mono">₹{inc.gross_total_income.toLocaleString('en-IN')}</span></div>
              <div className="flex justify-between"><span>HRA Exemption (EA 10(13A) G12)</span><span className="font-mono">-₹{inc.hra_exemption.toLocaleString('en-IN')}</span></div>
              <div className="flex justify-between"><span>Standard Deduction ({inc.standard_deduction_applied})</span><span className="font-mono">-₹{inc.standard_deduction_applied.toLocaleString('en-IN')}</span></div>
              <div className="flex justify-between"><span>Taxable Salary After Exemptions</span><span className="font-mono">₹{inc.taxable_salary_after_exemptions.toLocaleString('en-IN')}</span></div>
              <div className="flex justify-between"><span>Gross After Salary Adjustment</span><span className="font-mono">₹{inc.gross_after_salary_adjustment.toLocaleString('en-IN')}</span></div>
              <div className="flex justify-between"><span>Total Deductions Old Regime (80C+80D+80G...)</span><span className="font-mono">-₹{inc.total_deductions_old.toLocaleString('en-IN')}</span></div>
              <div className="flex justify-between border-t pt-1 font-bold"><span>Total Income OLD</span><span>₹{inc.total_income_old.toLocaleString('en-IN')}</span></div>
              <div className="flex justify-between font-bold text-blue-700"><span>Total Income NEW (no 80 deductions except std)</span><span>₹{inc.total_income_new.toLocaleString('en-IN')}</span></div>
            </div>
          </div>

          <div className="border rounded-lg p-4">
            <h3 className="font-semibold text-gray-700 mb-2">Regime Comparison – AY 26-27 Slabs</h3>
            <div className="grid grid-cols-2 gap-4 text-sm">
              <div className="bg-gray-50 p-3 rounded">
                <div className="font-bold text-gray-800 mb-1">Old Regime</div>
                <div>Taxable: ₹{oldTax.taxable_income.toLocaleString()}</div>
                <div>Base Tax: ₹{oldTax.base_tax.toLocaleString()}</div>
                <div>Rebate87A: -₹{oldTax.rebate87A.toLocaleString()}</div>
                <div>Surcharge {oldTax.surcharge_rate*100}%: ₹{oldTax.surcharge.toLocaleString()}</div>
                <div>Cess 4%: ₹{oldTax.cess.toLocaleString()}</div>
                <div className="font-bold border-t mt-1 pt-1">Total: ₹{oldTax.total_tax.toLocaleString()}</div>
                <div className="text-xs text-gray-500 mt-1">Slab: 0-2.5 nil, 2.5-5 5%, 5-10 20%, &gt;10 30% (senior 3L/5L)</div>
              </div>
              <div className="bg-blue-50 p-3 rounded border-2 border-blue-200">
                <div className="font-bold text-blue-800 mb-1">New Regime (Default) – 115BAC(1A)</div>
                <div>Taxable: ₹{newTax.taxable_income.toLocaleString()}</div>
                <div>Base Tax: ₹{newTax.base_tax.toLocaleString()}</div>
                <div>Rebate87A (upto 12L): -₹{newTax.rebate87A.toLocaleString()}</div>
                <div>Surcharge {newTax.surcharge_rate*100}%: ₹{newTax.surcharge.toLocaleString()}</div>
                <div>Cess 4%: ₹{newTax.cess.toLocaleString()}</div>
                <div className="font-bold border-t mt-1 pt-1 text-blue-800">Total: ₹{newTax.total_tax.toLocaleString()}</div>
                <div className="text-xs text-gray-600 mt-1">Slab: 0-4 nil, 4-8 5%, 8-12 10%, 12-16 15%, 16-20 20%, 20-24 25%, &gt;24 30% – Std Ded 75k, Rebate 60k till 12L (12.75L salaried tax-free)</div>
              </div>
            </div>
            <div className="mt-3 p-2 bg-green-50 rounded text-sm">
              Chosen Regime: <strong>{computation.chosen_regime}</strong> – Taxable ₹{computation.taxable_income.toLocaleString()} – Tax ₹{chosen.total_tax.toLocaleString()}
            </div>
          </div>
        </div>

        <div className="mt-6 grid grid-cols-1 md:grid-cols-2 gap-6">
          <div className="border rounded-lg p-4">
            <h3 className="font-semibold text-gray-700 mb-2">Prepaid Taxes (TDS + Advance + Self)</h3>
            <div className="text-sm space-y-1">
              <div className="flex justify-between"><span>TDS Credit</span><span>₹{computation.prepaid_taxes.tds.toLocaleString()}</span></div>
              <div className="flex justify-between"><span>Advance Tax</span><span>₹{computation.prepaid_taxes.advance.toLocaleString()}</span></div>
              <div className="flex justify-between"><span>Self Assessment</span><span>₹{computation.prepaid_taxes.self_assessment.toLocaleString()}</span></div>
              <div className="flex justify-between font-bold border-t pt-1"><span>Total Prepaid</span><span>₹{computation.prepaid_taxes.total_prepaid.toLocaleString()}</span></div>
            </div>
          </div>
          <div className="border rounded-lg p-4">
            <h3 className="font-semibold text-gray-700 mb-2">Interest u/s 234A/B/C – mdCalInterst234B.bas 707 lines, mdTaxCalc 1892 lines</h3>
            <div className="text-sm space-y-1">
              <div className="flex justify-between"><span>234A – Delay Filing (1% pm)</span><span>₹{interest['234A'].toLocaleString()}</span></div>
              <div className="flex justify-between"><span>234B – Advance shortfall &lt;90% (1% pm)</span><span>₹{interest['234B'].toLocaleString()}</span></div>
              <div className="flex justify-between"><span>234C – Deferment of Advance</span><span>₹{interest['234C'].toLocaleString()}</span></div>
              <div className="flex justify-between font-bold border-t pt-1"><span>Total Interest</span><span>₹{interest.total.toLocaleString()}</span></div>
            </div>
          </div>
        </div>

        <div className={`mt-6 p-4 rounded-lg text-center ${final.refund>0 ? 'bg-green-100 text-green-800' : 'bg-orange-100 text-orange-800'}`}>
          {final.refund>0 ? (
            <div><div className="text-2xl font-bold">Refund Due: ₹{final.refund.toLocaleString('en-IN')}</div><div className="text-sm">Excess prepaid over tax + interest</div></div>
          ) : (
            <div><div className="text-2xl font-bold">Tax Payable: ₹{final.payable.toLocaleString('en-IN')}</div><div className="text-sm">Total tax with interest ₹{final.total_tax_with_interest.toLocaleString()} – Prepaid ₹{computation.prepaid_taxes.total_prepaid.toLocaleString()}</div></div>
          )}
        </div>
      </div>

      {validation && (
        <div className={`rounded-xl shadow-sm p-4 ${validation.valid ? 'bg-green-50 border border-green-200' : 'bg-red-50 border border-red-200'}`}>
          <h3 className="font-semibold mb-2">Validation – {validation.valid ? '✓ All OK (mirrors fmsgboxStatus)' : `✗ ${validation.count} Error(s)`}</h3>
          {!validation.valid && (
            <ul className="text-sm list-disc ml-5 space-y-1 text-red-700">
              {validation.errors.map((e,i)=><li key={i}>{e}</li>)}
            </ul>
          )}
          {validation.warnings && validation.warnings.length>0 && (
            <ul className="text-sm list-disc ml-5 text-yellow-700 mt-2">
              {validation.warnings.map((w,i)=><li key={i}>{w}</li>)}
            </ul>
          )}
        </div>
      )}
    </div>
  )
}
