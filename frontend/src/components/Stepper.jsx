import React from 'react'

const steps = [
  {id: 1, title: 'Personal Info', short: 'Personal'},
  {id: 2, title: 'Income Details', short: 'Income'},
  {id: 3, title: 'House Property', short: 'HP'},
  {id: 4, title: 'Deductions', short: '80C-80U'},
  {id: 5, title: 'TDS / TCS', short: 'TDS/TCS'},
  {id: 6, title: 'Taxes Paid', short: 'Taxes'},
  {id: 7, title: 'Bank Accounts', short: 'Bank'},
  {id: 8, title: 'Summary & Tax', short: 'Summary'},
  {id: 9, title: 'Verification & JSON', short: 'JSON'},
]

export default function Stepper({currentStep, completedSteps = [], onStepClick}) {
  return (
    <div className="w-full bg-white shadow-sm border-b sticky top-0 z-10">
      <div className="max-w-7xl mx-auto px-4 py-3">
        <div className="flex items-center justify-between overflow-x-auto">
          {steps.map((step, idx) => (
            <React.Fragment key={step.id}>
              <div 
                className={`flex flex-col items-center cursor-pointer min-w-[80px] ${currentStep===step.id ? 'opacity-100' : 'opacity-70 hover:opacity-100'}`}
                onClick={()=>onStepClick && onStepClick(step.id)}
              >
                <div className={`w-9 h-9 rounded-full flex items-center justify-center text-sm font-bold transition-all
                  ${currentStep===step.id ? 'step-active shadow-lg scale-110' : completedSteps.includes(step.id) ? 'step-completed' : 'bg-gray-200 text-gray-600'}`}>
                  {completedSteps.includes(step.id) && currentStep!==step.id ? '✓' : step.id}
                </div>
                <span className={`text-xs mt-1 hidden md:block text-center ${currentStep===step.id ? 'font-semibold text-blue-700' : 'text-gray-600'}`}>{step.title}</span>
                <span className={`text-xs mt-1 md:hidden ${currentStep===step.id ? 'font-semibold text-blue-700' : 'text-gray-600'}`}>{step.short}</span>
              </div>
              {idx < steps.length-1 && (
                <div className={`flex-1 h-[2px] mx-1 md:mx-2 ${completedSteps.includes(step.id) ? 'bg-green-500' : 'bg-gray-200'}`}></div>
              )}
            </React.Fragment>
          ))}
        </div>
      </div>
    </div>
  )
}

export { steps }
