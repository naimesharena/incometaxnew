import axios from 'axios'

const API_BASE = import.meta.env.VITE_API_BASE ?? ''  // use relative URL for preview proxy; set VITE_API_BASE=http://localhost:8000 for local dev if needed, or '' for proxied

const api = axios.create({
  baseURL: API_BASE || '',
  timeout: 30000,
})

export const masterAPI = {
  getBankCodes: (q, limit=20) => api.get(`/api/master/bank-codes?q=${q}&limit=${limit}`),
  getPincode: (pin) => api.get(`/api/master/pincode/${pin}`),
  validateIFSC: (ifsc) => api.get(`/api/master/ifsc/validate?ifsc=${ifsc}`),
  searchIFSC: (q, limit=20) => api.get(`/api/master/ifsc/search?q=${q}&limit=${limit}`),
  getStates: () => api.get('/api/master/states'),
  getEmployerCategories: () => api.get('/api/master/employer-categories'),
  getDropdowns: () => api.get('/api/master/dropdowns'),
  getFileSections: () => api.get('/api/master/return-file-sections'),
  getHashMeta: () => api.get('/api/master/hash-meta'),
}

export const itrAPI = {
  calculate: (payload) => api.post('/api/itr/calculate', payload),
  validate: (payload) => api.post('/api/itr/validate', payload),
  generateJSON: (payload) => api.post('/api/itr/generate-json', payload),
  getForms: () => api.get('/api/itr/forms'),
  importJSON: (jsonData) => api.post('/api/itr/import-json', {json: jsonData}),
}

export default api
