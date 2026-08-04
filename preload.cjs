const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('db', {
  login: (email, password) => ipcRenderer.invoke('auth:login', email, password),
  signup: (signupData) => ipcRenderer.invoke('auth:signup', signupData),
  logout: () => ipcRenderer.invoke('auth:logout'),
  getCurrentUser: () => ipcRenderer.invoke('auth:current-user'),
  fetchData: () => ipcRenderer.invoke('db:fetch-data'),
  saveProducts: (products) => ipcRenderer.invoke('db:save-products', products),
  saveInvoices: (invoices) => ipcRenderer.invoke('db:save-invoices', invoices),
  saveExpenses: (expenses) => ipcRenderer.invoke('db:save-expenses', expenses),
  saveSettings: (settings) => ipcRenderer.invoke('db:save-settings', settings),
  sendInvoiceEmail: (invoice, user) => ipcRenderer.invoke('db:send-invoice-email', invoice, user),
  requestResetOtp: (email) => ipcRenderer.invoke('auth:request-reset-otp', email),
  confirmPasswordReset: (email, otp, newPassword) => ipcRenderer.invoke('auth:reset-password-confirm', email, otp, newPassword),
});
