const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  getSettings: () => ipcRenderer.invoke('settings:get'),
  saveSettings: (settings) => ipcRenderer.invoke('settings:save', settings),
  selectRootFolder: () => ipcRenderer.invoke('settings:selectRootFolder'),
  selectFiscalRootFolder: () => ipcRenderer.invoke('settings:selectFiscalRootFolder'),
  listUsers: () => ipcRenderer.invoke('users:list'),
  createUser: (payload) => ipcRenderer.invoke('users:create', payload),
  updateUser: (payload) => ipcRenderer.invoke('users:update', payload),
  setUserPassword: (payload) => ipcRenderer.invoke('users:setPassword', payload),
  deleteUser: (payload) => ipcRenderer.invoke('users:delete', payload),
  verifyLogin: (payload) => ipcRenderer.invoke('auth:verify', payload),
  loadList: (payload) => ipcRenderer.invoke('list:load', payload),
  loadFreightSummary: (payload) => ipcRenderer.invoke('freight:summary', payload),
  loadPcpEfficiency: (payload) => ipcRenderer.invoke('pcp:efficiency', payload),
  exportPcpEfficiencyImage: (payload) => ipcRenderer.invoke('pcp:efficiency:export-image', payload),
  loadPcpDashboard: (payload) => ipcRenderer.invoke('pcp:dashboard', payload),
  loadPcpMoldagem: (payload) => ipcRenderer.invoke('pcp:moldagem', payload),
  exportPcpDashboardPdf: (payload) => ipcRenderer.invoke('pcp:dashboard:export-pdf', payload),
  fiscalListNfs: () => ipcRenderer.invoke('fiscal:list-nfs'),
  fiscalGetNf: (payload) => ipcRenderer.invoke('fiscal:get-nf', payload),
  fiscalPreviewItems: (payload) => ipcRenderer.invoke('fiscal:preview-items', payload),
  fiscalRegisterNf: (payload) => ipcRenderer.invoke('fiscal:register-nf', payload),
  fiscalUpdateNf: (payload) => ipcRenderer.invoke('fiscal:update-nf', payload),
  fiscalDeleteNf: (payload) => ipcRenderer.invoke('fiscal:delete-nf', payload),
  fiscalHistory: () => ipcRenderer.invoke('fiscal:history'),
  fiscalFindNf: (payload) => ipcRenderer.invoke('fiscal:find-nf', payload)
});
