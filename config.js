// 永續 KPI 管理平台｜設定檔
// 這裡的值都不是密碼或機密（Client ID、Tenant ID 本來就會出現在登入網址中），可以放在網站上。
// 請依「設定步驟」把三個值填進引號內。
window.KPI_CONFIG = {
  // 步驟 3 取得：Microsoft Entra 應用程式註冊 → 概觀 →「目錄 (租用戶) 識別碼」
  tenantId: "4b03cc2e-4e82-4019-9150-47b2899c7495",
  // 步驟 3 取得：Microsoft Entra 應用程式註冊 → 概觀 →「應用程式 (用戶端) 識別碼」
  clientId: "ad6d8c70-4808-4c0f-8e7a-59bfe9612cb6",
  // 步驟 1 取得：Teams 文件庫中「永續KPI資料.xlsx」→ 複製連結（「具有現有存取權的人員」）
  fileShareUrl: "https://planetcomtw.sharepoint.com/:x:/r/sites/msteams_784ed6/Shared%20Documents/General/%E6%B0%B8%E7%BA%8CKPI%E8%B3%87%E6%96%99.xlsx?d=we01ff6fb956b4f7c9e965b5effc4f74a&csf=1&web=1&e=fvVsdv"
};
