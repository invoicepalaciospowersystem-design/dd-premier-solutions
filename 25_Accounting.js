// =====================================================
// FILE: 25_Accounting.gs
// =====================================================

function getAccountingDashboardData(sessionToken, companyId, periodMode, periodYear, periodMonth) {
  return getEconomyModuleDashboard(sessionToken, companyId, periodMode, periodYear, periodMonth);
}
