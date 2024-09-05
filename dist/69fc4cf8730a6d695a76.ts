import { featureUtils } from "../../library/featureUtils";
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("backButton").onclick = featureUtils.goBack;
    document.getElementById("selectAll").onclick = () => featureUtils.selectAll("selectAll", "employeeSelect");
    featureUtils.populateDropdown(featureUtils.getMonthArray(), "monthSelect");
    featureUtils.loadPieChart;
  }
});