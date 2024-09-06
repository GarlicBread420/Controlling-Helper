import { featureUtils } from "../../library/featureUtils";
var pieChart = featureUtils.loadPieChart();
Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("backButton").onclick = featureUtils.goBack;
    document.getElementById("selectAll").onclick = () => featureUtils.selectAll("selectAll", "entries");
    featureUtils.populateDropdown(featureUtils.getTestEmployeeArray(), "employeeSelect");
    featureUtils.populateDropdown(featureUtils.getMonthArray(), "monthSelect");
    pieChart;
    featureUtils.addRemoveDataCheckbox("entryOption1", pieChart, "gebuchte Tage", 80);
    featureUtils.addRemoveDataCheckbox("entryOption2", pieChart, "fakturierbare Tage", 50);
    featureUtils.addRemoveDataCheckbox("entryOption3", pieChart, "Urlaub", 10);
    featureUtils.addRemoveDataCheckbox("entryOption4", pieChart, "Krank", 1);
  }
});