import { featureUtils } from "../../library/featureUtils";

var pieChart = featureUtils.loadPieChart();

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("backButton").onclick = featureUtils.goBack; //configure "Back to Maiin" button
    document.getElementById("selectAll").onclick = () => featureUtils.selectAll("selectAll", "entries"); //configure "Select All" checkbox
    featureUtils.populateDropdown(featureUtils.getMonthArray(), "monthSelect"); //fill "Select Month" dropdown
    pieChart; //load Pie Chart without any data
    dataArray.forEach(([label, data]) => {
      //push data into Pie Chart
      featureUtils.addPieChartData(pieChart, label, data);
    });
    featureUtils.configureAllCheckboxes(pieChart); //configure entry checkboxes
    featureUtils.populateEmployeeDropdown(); //fill "Select Employee" dropdown
  }
});

//Array with test data
var dataArray = [
  ["gebuchte Tage", 80],
  ["fakturierbare Tage", 50],
  ["Urlaub", 10],
  ["Krank", 1],
];

dataArray.forEach(([label, data]) => {
  featureUtils.addRemoveDataCheckbox("selectAll", pieChart, label, data);
});