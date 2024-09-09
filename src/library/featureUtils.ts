import { Chart } from "chart.js/auto";

export function featureUtils() {}
//ToDo: Add documentation to all functions
//ToDo: maybe add functions only for the pieChart in a seperate library?

/**
 * Get an array with all months
 *
 * @returns Array with months
 */
featureUtils.getMonthArray = function () {
  return [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
  ];
};

featureUtils.getTestEmployeeArray = function () {
  return ["All Employees", "Employee 1", "Employee 2", "Employee 3"];
};

/**
 * Populate a dropdown
 *
 * @param pValueArray     Array with Values that should be displayed in the dropdown
 * @param pElementId      ID of the dropdown
 */
featureUtils.populateDropdown = function (pValueArray: string[], pElementId: string) {
  const element = document.getElementById(pElementId);
  pValueArray.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.text = value;
    element.appendChild(option);
  });
};

/**
 * Return back to main screen
 *
 */
featureUtils.goBack = function () {
  window.location.href = "../taskpane.html";
};

/**
 * Create a select all checkbox
 *
 * @param pSource    Id of the source checkbox (the 'select all' box)
 * @param pElement   Name of the checkboxes that will be checked
 */
featureUtils.selectAll = function (pSource: string, pElement: string) {
  var source: HTMLElement = document.getElementById(pSource);

  if (source instanceof HTMLInputElement) {
    let checkboxes = document.getElementsByName(pElement);

    for (let i = 0; i < checkboxes.length; i++) {
      const currentCheckbox = checkboxes[i];

      if (currentCheckbox instanceof HTMLInputElement) {
        currentCheckbox.checked = source.checked;
      }
    }
  }
};

/**
 * Loads the basic pie chart with no data
 * @returns the pie chart
 */
featureUtils.loadPieChart = function () {
  const xValues: string[] = [];
  const yValues: number[] = [];
  //ToDo one day: more colors, maybe a randomized function?
  const barColors: string[] = ["#8BC1F7", "#BDE2B9", "#A2D9D9", "#B2B0EA", "#F9E0A2", "#F4B678"];

  const data = {
    datasets: [
      {
        data: yValues,
        backgroundColor: barColors,
      },
    ],
    hoverOffset: 4,
    labels: xValues,
  };

  var pieChart = new Chart("pieChart", {
    type: "pie",
    data: data,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: "top",
        },
        title: {
          display: true,
          text: "Time",
        },
      },
    },
  });

  return pieChart;
};

/**
 * Add Data to a pie chart
 *
 * @param pChart chart where data should be added
 * @param pLabel label of the datasets
 * @param pNewData new data that should be added
 */
featureUtils.addPieChartData = function (pChart: Chart<"pie", number[], string>, pLabel, pNewData) {
  //push new label
  pChart.data.labels.push(pLabel);

  //push data
  pChart.data.datasets.forEach((dataset) => {
    dataset.data.push(pNewData);
  });
  pChart.update();
};

/**
 * Remove data from a pie chart
 *
 * @param pChart chart where data should be removed
 * @param pLabel label of the datasets
 * @param pData data that should be removed
 */
featureUtils.removePieChartData = function (pChart: Chart<"pie", number[], string>, pLabel, pData) {
  var data = pChart.data.datasets[0].data;
  var labels = pChart.data.labels;
  var removalIndexLabel = labels.indexOf(pLabel);
  var removalIndexData = data.indexOf(pData);

  //remove label
  labels.splice(removalIndexLabel, 1);

  //remove data
  pChart.data.datasets.forEach((dataset) => {
    dataset.data.splice(removalIndexData, 1);
  });

  pChart.update();
};

/**
 * Adds function to checkboxes that remove or add data
 * @param pSource id of the checkbox
 * @param pChart chart where data should be added/removed
 * @param pLabel label of the data
 * @param pData the data that should be added/removed
 */
featureUtils.addRemoveDataCheckbox = function (pSource: string, pChart: Chart<"pie", number[], string>, pLabel, pData) {
  var source: HTMLElement = document.getElementById(pSource);

  source.addEventListener("change", function () {
    if (source instanceof HTMLInputElement) {
      if (source.checked) {
        featureUtils.addPieChartData(pChart, pLabel, pData);
      } else {
        featureUtils.removePieChartData(pChart, pLabel, pData);
      }
    }
  });
};

featureUtils.configureAllCheckboxes = function (pChart: Chart<"pie", number[], string>) {
  featureUtils.addRemoveDataCheckbox("entryOption1", pChart, "gebuchte Tage", 80);
  featureUtils.addRemoveDataCheckbox("entryOption2", pChart, "fakturierbare Tage", 50);
  featureUtils.addRemoveDataCheckbox("entryOption3", pChart, "Urlaub", 10);
  featureUtils.addRemoveDataCheckbox("entryOption4", pChart, "Krank", 1);
};