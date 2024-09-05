import { Chart } from "chart.js/auto";

export function featureUtils() {}
//ToDo: Add documentation to all functions

/**
 * Get an array with all months
 *
 * @returns {String[]}  Array with months
 */
featureUtils.getMonthArray = function () {
  const months: string[] = [
    "Januar",
    "Februar",
    "März",
    "April",
    "Mai",
    "Juni",
    "Juli",
    "August",
    "September",
    "Oktober",
    "November",
    "Dezember",
  ];

  return months;
};

/**
 * Populate a dropdown
 *
 * @param {string[]} pValueArray     Array with Values that should be displayed in the dropdown
 * @param {string} pElementId        ID of the dropdown
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
 * @param {string} pSource    Id of the source checkbox (the 'select all' box)
 * @param {string} pElement   Name of the checkboxes that will be checked
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

featureUtils.loadPieChart = function () {
  const xValues: string[] = ["Employee 2", "Employee 3"];
  const yValues: number[] = [75, 25];
  const barColors: string[] = ["#8BC1F7", "#BDE2B9", "#A2D9D9", "#B2B0EA", "#F9E0A2", "#F4B678"];

  const data = {
    lables: xValues,
    datasets: [
      {
        data: yValues,
        backgroundColor: barColors,
      },
    ],
    hoverOffset: 4,
  };

  var pieChart = new Chart("pieChart", {
    type: "pie",
    data: data,
    options: {
      responsive: true,
      plugins: {
        legend: {
          position: "top",
        },
        title: {
          display: true,
          text: "Gebuchte Zeit",
        },
      },
    },
  });

  return pieChart;
};

/**
 *
 * @param pChart
 */
featureUtils.addPieChartData = function (pChart: Chart, pLabel: string, pNewData: string[]) {
  pChart.data.labels.push(pLabel);
  pChart.data.datasets.forEach((dataset) => {
    dataset.data.push();
  });
  pChart.update();
};

featureUtils.removePieChartData = function (pChart: Chart, pData: string|number) {
  let data = pChart.data;;
  let removalIndex = data.datasets.indexOf(pData)
  pChart.data.datasets.forEach((dataset) => {
    dataset.data.pop(pData);
  });
  pChart.update();
};

featureUtils.addRemoveDataCheckbox = function (pSource: string) {
  var source: HTMLElement = document.getElementById(pSource);

  source.addEventListener("change", function () {
    if (source instanceof HTMLInputElement) {
      if (source.checked) {
        featureUtils.addPieChartData();
      } else {
        featureUtils.removePieChartData();
      }
    }
  });
};
