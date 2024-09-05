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

// featureUtils.loadPieChart = function () {
//   const xValues = ["Employee 1", "Employee 2", "Employee 3"];
//   const yValues = [55, 30, 25];
//   const barColors = ["#b91d47", "#00aba9", "#2b5797"];

//   new Chart("PieChart", {
//     type: "pie",
//     data: {
//       labels: xValues,
//       datasets: [
//         {
//           backgroundColor: barColors,
//           data: yValues,
//         },
//       ],
//     },
//     options: {
//       title: {
//         display: true,
//         text: "Gebuchte Tage",
//       },
//     },
//   });
// };
