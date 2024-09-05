import { featureUtils } from "../../library/featureUtils";


Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("backButton").onclick = featureUtils.goBack;
  }
});

/* global Excel, Office */

let cancelAllOperations = false;

let entries = [];
const months = [
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



const helperFunctions = {
  findIndexByName: (headers, columnName) => headers.indexOf(columnName),
  // only correct for 2024
  returnworkdays2024:() => {
    const workdays = [22,21,20,21,19,20,23,21,21,22,20,20];
    let index = months.indexOf(document.getElementById("monthSelect").value);
    return workdays[index];
  },
  returnworkdays2025:() => {
    let index = months.indexOf(document.getElementById("monthSelect").value);
    const workdays = [21,20,21,20,20,19,23,20,22,22,20,21];
    return workdays[index];
  },
  convertTimeToPT: (data, columnIndex) => {
    let timeColumn = data.map((row) => row[columnIndex]);
    let convertedColumn = [];
    for (let i = 0; i<timeColumn.length; i++) {
      let time = toString(timeColumn[i]);
      let parts = time.split(":");
      let hours = parseInt(parts[0]);
      let minutes = parseInt(parts[1]);
      let pt = hours * 0.125 + minutes * (0.125/60);
      convertedColumn.push(pt);
      
    }
    let updatedData = replaceColumn(data,columnIndex, convertedColumn);

    return updatedData;
  },


};

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    populateDropdowns();
    loadEntriesFromSettings();
    populateMonthDropdown();
    document.getElementById("addEntryButton").onclick = addEntry;
    document.getElementById("addFormulaButton").onclick = addFormula;
    document.getElementById("formulaInput").oninput = (event) => resizeTextarea(event.target);
    document.getElementById("performAllCalculations").onclick = performAllCalculations;
    document.getElementById("autodetectMonth").onclick = autoDetectButtonAction;
    autoDetectButtonAction();
  }
});

function replaceColumn(data, columnIndex, newColumn) {
  return data.map((row, index) => {
    row[columnIndex] = newColumn[index];
    return row;
  });
}

async function populateDropdowns() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Kreis TRADE - 2024");
      const rangeC = sheet.getRange("C:C").getUsedRange();
      const rangeD = sheet.getRange("D:D").getUsedRange();
      const rangeE = sheet.getRange("E:E").getUsedRange();
      rangeC.load("values");
      rangeD.load("values");
      rangeE.load("values");

      await context.sync();

      const employeesSheet = context.workbook.worksheets.getItem("ExportDaten");

      const rangeEmployees = employeesSheet.getRange("E:E").getUsedRange();
      const rangeColumns = employeesSheet.getRange("1:1").getUsedRange();
      rangeEmployees.load("values");
      rangeColumns.load("values");

      await context.sync();

      const employees = ["Alle", ...(await getEmployeeNames(context))];

      const columnEEntries = [...new Set(rangeE.values.flat())];

      populateDropdown("employeeSelect", employees);
      populateDropdown("columnESelect", columnEEntries);
    });
  } catch (error) {
    console.error(error);
  }
}

/**
 *
 * @param {*} context
 * @returns {Promise<string[]>}
 */
async function getEmployeeNames(context) {
  const sheet = context.workbook.worksheets.getItem("ExportDaten");
  const range = sheet.getRange("E:E").getUsedRange();

  range.load("values");

  await context.sync();

  const employees = [...new Set(range.values.flat().slice(1))];

  return employees;
}

function populateDropdown(elementId, values) {
  const select = document.getElementById(elementId);
  select.innerHTML = "";
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.text = value;
    select.appendChild(option);
  });
}

function populateMonthDropdown() {
  const monthSelect = document.getElementById("monthSelect");
  months.forEach((month) => {
    const option = document.createElement("option");
    option.value = month;
    option.text = month;
    monthSelect.appendChild(option);

    monthSelect.addEventListener("change", function () {
      updateSelectedMonthUI();
    });
  });
}

function updateTable() {
  const tbody = document.getElementById("entriesTableBody");
  tbody.innerHTML = "";

  entries.forEach((entry, index) => {
      const row = document.createElement("tr");

      const nameCell = document.createElement("td");
      nameCell.textContent = entry.employee;
      row.appendChild(nameCell);

      const columnECell = document.createElement("td");
      columnECell.textContent = entry.columnE;
      row.appendChild(columnECell);

      const formulaCell = document.createElement("td");
      const formulaInput = document.createElement("textarea");
      formulaInput.type = "text";
      formulaInput.value = entry.formula;
      formulaInput.onchange = () => updateFormula(index, formulaInput.value);
      formulaInput.style.height = "150px";
      formulaInput.style.width = "300px";
      formulaInput.style.display = "none";
      formulaInput.oninput = () => resizeTextarea(formulaInput);

      const showFormulaButton = document.createElement("button");
      showFormulaButton.textContent = "Show Formula";
      showFormulaButton.onclick = () => showFormulaInTable(index);
      showFormulaButton.className = "showFormulaButton";

      formulaCell.appendChild(showFormulaButton);
      formulaCell.appendChild(formulaInput);
      row.appendChild(formulaCell);

      // Actions cell
      const actionsCell = document.createElement("td");

      const editButton = document.createElement("button");
      editButton.textContent = "Edit";
      editButton.onclick = () => toggleEditMode(index, editButton);
      actionsCell.appendChild(editButton);

      const removeButton = document.createElement("button");
      removeButton.textContent = "Remove";
      removeButton.onclick = () => removeEntry(index);
      actionsCell.appendChild(removeButton);

      const testButton = document.createElement("button");
      testButton.textContent = "Run";
      testButton.onclick = () => {
          onRunButtoncClick(index);
      };

      const pipelineCheckbox = document.createElement("input");
      pipelineCheckbox.type = "checkbox";
      pipelineCheckbox.checked = entry.checked; 
      pipelineCheckbox.onchange = () => updateCheckbox(index, pipelineCheckbox.checked);

      actionsCell.appendChild(testButton);
      actionsCell.appendChild(pipelineCheckbox);

      row.appendChild(actionsCell);

      tbody.appendChild(row);
  });
}


function showFormulaInTable(index) {
  const row = document.querySelectorAll("tbody tr")[index]; 
  const formulaInput = row.querySelector("textarea"); 
  const showFormulaButton = row.querySelector(".showFormulaButton"); 
  
  if (formulaInput.style.display === "block") {
    formulaInput.style.display = "none";
    showFormulaButton.textContent = "Show Formula";
  } else {
    formulaInput.style.display = "block";
    showFormulaButton.textContent = "Hide Formula";
  }
}


function resizeTextarea(textarea) {
  textarea.style.height = "auto";
  textarea.style.height = textarea.scrollHeight + "px";
}

function addEntry() {
  const employee = document.getElementById("employeeSelect").value;
  const columnE = document.getElementById("columnESelect").value;

  const entry = {
    employee: employee,
    columnE: columnE,
    formula: "",
    checked: false,
  };

  entries.push(entry);
  updateTable();
  saveEntriesToSettings(); // Save entries to settings
}

function addFormula() {
  const employee = document.getElementById("employeeSelect").value;
  const columnE = document.getElementById("columnESelect").value;
  const formula = document.getElementById("formulaInput").value;

  const entry = entries.find((e) => e.employee === employee && e.columnE === columnE);

  if (entry) {
    entry.formula = formula;
    updateTable();
    saveEntriesToSettings(); // Save entries to settings
  } else {
    console.error("Entry not found");
  }
}

function loadEntriesFromSettings() {
  const savedEntries = Office.context.document.settings.get("savedEntries");
  if (savedEntries) {
    entries = savedEntries;
    updateTable();
  }
}

function saveEntriesToSettings() {
  Office.context.document.settings.set("savedEntries", entries);
  Office.context.document.settings.saveAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Failed) {
      console.error("Failed to save settings: " + result.error.message);
    } else {
      console.log("Settings saved.");
    }
  });
}

function updateFormula(index, newFormula) {
  entries[index].formula = newFormula;
  saveEntriesToSettings();
}

function removeEntry(index) {
  entries.splice(index, 1);
  updateTable();
  saveEntriesToSettings();
}

function toggleEditMode(index, button) {

  if (button.textContent === "Edit") {
    button.textContent = "Done";

    // Set the dropdowns and input fields to the corresponding values
    const entry = entries[index];
    document.getElementById("employeeSelect").value = entry.employee;
    document.getElementById("columnESelect").value = entry.columnE;
    document.getElementById("formulaInput").value = entry.formula;

    const input = document.createElement("input");
    input.type = "text";
  } else {
    button.textContent = "Edit";
  }
}

function alignRanges(currentColumnRange, newColumn, context)
  {
    const sheet = context.workbook.worksheets.getItem("Kreis TRADE - 2024");
    const [startAddress, endAddress] = currentColumnRange.address.split(":");
    const newStartAddress = startAddress.replace("C", newColumn);
    const newEndAddress = endAddress.replace("C", newColumn);
  
    const newAdress = `${newStartAddress}:${newEndAddress}`;
    const alignedRange = sheet.getRange(newAdress);
    return alignedRange;
  }

async function findCellByEntry(index, context) {
  const entry = entries[index];

  const [firstName, lastName] = entry.employee.split(" ").map((name) => name.trim());

  const sheet = context.workbook.worksheets.getItem("Kreis TRADE - 2024");
  sheet.load("name");
  const rangeC = sheet.getRange("C:C").getUsedRange();
  rangeC.load("address, values");
  await context.sync();

  const rangeD = alignRanges(rangeC, "D", context);


  rangeD.load("address");
  await context.sync();

  rangeC.load("values");
  rangeD.load("values");

  await context.sync();

  let startRow = -1;
  for (let i = 0; i < rangeC.values.length; i++) {
    const currentFirstName = rangeC.values[i] && typeof rangeC.values[i][0] === 'string' ? rangeC.values[i][0].trim() : "";
    const currentLastName = rangeD.values[i] && typeof rangeD.values[i][0] === 'string' ? rangeD.values[i][0].trim() : "";
    if (currentFirstName === firstName && currentLastName === lastName) {
      startRow = i;
      entries[index].startRow = startRow;
      break;
    }
  }

  return startRow;
}

async function loadExportDaten(context) {
  const employeesSheet = context.workbook.worksheets.getItem("ExportDaten");
  const range = employeesSheet.getUsedRange();
  range.load("values");
  await context.sync();
  return range.values;
}

async function findCellByName(name, context) {
  const firstName = name.split(" ")[0];
  const lastName = name.split(" ")[1];

  const sheet = context.workbook.worksheets.getItem("Kreis TRADE - 2024");
  sheet.load("name");


  const rangeC = sheet.getRange("C:C").getUsedRange();
  rangeC.load("address, values");
  await context.sync();

  const rangeD = alignRanges(rangeC, "D", context);

  // TODO this loading (and syncing) is necessary once, but not n times if we want to use the "All" feature.
  // We should load the values once and then use them in the findCellByName function.
  // Presumably, the function does not need to be async anymore.
  rangeC.load("values");
  rangeD.load("values");

  await context.sync();

  // await context.sync();

  let startRow = -1;
  for (let i = 0; i < rangeC.values.length; i++) {
    const currentFirstName = rangeC.values[i] && typeof rangeC.values[i][0] === 'string' ? rangeC.values[i][0].trim() : "";
    const currentLastName = rangeD.values[i] && typeof rangeD.values[i][0] === 'string' ? rangeD.values[i][0].trim() : "";
    

    if (currentFirstName === firstName && currentLastName === lastName) {
        startRow = i;
        // Further processing
        break;
    }
}


  if (startRow === -1) {
    // console.log(`Name not found`);
  }
  // console.log(startRow);
  return startRow;
}

/**
 *
 * @param {number} index
 */
async function onRunButtoncClick(index) {
  cancelAllOperations = false;

  // call findCellByName, load data from the ExportDaten sheet, execute formula, findMatchingRowAndWriteSum
  await Excel.run(async (context) => {
    const entry = entries[index];
    const data = await loadExportDaten(context);
    const employeeNames = await getEmployeeNames(context);

    if (entry.employee === "Alle") {
   
      // debugging code
      // const employeeCellPositions = await findAllRelevantCells(context);
      // console.log("employeepositions", employeeCellPositions);
      // console.log("selectedMonth", document.getElementById("monthSelect").value);
      console.log("Starting calculations for all employees.");
      for (let employeeName of employeeNames) {
        if (cancelAllOperations) {
          console.log("User cancelled all operations.");
          return;
        }

        let startRow = await findCellByName(employeeName, context);
        if (startRow != -1) {
          let result = await executeFormula(entry, data, employeeName, helperFunctions);
          // console.log("result", result);
          await findMatchingRowAndWriteResult(context, entry, startRow, result);
        }
      }
      console.log("Finished calculations for all employees.");
    } else {
      const startRow = await findCellByEntry(index, context);

      if (startRow !== -1 && data) {
        const result = await executeFormula(entry, data, entry.employee, helperFunctions);
        await findMatchingRowAndWriteResult(context, entry, startRow, result);
      }
    }

  });
}

// /**
//  *
//  * @param {*} context
//  * @returns {Promise<number[]>}
//  */
// async function findAllRelevantCells(context) {
//   const employeeNames = await getEmployeeNames(context);
//   const matchingRows = [];

//   console.log(employeeNames);

//   for (const name of employeeNames) {
//     matchingRows.push(await findCellByName(name, context));
//   }
//   return matchingRows;
// }

async function displayLoadingMessage() {
  const loadingMessage = document.getElementById("loadingMessage");
  loadingMessage.style.display = "block";
}

async function hideLoadingMessage() {
  const loadingMessage = document.getElementById("loadingMessage");
  loadingMessage.style.display = "none";
}

async function updateProgressIndicator(progress) {
  const progressIndicator = document.getElementById("progressIndicator");
  progressIndicator.textContent = progress;
}


async function executeFormula(entry, data, employeeName, helperFunctions = {}) {
  try {
    // console.log("Helper Functions inside executeFormula:", helperFunctions);

    // Use the formula from the entry
    const formula = `
      ${entry.formula}
    `;

    // Convert formula string to a function that also includes the helper functions
    const formulaFunction = new Function("data", "employeeName", "helpers", formula);

    // Execute the formula with the data, employeeName, and helperFunctions object
    return formulaFunction(data, employeeName, helperFunctions);
  } catch (error) {
    console.error("Error executing formula: ", error);
  }
}

async function findMatchingRowAndWriteResult(context, columnEEntry, startRow, sum) {
  const sheet = context.workbook.worksheets.getItem("Kreis TRADE - 2024");
  const rangeE = sheet.getRange("E:E").getUsedRange();
  rangeE.load("values");
  await context.sync();

  // Declare months locally
  let months = [
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

  const monthSelect = document.getElementById("monthSelect").value;
  const columnIndex = months.indexOf(monthSelect) + 6; // Assuming month columns start from column G (index 6)



  if (columnIndex === -1) {
    console.error("Selected month not found in months array.");
    return;
  }

  for (let j = startRow; j < rangeE.values.length; j++) {
    const currentE = rangeE.values[j][0].trim();

    if (cancelAllOperations) {
      console.log("User cancelled all operations.");
      return;
    }

    if (currentE === columnEEntry.columnE) {
      const cell = sheet.getRangeByIndexes(j + 4, columnIndex, 1, 1);
      cell.load("address, values");
      
      if (!document.getElementById("overWriteAllCheckbox").checked) {

      await context.sync();

      if (cell.values[0][0] !== null && cell.values[0][0] !== "") {
        
        const overwrite = await showOverwriteWarning(cell.address);
        if (!overwrite) {
          console.log("User cancelled the operation.");
          return;
        }
        
      }

    }

      cell.values = [[sum]];
      await context.sync();
      console.log(`Wrote sum ${sum} to cell ${cell.address}`);
      break;
    }
  }

}


function showCalculateAllConfirmMessage() {
  return new Promise((resolve) => {
      const rows = document.querySelectorAll("tbody tr");
      let selectedCategories = [];
      let selectedMonth = document.getElementById("monthSelect").value;

      // Collect the selected categories
      rows.forEach((row) => {
          const checkbox = row.querySelector('input[type="checkbox"]');
          if (checkbox.checked) {
              const category = row.querySelector("td:nth-child(2)").textContent;
              selectedCategories.push(category);
          }
      });

      // Construct the confirmation message
      const message = `You are about to calculate for the month <strong>${selectedMonth}</strong>:<br/><br/>${selectedCategories.map(category => `${category}<br/>`).join("")} <br/> Are you sure you want to proceed? <br/> <br/> 
      Are u ABSOLUTELY sure? <br/> <br/> If you already messed up, luckily excel-online version control has your back, just search for "Versionsverlauf" and reset the Excel to before you pressed this button. <br/>`;


      // Display the confirmation message in the modal
      document.getElementById("confirmationMessage").innerHTML = message;
      document.getElementById("confirmationModal").style.display = "block";
      document.getElementById("modalOverlay").style.display = "block";

      // Handle the OK button
      document.getElementById("confirmButton").onclick = () => {
          document.getElementById("confirmationModal").style.display = "none";
          document.getElementById("modalOverlay").style.display = "none";
          resolve(true);
      };

      // Handle the Cancel button
      document.getElementById("cancelButton").onclick = () => {
          document.getElementById("confirmationModal").style.display = "none";
          document.getElementById("modalOverlay").style.display = "none";
          resolve(false);
      };

  


  });
}


async function performAllCalculations() {
  const rows = document.querySelectorAll("tbody tr");



  const userConfirmed = await showCalculateAllConfirmMessage();
  if (!userConfirmed ) {
    console.log("User cancelled the operation.");
    return;
  }
  

  displayLoadingMessage();



  let numberOfCalculations = 0;


  rows.forEach(async (row) => {
    const checkbox = row.querySelector('input[type="checkbox"]');
    if (checkbox.checked) {
      numberOfCalculations++;
    }
  });

  let currentCalculation = 0;

  for (let index = 0; index < rows.length; index++) {



    const row = rows[index];
    const checkbox = row.querySelector('input[type="checkbox"]');
    if (checkbox.checked) {
      currentCalculation++;
      if (cancelAllOperations) {
        console.log("User cancelled all operations.");
        continue;
      }
      await onRunButtoncClick(index);
      await updateProgressIndicator(`Calculating ${currentCalculation} of ${numberOfCalculations}`);
    }
  }
  hideLoadingMessage();
  currentCalculation = 0;
  cancelAllOperations = false;
}

async function updateCheckbox(index, isChecked) {
  entries[index].checked = isChecked;
  saveEntriesToSettings();
}

async function findMostFrequentMonth() {
  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("ExportDaten");
      const headerRange = sheet.getRange("A1:Z1");
      headerRange.load("values");
      await context.sync();

      let dateColumnIndex = -1;
      const headers = headerRange.values[0];
      for (let i = 0; i < headers.length; i++) {
        const headerText = headers[i].toLowerCase();
        if (headerText.includes("datum") || headerText.includes("date")) {
          dateColumnIndex = i;
          break;
        }
      }

      if (dateColumnIndex === -1) {
        console.error("Date or Datum column not found.");
        return null;
      }

      const usedRange = sheet.getUsedRange();
      usedRange.load("rowCount");
      await context.sync();

      const dateColumnRange = sheet.getRangeByIndexes(1, dateColumnIndex, usedRange.rowCount - 1, 1);
      dateColumnRange.load("values");
      await context.sync();

      const dateCounts = {};

      dateColumnRange.values.forEach((row) => {
        const dateValue = row[0];
        if (dateValue) {
          const parts = dateValue.split('.');
          const formattedDate = `${parts[1]}/${parts[0]}/${parts[2]}`;
          const date = new Date(formattedDate);

          if (!isNaN(date.getTime())) {
            const month = String(date.getMonth() + 1).padStart(2, '0');
            const year = date.getFullYear();
            const monthYear = `${month}/${year}`;

            dateCounts[monthYear] = (dateCounts[monthYear] || 0) + 1;
          } else {
            console.warn("Invalid date encountered:", dateValue);
          }
        }
      });

      let mostFrequentMonthYear = null;
      let maxCount = 0;

      for (const [monthYear, count] of Object.entries(dateCounts)) {
        if (count > maxCount) {
          maxCount = count;
          mostFrequentMonthYear = monthYear;
        }
      }

      if (mostFrequentMonthYear) {
        const [monthNumber] = mostFrequentMonthYear.split('/');
        const monthIndex = parseInt(monthNumber, 10) - 1;

        if (monthIndex >= 0 && monthIndex < months.length) {
          const mostFrequentMonthName = months[monthIndex];
          // console.log(`The most frequent month is: ${mostFrequentMonthName}`);
          return mostFrequentMonthName;
        }
      }

      console.log("No valid dates found or no frequent month detected.");
      return null;
    });
  } catch (error) {
    console.error("An error occurred in findMostFrequentMonth:", error);
    return null;
  }
}

async function autoDetectButtonAction() {
  let mostFrequentMonth = await findMostFrequentMonth();
  
  const monthSelect = document.getElementById("monthSelect");
  // console.log("Options in monthSelect:", monthSelect.options);
  

  let found = false;
  for (let i = 0; i < monthSelect.options.length; i++) {
      if (monthSelect.options[i].value === mostFrequentMonth) {
          found = true;
          break;
      }
  }

  if (found) {
      monthSelect.value = mostFrequentMonth;
  } else {
      console.error("No matching option found for:", mostFrequentMonth);
  }
}

async function showOverwriteWarning(position) {
  return new Promise((resolve) => {
      // Construct the warning message
      const message = `Do you want to overwrite the value in <strong>${position}</strong>?`;

      // Display the warning message in the modal
      document.getElementById("confirmationMessage").innerHTML = message;
      document.getElementById("confirmationModal").style.display = "block";
      document.getElementById("modalOverlay").style.display = "block";

      // Handle the OK button
      document.getElementById("confirmButton").onclick = () => {
          document.getElementById("confirmationModal").style.display = "none";
          document.getElementById("modalOverlay").style.display = "none";
          resolve(true);
      };

      // Handle the Cancel button
      document.getElementById("cancelButton").onclick = () => {
          document.getElementById("confirmationModal").style.display = "none";
          document.getElementById("modalOverlay").style.display = "none";
          resolve(false);
      };
      // Handle the CancelAll button
      document.getElementById("cancelAllButton").onclick = () => {
        document.getElementById("confirmationModal").style.display = "none";
        document.getElementById("modalOverlay").style.display = "none";
        cancelAllOperations = true;
        console.log("Cancel All button clicked.");
        resolve(false);

        
    };
  });
}