// Types for entries and other elements
interface Entry {
  employee: string;
  columnE: string;
  formula: string;
  checked: boolean;
}

interface Formula {
  employee: string;
  columnE: string;
  formula: string;
}

let entries: Entry[] = [];

/**
 * Updates the table by rendering the entries into the DOM.
 * Clears the table and adds a row for each entry in the `entries` array.
 */
function updateTable(): void {
  const tbody = document.querySelector("tbody");
  if (!tbody) return;

  tbody.innerHTML = ""; // Clear the table

  entries.forEach((entry, index) => {
    const row = document.createElement("tr");

    // Employee column
    const employeeCell = document.createElement("td");
    employeeCell.textContent = entry.employee;
    row.appendChild(employeeCell);

    // Column E
    const columnECell = document.createElement("td");
    columnECell.textContent = entry.columnE;
    row.appendChild(columnECell);

    // Actions column
    const actionsCell = document.createElement("td");

    // Edit button
    const editButton = document.createElement("button");
    editButton.textContent = "Edit";
    editButton.onclick = () => editEntry(index);
    actionsCell.appendChild(editButton);

    // Remove button
    const removeButton = document.createElement("button");
    removeButton.textContent = "Remove";
    removeButton.onclick = () => removeEntry(index);
    actionsCell.appendChild(removeButton);

    // Run button
    const testButton = document.createElement("button");
    testButton.textContent = "Run";
    testButton.onclick = () => onRunButtonClick(index);
    actionsCell.appendChild(testButton);

    // Checkbox
    const pipelineCheckbox = document.createElement("input");
    pipelineCheckbox.type = "checkbox";
    pipelineCheckbox.checked = entry.checked;
    pipelineCheckbox.onchange = () => updateCheckbox(index, pipelineCheckbox.checked);
    actionsCell.appendChild(pipelineCheckbox);

    row.appendChild(actionsCell);
    tbody.appendChild(row);
  });
}

/**
 * Toggles the visibility of the formula textarea in the table for a specific entry.
 * @param index - The index of the row where the formula textarea should be toggled.
 */
function showFormulaInTable(index: number): void {
  const row = document.querySelectorAll("tbody tr")[index];
  const formulaInput = row.querySelector("textarea") as HTMLTextAreaElement;
  const showFormulaButton = row.querySelector(".showFormulaButton") as HTMLButtonElement;

  if (formulaInput.style.display === "block") {
    formulaInput.style.display = "none";
    showFormulaButton.textContent = "Show Formula";
  } else {
    formulaInput.style.display = "block";
    showFormulaButton.textContent = "Hide Formula";
  }
}

/**
 * Dynamically resizes the height of the textarea to fit its content.
 * @param textarea - The textarea element to resize.
 */
function resizeTextarea(textarea: HTMLTextAreaElement): void {
  textarea.style.height = "auto"; // Reset height
  textarea.style.height = `${textarea.scrollHeight}px`; // Set height based on content
}

/**
 * Adds a new entry to the `entries` array and updates the table.
 */
function addEntry(): void {
  const employee = (document.getElementById("employeeSelect") as HTMLSelectElement).value;
  const columnE = (document.getElementById("columnESelect") as HTMLSelectElement).value;

  const entry: Entry = {
    employee,
    columnE,
    formula: "",
    checked: false,
  };

  entries.push(entry);
  updateTable();
  saveEntriesToSettings(); // Save entries to localStorage
}

/**
 * Adds a new formula entry to the system.
 * This function handles user input for adding a formula.
 */
function addFormula(): void {
  const employee = (document.getElementById("employeeSelect") as HTMLSelectElement).value;
  const columnE = (document.getElementById("columnESelect") as HTMLSelectElement).value;
  const formula = (document.getElementById("formulaInput") as HTMLTextAreaElement).value;

  const newFormula: Formula = { employee, columnE, formula };

  // Add logic to handle the new formula here
}

/**
 * Edits an existing entry at the specified index.
 * @param index - The index of the entry to edit.
 */
function editEntry(index: number): void {
  // Logic for editing an entry
}

/**
 * Removes an entry from the `entries` array and updates the table.
 * @param index - The index of the entry to remove.
 */
function removeEntry(index: number): void {
  entries.splice(index, 1);
  updateTable();
  saveEntriesToSettings(); // Update saved entries after removal
}

/**
 * Updates the checkbox state for a specific entry.
 * @param index - The index of the entry to update.
 * @param isChecked - The new checked state of the checkbox.
 */
function updateCheckbox(index: number, isChecked: boolean): void {
  entries[index].checked = isChecked;
  saveEntriesToSettings(); // Save the updated checkbox state
}

/**
 * Saves the current entries to localStorage to persist them across page reloads.
 */
function saveEntriesToSettings(): void {
  localStorage.setItem("entries", JSON.stringify(entries));
}

/**
 * Displays a confirmation modal for overwriting data.
 * @param position - The name of the position where the value will be overwritten.
 * @returns A Promise that resolves with a boolean indicating the user's choice.
 */
async function showOverwriteWarning(position: string): Promise<boolean> {
  return new Promise((resolve) => {
    const message = `Do you want to overwrite the value in <strong>${position}</strong>?`;

    document.getElementById("confirmationMessage")!.innerHTML = message;
    document.getElementById("confirmationModal")!.style.display = "block";
    document.getElementById("modalOverlay")!.style.display = "block";

    document.getElementById("confirmButton")!.onclick = () => {
      document.getElementById("confirmationModal")!.style.display = "none";
      document.getElementById("modalOverlay")!.style.display = "none";
      resolve(true);
    };

    document.getElementById("cancelButton")!.onclick = () => {
      document.getElementById("confirmationModal")!.style.display = "none";
      document.getElementById("modalOverlay")!.style.display = "none";
      resolve(false);
    };

    document.getElementById("cancelAllButton")!.onclick = () => {
      document.getElementById("confirmationModal")!.style.display = "none";
      document.getElementById("modalOverlay")!.style.display = "none";
      resolve(false);
    };
  });
}

/**
 * Handles the action when the run button is clicked for a specific entry.
 * @param index - The index of the entry to run.
 */
function onRunButtonClick(index: number): void {
  // Logic for running an entry
}
