import SpreadsheetWs from './ss-ws.js';

import { Result, okResult, errResult, Err } from 'cs544-js-utils';

import { Errors, makeElement } from './utils.js';

const [N_ROWS, N_COLS] = [10, 10];

export default async function make(ws: SpreadsheetWs, ssName: string) {
  return await Spreadsheet.make(ws, ssName);
}


class Spreadsheet {

  private readonly ws: SpreadsheetWs;
  private readonly ssName: string;
  private readonly errors: Errors;

  //TODO: add more instance variables
  private focusedCellId: string | null = null; 
  private isFocused = false;
  private currentlyFocusedCellId: string | null = null;

  // Instance variables to store copy source cell information
  private copySourceCellId: string | null = null;
  private copySourceExpr: string | null = null;
  
  constructor(ws: SpreadsheetWs, ssName: string) {
    this.ws = ws; 
    this.ssName = ssName;
    this.errors = new Errors();
    this.makeEmptySS();
    this.addListeners();
    //TODO: initialize added instance variables
  }

  static async make(ws: SpreadsheetWs, ssName: string) {
    const ss = new Spreadsheet(ws, ssName);
    await ss.load();
    return ss;
  }

  /** add listeners for different events on table elements */
  private addListeners() {
    //TODO: add listeners for #clear and .cell
    
    const clearButton = document.querySelector('#clear')!;
    clearButton.addEventListener('click', this.clearSpreadsheet);

    const cellElements = document.querySelectorAll('.cell');
    cellElements.forEach((cell) => {
      cell.addEventListener('focusin', this.focusCell);
      cell.addEventListener('focusout', this.blurCell);
      cell.addEventListener('copy', this.copyCell);
      cell.addEventListener('paste', this.pasteCell);
    });

  }

  /** listener for a click event on #clear button */
  private readonly clearSpreadsheet = async (ev: Event) => {

    const result = await this.ws.clear(this.ssName);
    if (result.isOk) { 
      // If successful, clear the cells in the DOM
      const cellElements = document.querySelectorAll('.cell');
      cellElements.forEach((cell) => {
        cell.textContent = '';
        cell.removeAttribute('data-expr');
        cell.removeAttribute('data-value');
      });
    } else {
      // Handle the error result by reporting the errors 
      this.errors.display(result.errors);
    }
  };
 
    
  /** listener for a focus event on a spreadsheet data cell */
  private readonly focusCell = (ev: Event) => {
    const targetCell = ev.target as HTMLElement;

    // Set the cell's text content to the value of its data-expr attribute
    const cellExpr = targetCell.getAttribute('data-expr');
    targetCell.textContent = cellExpr || '';
  
    // Update the focusedCellId and isFocused
    this.focusedCellId = targetCell.id;
    this.isFocused = true;
  

    this.currentlyFocusedCellId = targetCell.id;
  
    // Add the focusout event listener to handle focusout behavior
      document.addEventListener('focusout', this.focusoutCell);
      this.errors.clear();
  };

  private readonly focusoutCell = (ev: Event) => {
    // Get the target element of the focusout event
    const targetElement = ev.target as HTMLElement;
  
    if (targetElement.classList.contains('cell')) {
      // Set the cell's text content to the value of its data-value attribute
      const cellValue = targetElement.getAttribute('data-value');
      targetElement.textContent = cellValue || '';
      // Update the focusedCellId and isFocused flags
      this.focusedCellId = null;
      this.isFocused = false;
    }
  
    // Remove the focusout event listener
    document.removeEventListener('focusout', this.focusoutCell);
  };
  

/** listener for a blur event on a spreadsheet data cell */

  private readonly blurCell = async (ev: Event) => {
      
const targetCell = ev.target as HTMLElement;
const cellTextContent = targetCell.textContent?.trim() ?? ''; // Provide a default value when cellTextContent is undefined
const cellId = targetCell.id;

if (this.isFocused) {
  if (cellTextContent !== '') {
    // If the trimmed content is not the same as the original expression, call the web service to evaluate the cell expression

    
    const result = await this.ws.evaluate(this.ssName, cellId, cellTextContent);

    if (result.isOk) {
      const evaluatedValue = result.val;

      // Update the cell's data-expr and data-value attributes with the result
      targetCell.dataset.expr = cellTextContent;

      // Set the cell's evaluated value based on its ID
      for (const [cellId, value] of Object.entries(evaluatedValue)) {
        const cellElement = document.getElementById(cellId);
        if (cellElement && cellElement.id !== this.currentlyFocusedCellId) {
          cellElement.textContent = value.toString();
          cellElement.dataset.value = value.toString();
        }
      }
    } else {
       // Handle the error result 
       this.errors.display(result.errors);
    }
    
 
  } else {
    // If the trimmed content is empty, call the web service to remove the cell
    const result = await this.ws.remove(this.ssName, cellId);
    if (result.isOk) {
      // If successful, clear the cell in the DOM
      targetCell.textContent = '';
      targetCell.removeAttribute('data-expr');
      targetCell.removeAttribute('data-value');

      const evaluatedValue = result.val;
      
      for (const [cellId, value] of Object.entries(evaluatedValue)) {
        const cellElement = document.getElementById(cellId);
        if (cellElement && cellElement.id !== this.currentlyFocusedCellId) {
          cellElement.textContent = value.toString();
          cellElement.dataset.value = value.toString();
        }
      }
    } else {
      // Handle the error result by reporting the errors using the errors instance variable
      this.errors.display(result.errors);
    }
  }

  // Clear the focusedCellId after updating the cell
  this.focusedCellId = null;

  // Clear the currently focused cell id to prevent updates to the focused cell
  this.currentlyFocusedCellId = null;
}

  };
  
 
  
  /** listener for a copy event on a spreadsheet data cell */
  private readonly copyCell = (ev: Event) => {
      // Remember the cellId of the source cell in an instance variable
    const sourceCell = ev.target as HTMLElement;
    this.copySourceCellId = sourceCell.id;

    // Add the 'is-copy-source' class to the source cell element
    sourceCell.classList.add('is-copy-source');
    
  };

  /** listener for a paste event on a spreadsheet data cell */
  private readonly pasteCell = async (ev: Event) => {
    ev.preventDefault();
    if (!this.copySourceCellId) {
      // No source cell to copy from, return early
      return;
    }
  
    const destinationCell = ev.target as HTMLElement;
    const destinationCellId = destinationCell.id;
  
    // Copy the formula from the source cell to the destination cell and adjust relative cell references
    const copyResult = await this.ws.copy(this.ssName, destinationCellId, this.copySourceCellId);
  
    if (copyResult.isOk) {
      const updatedCells = copyResult.val;
      const res1 = await this.ws.query(this.ssName, destinationCellId);
      if(res1.isOk){
        const cell = document.getElementById(destinationCellId);
        if(cell){
          cell.textContent = res1.val.expr;
          cell.dataset.expr = res1.val.expr;
        }
      }
      // Update the destination cell and other affected cells in the DOM with the updated values
      for (const [cellId, cellValue] of Object.entries(updatedCells)) {
        const cellElement = document.getElementById(cellId);
        if (cellElement) {
          if(this.currentlyFocusedCellId && cellElement.id != this.currentlyFocusedCellId){
          cellElement.dataset.value = cellValue.toString();
          //cellElement.setAttribute('data-value', cellValue.toString());
          cellElement.textContent = cellValue.toString();
          }
        }
      }

      
       // Check if there is a source cell (cell previously copied)
    if (this.copySourceCellId) {
      // Remove the 'is-copy-source' class from the source cell element
      const sourceCell = document.getElementById(this.copySourceCellId);
      if (sourceCell) {
        sourceCell.classList.remove('is-copy-source');
      }

      // Clear the copy source cell information after pasting is done
      this.copySourceCellId = null;
      this.copySourceExpr = null;
    }

       
      } else {
        // Handle the error result by reporting the errors using the errors instance variable
        this.errors.display(copyResult.errors);
      }
   
  };

  /** Replace entire spreadsheet with that from the web services.
   *  Specifically, for each active cell set its data-value and 
   *  data-expr attributes to the corresponding values returned
   *  by the web service and set its text content to the cell value.
   */
  /** load initial spreadsheet data into DOM */

  private async load() {
    this.errors.clear();
  
    // Query the web service to get the active cells' data
    const response = await this.ws.dumpWithValues(this.ssName);
  
    if (response.isOk) {
      // Get the cell data from the response result
      const cellData = response.val;
  
      // Create a data structure to store cell data
      const cellDataMap: { [cellId: string]: { expr: string; value: number } } = {};
  
      // Loop through each cell's data and update the data structure accordingly
      for (const [cellId, expr, value] of cellData) {
        // Update the data structure with the cell data
        cellDataMap[cellId] = { expr, value };
      }
  
      // Update the DOM based on the data structure
      for (const cellId of Object.keys(cellDataMap)) {
        const cellElement = document.getElementById(cellId);
  
        if (cellElement) {
          // Update the data-expr and data-value attributes if the cell is not currently focused
          if (cellId !== this.currentlyFocusedCellId) {
            // Update the data-expr and data-value attributes
            cellElement.dataset.expr = cellDataMap[cellId].expr;
            cellElement.dataset.value = cellDataMap[cellId].value.toString();
  
            // Set the text content to the cell value
            cellElement.textContent = cellDataMap[cellId].value.toString();
          }
        }
      }
    } else {
      // Handle the error result by reporting the errors using the errors instance variable
      this.errors.display(response.errors);
    }
  }
  
  
private makeEmptySS() {
  const ssDiv = document.querySelector('#ss')!;
  ssDiv.innerHTML = '';
  const ssTable = makeElement('table');
  const header = makeElement('tr');
  const clearCell = makeElement('td');
  const clear = makeElement('button', { id: 'clear', type: 'button' }, 'Clear');
  clearCell.append(clear);
  header.append(clearCell);
  const A = 'A'.charCodeAt(0);
  for (let i = 0; i < N_COLS; i++) {
    header.append(makeElement('th', {}, String.fromCharCode(A + i)));
  }
  ssTable.append(header);
  for (let i = 0; i < N_ROWS; i++) {
    const row = makeElement('tr');
    row.append(makeElement('th', {}, (i + 1).toString()));
    const a = 'a'.charCodeAt(0);
    for (let j = 0; j < N_COLS; j++) {
      const colId = String.fromCharCode(a + j);
      const id = colId + (i + 1);
      const cell = makeElement('td', { id, class: 'cell', contenteditable: 'true', tabindex: '0' });
      row.append(cell);
    }
    ssTable.append(row);
  }
  ssDiv.append(ssTable);
}


}
