import { default as parse, CellRef, Ast } from './expr-parser.js';
import { Result, okResult, errResult, ErrResult } from 'cs544-js-utils';

class CellInfo {
  id: string;
  expr: string;
  ast: Ast | null;
  value: number;
  dependents: Set<string>;

  constructor(id: string, expr: string) {
    this.id = id;
    this.expr = expr;
    this.ast = null;
    this.value = 0;
    this.dependents = new Set();
  }
}

export default async function makeSpreadsheet(name: string): Promise<Result<Spreadsheet>> {
  return okResult(new Spreadsheet(name));
}

type Updates = { [cellId: string]: number };

export class Spreadsheet {
  readonly name: string;                      // Name of the spreadsheet
  cells: { [cellId: string]: CellInfo };     // Object mapping cell IDs to CellInfo objects

  constructor(name: string) {
    this.name = name;
    this.cells = {};
  }


  // evaluates an expression for a specific cell in the spreadsheet. It takes the cellId and the expr as parameters. It first parses the expression using the parse function .
  // If the parsing is successful, it updates the cell information in the spreadsheet with the new expression and evaluates the expression to compute its result. It also updates the dependencies of the cell and recursively updates all dependent cells. The function returns a Promise containing the updates made to the cells.

  async eval(cellId: string, expr: string): Promise<Result<Updates>> {
    const oldCells = JSON.parse(JSON.stringify(this.cells));
    try {
      const astResult = parse(expr, cellId) as Result<Ast>;
      //const astResult = parse(expr) as Result<Ast>;
      // console.log(JSON.stringify(astResult, null, 2));
      // console.log("start " , cellId,expr);


      // Parse the expression and check for syntax errors

      if (astResult.isOk) {
        const ast = astResult.val;
        const baseCellRef = CellRef.parse(cellId);
        //  console.log("baseCellRef" + baseCellRef);

        if (baseCellRef.isOk) {
          const cellInfo = this.cells[cellId] || new CellInfo(cellId, expr);
          const previousAst = cellInfo.ast;
          cellInfo.ast = ast;


          // Update dependencies and evaluate the current formula
          const referencedCellIds = this.extractReferencedCellIds(ast, baseCellRef.val);
          // console.log("referencedCellIds" , referencedCellIds);

          // Remove cellId from dependents of cells referenced by the previous formula
          if (previousAst) {
            const previousReferencedCellIds = this.extractReferencedCellIds(previousAst, baseCellRef.val);
            // console.log("previousReferencedCellIds",previousReferencedCellIds);

            previousReferencedCellIds.forEach(previousReferencedCellId => {
              const previousReferencedCellInfo = this.cells[previousReferencedCellId];
              // console.log("previousReferencedCellInfo", previousReferencedCellInfo);

              if (previousReferencedCellInfo) {
                previousReferencedCellInfo.dependents.delete(cellId);
              }
            });
          }

          //Add cellId to dependents of cells referenced by the current formula
          let circularRefFound = referencedCellIds.includes(cellId);

          //console.log("circularRefFound",circularRefFound);

          for (const referencedCellId of referencedCellIds) {
            if (referencedCellId === cellId) {
              circularRefFound = true;
              break; // Exit the loop if circular reference is found
            }

            let referencedCellInfo = this.cells[referencedCellId];
            // console.log("referencedCellId", referencedCellId, cellId, this.cells);
            if (referencedCellInfo) {
              if (!(referencedCellInfo.dependents instanceof Set)) {
                referencedCellInfo.dependents = new Set();
              }
              referencedCellInfo.dependents.add(cellId);
            } else {
              referencedCellInfo = new CellInfo(referencedCellId, '0');
              referencedCellInfo.dependents = new Set(); // Initialize dependents as an empty Set
              this.cells = { ...this.cells, [referencedCellId]: referencedCellInfo };
              referencedCellInfo.dependents.add(cellId);
            }
          }


          if (circularRefFound) {
            //console.log("Here");
            throw errResult('CIRCULAR_REF', 'expected direct circular reference');
          }

          //  console.log("cellInfos", this.cells);
          const result = this.evalAst(ast, baseCellRef.val, cellId);
          cellInfo.value = result;
          // console.log("result",result);

          this.cells = { ...this.cells, [cellId]: cellInfo };
          const updates: Updates = { [cellId]: result };

          // Update all dependent cells recursively
          const updatedCells = this.updateDependentCells(cellId);
          Object.assign(updates, updatedCells);

          return okResult(updates);
        } else {
          return errResult({ code: 'SYNTAX', message: 'Invalid numeric expression' });
        }
      } else {
        return errResult('syntax error', 'SYNTAX');
      }
    } catch (error) {
      this.cells = JSON.parse(JSON.stringify(oldCells));    //In case of an error, the code restores the original state by assigning the oldCells back to this.cells

      return errResult('CIRCULAR_REF', 'CIRCULAR_REF');
    }
  }

  // This function updates all dependent cells recursively for a given cell.
  updateDependentCells(cellId: string, processedCells: Set<string> = new Set(), path: string[] = []): Updates {
    const cellInfo = this.cells[cellId];
    //console.log("processedCells ",processedCells);
    if (!cellInfo || !cellInfo.ast) {
      //console.log("path includes in");

      return {};
    }

    if (processedCells.has(cellId)) {
      return {}; // Avoid infinite recursion if cell is already processed
    }

    const dependentIds = Array.from(cellInfo.dependents);
    const updates: Updates = {};

    for (const dependentId of dependentIds) {
      const dependentCellInfo = this.cells[dependentId];
      if (dependentCellInfo && dependentCellInfo.ast) {

        //For each dependent cell, it checks if there is a circular reference by checking if the dependent cell ID is already present in the path array. If a circular reference is detected, it throws an error. 

        if (path.includes(dependentId)) {
          // Circular reference detected
          throw errResult('CIRCULAR_REF', 'CIRCULAR_REF');
        }
        //By throwing an error, the code prevents partial updates and ensures that the spreadsheet remains unchanged in case of circular dependencies.


        
        processedCells.add(cellId);
        const dependentCellRefResult = CellRef.parse(dependentId);
        if (dependentCellRefResult.isOk) {

          const dependentValue = this.evalAst(dependentCellInfo.ast, dependentCellRefResult.val, dependentId);
          dependentCellInfo.value = dependentValue;
          updates[dependentId] = dependentValue;

          //Otherwise, it recursively updates the dependent cell by evaluating its AST and updating its value. It returns an object containing the updates made to the dependent cells.

          const recursiveUpdates = this.updateDependentCells(dependentId, processedCells, [...path, cellId]);
          Object.assign(updates, recursiveUpdates);
        } else {
          throw new Error(`Invalid cell reference: ${dependentId}`);
        }
      }
    }

    return updates;
  }

  // Extract cell IDs referenced in an AST
  extractReferencedCellIds(ast: Ast, baseCellRef: CellRef): string[] {
    //It traverses the AST and for each reference node (ref), it parses the cell reference using the baseCellRef and adds the cell ID to the referencedCellIds array. It recursively traverses the AST for application nodes (app). Finally, it returns the referencedCellIds array.

    const referencedCellIds: string[] = [];

    function traverseAst(node: Ast) {
      switch (node.kind) {
        case 'ref':
          const cellRef = CellRef.parse(baseCellRef.toText(node.value));
          if (cellRef.isOk) {
            const cellId = cellRef.val.toText();
            referencedCellIds.push(cellId);
          }
          break;
        case 'app':
          node.kids.forEach(kid => traverseAst(kid));
          break;
        default:
          break;
      }
    }

    traverseAst(ast);
    return referencedCellIds;
  }

  // Evaluate an AST and compute the result
  evalAst(node: Ast, baseCellRef: CellRef, cellId: string): number {
    switch (node.kind) {
      case 'num':
        return node.value;

      case 'ref':
        const cellRef = CellRef.parse(baseCellRef.toText(node.value));
        if (cellRef.isOk) {
          const refCellId = cellRef.val.toText();
          const cellInfo = this.cells[refCellId];
          return cellInfo ? cellInfo.value : 0;
        }
        return 0;

      case 'app':
        const kidValues = node.kids.map(kid => this.evalAst(kid, baseCellRef, cellId));
        const fn = FNS[node.fn];
        if (typeof fn === 'function') {
          return fn.apply(null, kidValues);
        } else {
          throw new Error(`Unsupported function: ${node.fn}`);
        }

      default:
        throw new Error('Unsupported AST node kind');
    }
  }
}

const FNS = {
  '+': (a: number, b: number): number => a + b,
  '-': (a: number, b?: number): number => (b === undefined ? -a : a - b),
  '*': (a: number, b: number): number => a * b,
  '/': (a: number, b: number): number => a / b,
  min: (a: number, b: number): number => Math.min(a, b),
  max: (a: number, b: number): number => Math.max(a, b),
};
