import { Result, okResult, errResult } from 'cs544-js-utils';
import * as mongo from 'mongodb';

export async function makeSpreadsheetDao(mongodbUrl: string, ssName: string): Promise<Result<SpreadsheetDao>> {
  // try {
    if (!mongodbUrl.startsWith('mongodb://')) {
      return errResult('Invalid database URL', 'DB');
    }
  
  return SpreadsheetDao.make(mongodbUrl,ssName);

}



//represents a Data Access Object (DAO) for working with a spreadsheet. It has private properties for the MongoDB URL, spreadsheet name, MongoDB client, database, and collection
export class SpreadsheetDao {
  constructor(
    private mongodbUrl: string,
    private ssName: string,
    private client: mongo.MongoClient,
    private db: mongo.Db,
    private collection: mongo.Collection<mongo.Document>,
  ) { }

  //factory method

  //connects to the MongoDB database using the provided URL, creates a collection named 'spreadsheets', and creates an index on the _id field. If successful, it returns a Result with the created SpreadsheetDao object. If an error occurs, it returns an error Result with a message and an error code
  static async make(dbUrl: string, ssName: string): Promise<Result<SpreadsheetDao>> {
    try {
      const client = await mongo.MongoClient.connect(dbUrl);
      const db = client.db(ssName);
      const collection = db.collection('spreadsheets');
      // Create indexes if they don't exist
      await collection.createIndex({ _id: 1 }); // Index on _id field

      return okResult(new SpreadsheetDao(dbUrl, ssName, client, db, collection));
    } catch (error) {
      return errResult(error.message, 'DB');
    }
  }

  /** Release all resources held by persistent spreadsheet. *  Specifically, close any database connections. */
  async close(): Promise<Result<undefined>> {
    try {
      await this.client.close();
      return okResult(undefined);
    } catch (e) {
      return errResult(e.message, 'DB');
    }
  }

  /** return name of this spreadsheet */
  getSpreadsheetName(): string {
    return this.ssName;  //returns the name of the spreadsheet associated with the SpreadsheetDao object
  }

  /** Set cell with id cellId to string expr. */
  async setCellExpr(cellId: string, expr: string): Promise<Result<undefined>> {

    //the setCellExpr function updates the expression of a cell in the spreadsheet using MongoDB's updateOne method, and it handles different scenarios such as successful updates, circular references, and error handling.

    try {

      // console.log("visitedCells : ",visitedCells);
      // console.log("cellId: ", cellId);

     
      // visitedCells.add(cellId);
  
      const [spreadsheetName, formattedCellId] = cellId.split(':'); //the cellId is split into spreadsheetName and formattedCellId using the ':' separator. This allows the function to identify the specific cell within the spreadsheet

      const filter: mongo.Filter<mongo.Document> = {
        spreadsheetName: spreadsheetName,
        cellId: formattedCellId,
      }; //A filter object is created using the spreadsheetName and formattedCellId. This filter is used to identify the document in the MongoDB collection that corresponds to the cell

      // console.log("filter : ",filter);
  
      const updateResult = await this.collection.updateOne(
        filter,
        { $set: { expr: expr } },
        { upsert: true }
      ); //The update object uses the $set operator to set the expr field of the document to the provided expr value. The upsert option is set to true, which means that if a document matching the filter doesn't exist, it will be created.
      //The updateResult is awaited, which is a result object that contains information about the update operation, such as the number of documents matched, modified, or upserted.

  
      if (updateResult.upsertedCount === 1 || updateResult.modifiedCount === 1) {
        //If the upsertedCount or modifiedCount in the updateResult is equal to 1, it means that the cell expression was successfully set or updated. In this case, an okResult is returned with undefined as the value, indicating a successful operation.
        return okResult(undefined);
      } else {
        return errResult(`Circular reference involving ${cellId}`, 'DB'); //If neither upsertedCount nor modifiedCount is equal to 1, it implies that the update operation didn't modify any document or resulted in a circular reference. In this case, an errResult is returned with an error message indicating the circular reference involving the cellId.
      }
    } catch (error) {
      return errResult({ code: 'DB', message: error.message });
    }
  }
  

  /** Return expr for cell cellId; return '' for an empty/unknown cell.  */

  async query(cellId: string): Promise<Result<string>> {

    // It queries the database to retrieve the expression of the cell identified by cellId. It returns a Promise that resolves to a Result containing the expression as a string if the cell is found. If the cell is not found, it returns an empty string. If an error occurs during the query, it returns an error Result with a message and an error code.

    try {
      const [spreadsheetName, formattedCellId] = cellId.split(':');

      const doc = await this.collection.findOne({ spreadsheetName: spreadsheetName, cellId: formattedCellId });
      // console.log("doc inside query method:" , doc);

      if (doc) {
        return okResult(doc.expr);
      } else {
        return okResult(''); // Return empty string when cell is not found
      }
    } catch (e) {
      return errResult(e.message, 'DB');
    }
  }

  async clear(): Promise<Result<undefined>> {
    try {
      await this.collection.deleteMany({});      // Delete all documents in the collection
      return okResult(undefined);
    } catch (e) {
      return errResult(e.message, 'DB');
    }
  }

  /** Remove all info for cellId from this spreadsheet. */
  async remove(cellId: string): Promise<Result<undefined>> {
    try {
      const [spreadsheetName, formattedCellId] = cellId.split(':');
      await this.collection.deleteOne({ spreadsheetName, cellId: formattedCellId });
      return okResult(undefined);
    } catch (e) {
      return errResult(e.message, 'DB');
    }
  }


  /** Return array of [cellId, expr] pairs for all cells in this spreadsheet */
  async getData(): Promise<Result<[string, string][]>> {
    try {
      const docs = await this.collection.find().toArray();  //find method is called on the MongoDB collection without any filter, which retrieves all documents from the collection. The toArray method is then called on the resulting cursor to convert the documents into an array.


      // console.log("docs inside getData(): ", docs);
  
      const data: [string, string][] = docs.map((doc) => {

        //The data variable is initialized as an empty array. This array will store the cell data retrieved from the documents.
        //map method is called on the docs array to iterate over each document and transform it into a cell data tuple of the form [cellId, expr].

        const cellId = doc.spreadsheetName ?? ''; // Use empty string if spreadsheetName is null or undefined
        const expr = doc.expr ?? '';
        return [cellId, expr];
      });
      
      
  
      // console.log("data object inside getData():", data);
  
      return okResult(data);
    } catch (e) {
      return errResult(e.message, 'DB');
    }
  }
  


}



