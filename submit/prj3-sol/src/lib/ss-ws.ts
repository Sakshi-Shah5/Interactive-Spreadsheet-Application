import cors from 'cors';
import Express from 'express';
import bodyparser from 'body-parser';
import assert from 'assert';
import STATUS from 'http-status';
import { Result, okResult, errResult, Err, ErrResult } from 'cs544-js-utils';
import { SpreadsheetServices as SSServices } from 'cs544-prj2-sol';
import { SelfLink, SuccessEnvelope, ErrorEnvelope } from './response-envelopes.js';

export type App = Express.Application;

export function makeApp(ssServices: SSServices, base = '/api'): App {
  const app = Express();
  app.locals.ssServices = ssServices;
  app.locals.base = base;
  setupRoutes(app);
  return app;
}


/******************************** Routing ******************************/

const CORS_OPTIONS = {
  origin: '*',
  methods: 'GET,HEAD,PUT,PATCH,POST,DELETE',
  preflightContinue: false,
  optionsSuccessStatus: 204,
  exposedHeaders: 'Location',
};

function setupRoutes(app: Express.Application) {
  const base = app.locals.base;
  app.use(cors(CORS_OPTIONS));  //will be explained towards end of course
  app.use(Express.json());  //all request bodies parsed as JSON.

   // Dummy route for debugging
   app.get(`${base}/debug`, (req, res) => {
    res.json({ message: 'Debug route reached successfully' });
  });

  //routes for individual cells
  app.put(`${base}/:ssName`, makePutCellHandler(app));
  app.get(`${base}/:ssName/:cellId`, makeGetCellHandler(app));
  app.patch(`${base}/:ssName/:cellId`, makePatchCellHandler(app));
  app.delete(`${base}/:ssName/:cellId`, makeDeleteCellHandler(app));

  //routes for entire spreadsheets
  app.get(`${base}/:ssName`, makeGetSpreadsheetHandler(app));
  app.delete(`${base}/:ssName`, makeClearSpreadsheetHandler(app));

  //generic handlers: must be last
  app.use(make404Handler(app));
  app.use(makeErrorsHandler(app));
}

/* A handler can be created by calling a function typically structured as
   follows:

   function makeOPHandler(app: Express.Application) {
     return async function(req: Express.Request, res: Express.Response) {
       try {
         const { ROUTE_PARAM1, ... } = req.params; //if needed
         const { QUERY_PARAM1, ... } = req.query;  //if needed
   VALIDATE_IF_NECESSARY();
   const SOME_RESULT = await app.locals.ssServices.OP(...);
   if (!SOME_RESULT.isOk) throw SOME_RESULT;
         res.json(selfResult(req, SOME_RESULT.val));
       }
       catch(err) {
         const mapped = mapResultErrors(err);
         res.status(mapped.status).json(mapped);
       }
     };
   }
*/

/****************** Handlers for Spreadsheet Cells *********************/


// Handler for PUT requests to update a cell in a spreadsheet
function makePutCellHandler(app: Express.Application) {
  return async function (req: Express.Request, res: Express.Response) {
    try {
      const { ssName, cellId } = req.params;
      const ssServices = app.locals.ssServices;


      // Evaluate the expression and store the result
      const evalResult = await ssServices.load(ssName, req.body);

      if (!evalResult.isOk) throw evalResult;

      res.json(selfResult(req, evalResult.val));
    } catch (err) {
      const mapped = mapResultErrors(err);
      res.status(mapped.status).json(mapped);
    }
  };
}

// Handler for GET requests to retrieve the value of a cell in a spreadsheet
function makeGetCellHandler(app: Express.Application) {
  return async function (req: Express.Request, res: Express.Response) {
    try {
      const { ssName, cellId } = req.params;
      const ssServices = app.locals.ssServices;

      // Call the query() method to retrieve the cell value
      const result = await ssServices.query(ssName, cellId);
      if (!result.isOk) throw result;

      res.json(selfResult(req, result.val));
    } catch (err) {
      const mapped = mapResultErrors(err);
      res.status(mapped.status).json(mapped);
    }
  };
}

// Handler for PATCH requests to update a cell using an expression or copying from another cell
function makePatchCellHandler(app: Express.Application) {
  return async function (req: Express.Request, res: Express.Response) {
    try {
      const { ssName, cellId } = req.params;
      const { expr, srcCellId } = req.query;
      const ssServices = app.locals.ssServices;

      if (!expr && !srcCellId) {
        throw new Error('Missing query parameter. Please provide either "expr" or "srcCellId".');
      }

      //if both "expr" and "srcCellId" are provided in the query parameters
      if (expr && srcCellId) {
        throw new Error('Invalid query parameters. Please provide only one of "expr" or "srcCellId".');
      }

      let result;

      if (expr) {
        result = await ssServices.evaluate(ssName, cellId, expr); // If "expr" is provided, call the evaluate() method to update the cell with the expression
      } else if (srcCellId) {
        result = await ssServices.copy(ssName, cellId, srcCellId); // If "srcCellId" is provided, call the copy() method to update the cell by copying from another cell
      }

      if (!result.isOk) throw result;

      res.json(selfResult(req, result.val));
    } catch (err) {
      const mapped = mapResultErrors(err);
      res.status(mapped.status).json(mapped);
    }
  };
}

// Handler for DELETE requests to remove a cell from a spreadsheet
function makeDeleteCellHandler(app: Express.Application) {
  return async function (req: Express.Request, res: Express.Response) {
    try {
      const { ssName, cellId } = req.params;
      const ssServices = app.locals.ssServices;

      // Remove the cell from the spreadsheet
      const deleteResult = await ssServices.remove(ssName, cellId);
      if (!deleteResult.isOk) throw deleteResult;

      res.json(selfResult(req, deleteResult.val)); // Return the deleted cells

    } catch (err) {
      const mapped = mapResultErrors(err);
      res.status(mapped.status).json(mapped);
    }
  };
}



/**************** Handlers for Complete Spreadsheets *******************/


// Handler for GET requests to retrieve the entire spreadsheet
function makeGetSpreadsheetHandler(app: Express.Application) {
  return async function (req: Express.Request, res: Express.Response) {
    try {
      const { ssName } = req.params;
      const ssServices = app.locals.ssServices;

      const dumpResult = await ssServices.dump(ssName, true); //call the dump() method to retrieve the entire spreadsheet data

      if (!dumpResult.isOk) throw dumpResult;

      const cells = dumpResult.val;

      res.json(selfResult(req, cells, STATUS.OK));
    } catch (err) {
      const mapped = mapResultErrors(err);
      res.status(mapped.status).json(mapped);
    }
  };
}

// Handler for DELETE requests to clear the entire spreadsheet
function makeClearSpreadsheetHandler(app: Express.Application) {
  return async function (req: Express.Request, res: Express.Response) {
    try {
      const { ssName } = req.params;
      const ssServices = app.locals.ssServices;

      // Clear the spreadsheet
      const clearResult = await ssServices.clear(ssName);
      if (!clearResult.isOk) throw clearResult;

      const deletedCells = clearResult.val;

      res.json(selfResult(req, deletedCells, STATUS.OK));
    } catch (err) {
      const mapped = mapResultErrors(err);
      res.status(mapped.status).json(mapped);
    }
  };
}


/*************************** Generic Handlers **************************/

/** Default handler for when there is no route for a particular method
 *  and path.
  */
function make404Handler(app: Express.Application) {
  return async function (req: Express.Request, res: Express.Response) {
    const message = `${req.method} not supported for ${req.originalUrl}`;
    const result = {
      status: STATUS.NOT_FOUND,
      errors: [{ options: { code: 'NOT_FOUND' }, message, },],
    };
    res.status(404).json(result);
  };
}


/** Ensures a server error results in nice JSON sent back to client
 *  with details logged on console.
 */
function makeErrorsHandler(app: Express.Application) {
  return async function (err: Error, req: Express.Request, res: Express.Response,
    next: Express.NextFunction) {
    const message = err.message ?? err.toString();
    const result = {
      status: STATUS.INTERNAL_SERVER_ERROR,
      errors: [{ options: { code: 'INTERNAL' }, message }],
    };
    res.status(STATUS.INTERNAL_SERVER_ERROR as number).json(result);
    console.error(result.errors);
  };
}


/************************* HATEOAS Utilities ***************************/

/** Return original URL for req */
function requestUrl(req: Express.Request) {
  return `${req.protocol}://${req.get('host')}${req.originalUrl}`;
}

function selfHref(req: Express.Request, id: string = '') {
  const url = new URL(requestUrl(req));
  return url.pathname + (id ? `/${id}` : url.search);
}

function selfResult<T>(req: Express.Request, result: T,
  status: number = STATUS.OK)
  : SuccessEnvelope<T> {
  return {
    isOk: true,
    status,
    links: { self: { href: selfHref(req), method: req.method } },
    result,
  };
}



/*************************** Mapping Errors ****************************/

//map from domain errors to HTTP status codes.  If not mentioned in
//this map, an unknown error will have HTTP status BAD_REQUEST.
const ERROR_MAP: { [code: string]: number } = {
  EXISTS: STATUS.CONFLICT,
  NOT_FOUND: STATUS.NOT_FOUND,
  BAD_REQ: STATUS.BAD_REQUEST,
  AUTH: STATUS.UNAUTHORIZED,
  DB: STATUS.INTERNAL_SERVER_ERROR,
  INTERNAL: STATUS.INTERNAL_SERVER_ERROR,
  BAD_REQ_PATCH_NO_PARAMS: STATUS.BAD_REQUEST,
  BAD_REQ_PATCH_BOTH_PARAMS: STATUS.BAD_REQUEST,
};

/** Return first status corresponding to first options.code in
 *  errors, but SERVER_ERROR dominates other statuses.  Returns
 *  BAD_REQUEST if no code found.
 */
function getHttpStatus(errors: Err[]): number {
  let status: number = 0;
  for (const err of errors) {
    if (err instanceof Err) {
      const code = err?.options?.code;
      const errStatus = (code !== undefined) ? ERROR_MAP[code] : -1;
      if (errStatus > 0 && status === 0) status = errStatus;
      if (errStatus === STATUS.INTERNAL_SERVER_ERROR) status = errStatus;
    }
  }
  return status !== 0 ? status : STATUS.BAD_REQUEST;
}

/** Map domain/internal errors into suitable HTTP errors.  Return'd
 *  object will have a "status" property corresponding to HTTP status
 *  code.
 */
function mapResultErrors(err: Error | ErrResult): ErrorEnvelope {
  const errors = (err instanceof Error)
    ? [new Err(err.message ?? err.toString(), { code: 'BAD_REQ' }),]
    : err.errors;
  const status = getHttpStatus(errors);
  if (status === STATUS.SERVER_ERROR) console.error(errors);
  return { isOk: false, status, errors, };
}

