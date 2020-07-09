import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import request from '../../../../request';
import {
  CommandOption,
  CommandValidate, CommandError, CommandTypes
} from '../../../../Command';
import SpoCommand from '../../../base/SpoCommand';
import Utils from '../../../../Utils';
import { FolderExtensions } from '../../FolderExtensions';
import * as path from 'path';
import { Transform } from 'stream';
const vorpal: Vorpal = require('../../../../vorpal-init');
import * as csv from '@fast-csv/parse';
import { v4 } from 'uuid';
import { createReadStream, ReadStream } from 'fs';
import requestPromise = require('request-promise-native');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  batchSize?: number;
  contentType?: string;
  folder?: string;
  listId?: string;
  listTitle?: string;
  path: string;
  webUrl: string;
}

interface FieldNames {
  value: { InternalName: string }[];
}

interface ContentTypes {
  value: {
    Id: {
      StringValue: string
    },
    Name: string
  }[];
}

interface GetWebResponse {
  Url: string
}

interface GetRootFolderResponse {
  ServerRelativeUrl: string
}

class SpoListItemBatchAddCommand extends SpoCommand {
  public get name(): string {
    return commands.LISTITEM_BATCH_ADD;
  }

  public get description(): string {
    return 'Creates list items from the specified .csv file';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.batchSize = typeof args.options.batchSize !== 'undefined';
    telemetryProps.contentType = typeof args.options.contentType !== 'undefined';
    telemetryProps.folder = typeof args.options.folder !== 'undefined';
    telemetryProps.listId = typeof args.options.listId !== 'undefined';
    telemetryProps.listTitle = typeof args.options.listTitle !== 'undefined';
    return telemetryProps;
  }

  public types(): CommandTypes | undefined {
    return {
      string: ['c', 'contentType']
    };
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let lineNumber: number = 0;
    let contentTypeName: string | null = null;

    let maxBytesInBatch: number = 1000000; // max is  1048576
    let rowsInBatch: number = 0;
    let batchCounter: number = 0;
    let recordsToAdd: string = "";
    let csvHeaders: string[];

    const fullPath: string = path.resolve(args.options.path);
    const fileName: string = Utils.getSafeFileName(path.basename(fullPath));
    const listIdArgument: string = args.options.listId || '';
    const listTitleArgument: string = args.options.listTitle || '';
    const batchSize: number = args.options.batchSize || 10;

    let listRestUrl = args.options.listId ?
      `${args.options.webUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
      : `${args.options.webUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`;

    SpoListItemBatchAddCommand
      .validateContentType(args.options.contentType, listRestUrl, this.verbose, cmd)
      .catch((error: string): void => {
        cb(new CommandError(error));
      })
      .then((): Promise<string> => SpoListItemBatchAddCommand.getCaseSensitiveWebUrl(args.options.webUrl))
      .then((caseCorrectedWebUrl: string): Promise<string> => {
        listRestUrl = args.options.listId ?
          `${caseCorrectedWebUrl}/_api/web/lists(guid'${encodeURIComponent(listIdArgument)}')`
          : `${caseCorrectedWebUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listTitleArgument)}')`;

        return SpoListItemBatchAddCommand.getFolderUrl(args.options.folder, listRestUrl as string, caseCorrectedWebUrl as string, this.verbose, this.debug, cmd);
      })
      .then((folderServerRelativeUrl: string): any => {
        if (this.verbose) {
          cmd.log(`Creating items in list ${folderServerRelativeUrl}`);
        }

        //start the batch -- each batch will get assigned its own id
        let changeSetId: string = v4();
        const endpoint: string = `${listRestUrl}/AddValidateUpdateItemUsingPath()`;

        // get the csv  file passed in from the cmd line
        const fileStream: ReadStream = createReadStream(fileName);
        const csvStream: any = csv.parseStream(fileStream, { headers: false });
        const verboseMode: boolean = this.verbose;

        csvStream
          .pipe(new Transform({ // https://github.com/C2FO/fast-csv/issues/328 Need to transform if we are batching async
            objectMode: true,
            write(row: any, encoding: string, callback: (error?: (Error | null)) => void): void {
              if (lineNumber === 0) {
                // Process csv Headers (fast csv headers don't work if using transform)
                csvHeaders = row;

                // fetch the valid field names from the list. If you pass a bad field name to AddValidateUpdateItemUsingPath it returns xml not JSON
                const fetchFieldsRequest: requestPromise.OptionsWithUrl = {
                  url: `${listRestUrl}/fields?$select=InternalName&$filter=ReadOnlyField eq false`,
                  json: true,
                  headers: {
                    'Accept': `application/json;odata=nometadata`
                  },
                }
                request
                  .get<FieldNames>(fetchFieldsRequest)
                  .then((realFields: FieldNames): void => {
                    for (let header of csvHeaders) {
                      let fieldFound: boolean = false;

                      for (let spField of realFields.value) {
                        if (header === spField.InternalName) {
                          fieldFound = true;
                          break;
                        }
                      }

                      if (!fieldFound) {
                        cmd.log(`Field ${header} was not found in the list. Valid fields:`);

                        for (let realField of realFields.value) {
                          cmd.log(realField.InternalName);
                        }

                        cb(new CommandError(`Field ${header} was not found in the list`));
                      }
                    }

                    lineNumber++
                    this.push(row);
                    callback();
                  })
                  .catch((error) => {
                    cb(new CommandError(error))
                  });
              }
              else {
                // Process csv Data
                lineNumber++;
                rowsInBatch++;

                const requestBody: any = {
                  formValues: SpoListItemBatchAddCommand.mapRequestBody(row, csvHeaders)
                };

                if (args.options.folder) {
                  requestBody.listItemCreateInfo = {
                    FolderPath: {
                      DecodedUrl: folderServerRelativeUrl
                    }
                  };
                }

                if (args.options.contentType && contentTypeName !== '') {
                  requestBody.formValues.push({
                    FieldName: 'ContentType',
                    FieldValue: contentTypeName
                  });
                }

                // row is ready
                recordsToAdd += '--changeset_' + changeSetId + '\r\n' +
                  'Content-Type: application/http' + '\r\n' +
                  'Content-Transfer-Encoding: binary' + '\r\n' +
                  '\r\n' +
                  'POST ' + endpoint + ' HTTP/1.1' + '\r\n' +
                  'Content-Type: application/json;odata=verbose' + '\r\n' +
                  'Accept: application/json;odata=verbose' + '\r\n' +
                  '\r\n' +
                  `${JSON.stringify(requestBody)}` + '\r\n' +
                  '\r\n';

                /***  Send the batch if the buffer is getting full **/
                if (rowsInBatch >= batchSize || recordsToAdd.length >= maxBytesInBatch) {
                  recordsToAdd += '--changeset_' + changeSetId + '--' + '\r\n';
                  ++batchCounter;
                  SpoListItemBatchAddCommand
                    .sendABatch(batchCounter, rowsInBatch, changeSetId, recordsToAdd, args.options.webUrl, verboseMode, cmd)
                    .catch(e => cb(new CommandError(e)))
                    .then((response: string | void) => {
                      SpoListItemBatchAddCommand.parseResults(response as string, cmd, cb);
                      recordsToAdd = ``;
                      rowsInBatch = 0;
                      changeSetId = v4();
                      this.push(row);
                      callback();
                    });
                }
                else {
                  this.push(row);
                  callback();
                }
              }
            },
          }))
          .on("data", function () { }) // don't delete this, or onEnd won't fire
          .on("end", function () {
            if (recordsToAdd.length > 0) {
              ++batchCounter;
              recordsToAdd += '--changeset_' + changeSetId + '--' + '\r\n';

              if (verboseMode) {
                cmd.log(`Sending final batch #${batchCounter} with ${rowsInBatch} items`);
              }

              SpoListItemBatchAddCommand
                .sendABatch(batchCounter, rowsInBatch, changeSetId, recordsToAdd, args.options.webUrl, verboseMode, cmd)
                .catch(e => cb(new CommandError(e)))
                .then((response: string | void) => {
                  SpoListItemBatchAddCommand.parseResults(response as string, cmd, cb);
                })
                .finally(() => {
                  cmd.log(`Processed ${lineNumber} Rows`);
                  cb();
                });
            }
            else {
              cmd.log(`Processed ${lineNumber} Rows`);
              cb();
            }
          })
          .on("error", (error: any) => {
            cb(error);
          });
      });
  }

  private static parseResults(response: string, cmd: CommandInstance, cb: (err?: any) => void): void {
    const responseLines: string[] = response.toString().split('\n');

    // read each line until you find JSON... 
    for (let responseLine of responseLines) {
      try {
        // check for error
        // any 500 errors (like timeout), just stop
        if (responseLine.startsWith("HTTP/1.1 5")) {
          cmd.log("An HTTP 5xx error was returned from SharePoint. Please retry with a lower --batchSize")
          cb(new CommandError(responseLine));
        }

        // parse the JSON response
        const responseJson: any = JSON.parse(responseLine);
        for (let result of responseJson.d.AddValidateUpdateItemUsingPath.results) {
          if (result.HasException) {
            cmd.log(result);
          }
        }
      }
      catch { }
    }
  }

  private static mapRequestBody(row: any, csvHeaders: string[]): { FieldName: string; FieldValue: string; }[] {
    const requestBody: { FieldName: string; FieldValue: string; }[] = [];
    Object.keys(row).forEach(async key => {
      requestBody.push({ FieldName: csvHeaders[parseInt(key)], FieldValue: (<any>row)[key] });
    });

    return requestBody;
  }

  private static sendABatch(batchCounter: number, rowsInBatch: number, changeSetId: string, recordsToAdd: string, webUrl: string, verbose: boolean, cmd: CommandInstance): Promise<string> {
    const batchContents: string[] = [];
    const batchId = v4();
    batchContents.push('--batch_' + batchId);

    if (verbose) {
      cmd.log(`Sending batch #${batchCounter} with ${rowsInBatch} items`);
    }

    batchContents.push('Content-Type: multipart/mixed; boundary="changeset_' + changeSetId + '"');
    batchContents.push('Content-Length: ' + recordsToAdd.length);
    batchContents.push('Content-Transfer-Encoding: binary');
    batchContents.push('');
    batchContents.push(recordsToAdd);

    batchContents.push('--batch_' + batchId + '--');

    const requestOptions: requestPromise.OptionsWithUrl = {
      url: `${webUrl}/_api/$batch`,
      headers: {
        'Content-Type': `multipart/mixed; boundary="batch_${batchId}"`
      },
      body: batchContents.join('\r\n')
    };

    return request.post(requestOptions);
  }

  private static async validateContentType(contentTypeName: string | undefined, listRestUrl: string, verbose: boolean, cmd: CommandInstance): Promise<string | void> {
    if (contentTypeName === undefined) {
      return Promise.resolve();
    }

    if (verbose) {
      cmd.log(`Getting content types for list...`);
    }

    const ctRequestOptions: requestPromise.OptionsWithUrl = {
      url: `${listRestUrl}/contenttypes?$select=Name,Id`,
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    return request
      .get<ContentTypes>(ctRequestOptions)
      .then((response: ContentTypes): Promise<void> => {
        const foundContentType = response.value.filter(ct => {
          const contentTypeMatch: boolean = ct.Id.StringValue === contentTypeName || ct.Name === contentTypeName;

          if (verbose) {
            cmd.log(`Checking content type ${ct.Name}: ${contentTypeMatch}`);
          }

          return contentTypeMatch;
        });

        if (verbose) {
          cmd.log('content type filter output...');
          cmd.log(foundContentType);
        }

        if (foundContentType.length !== 1) {
          return Promise.reject(`Specified content type '${contentTypeName}' doesn't exist on the target list`);
        }
        else {
          return Promise.resolve();
        }
      });
  }

  private static getFolderUrl(folderName: string | undefined, listRestUrl: string, webUrl: string, verbose: boolean, debug: boolean, cmd: CommandInstance): Promise<string> {
    if (folderName === undefined) {
      const requestOptions: requestPromise.OptionsWithUrl = {
        url: listRestUrl + "/RootFolder",
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      };
      return request
        .get<GetRootFolderResponse>(requestOptions)
        .then((response: GetRootFolderResponse): Promise<string> => Promise.resolve(response.ServerRelativeUrl));
    }
    else {
      const requestOptions: requestPromise.OptionsWithUrl = {
        url: `${listRestUrl}/rootFolder`,
        headers: {
          'accept': 'application/json;odata=nometadata'
        },
        json: true
      }

      let targetFolderServerRelativeUrl: string;

      return request
        .get<GetRootFolderResponse>(requestOptions)
        .then((rootFolderResponse: GetRootFolderResponse) => {
          targetFolderServerRelativeUrl = Utils.getServerRelativePath(rootFolderResponse.ServerRelativeUrl, folderName);
          const folderExtensions: FolderExtensions = new FolderExtensions(cmd, debug);
          return folderExtensions.ensureFolder(webUrl, targetFolderServerRelativeUrl);
        })
        .then(_ => Promise.resolve(targetFolderServerRelativeUrl));
    }
  }

  private static getCaseSensitiveWebUrl(webUrl: string): Promise<string> {
    const requestOptions: requestPromise.OptionsWithUrl = {
      url: webUrl + "/_api/web",
      headers: {
        'accept': 'application/json;odata=nometadata'
      },
      json: true
    };

    return request
      .get<GetWebResponse>(requestOptions)
      .then((response: GetWebResponse): Promise<string> => Promise.resolve(response.Url));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'URL of the site where the list is located'
      },
      {
        option: '-l, --listId [listId]',
        description: 'ID of the list where items should be added. Specify listId or listTitle but not both'
      },
      {
        option: '-t, --listTitle [listTitle]',
        description: 'Title of the list where items should be added. Specify listId or listTitle but not both'
      },
      {
        option: '-p, --path <path>',
        description: 'Path of the csv file with records to be added to the list'
      },
      {
        option: '-c, --contentType [contentType]',
        description: 'Name or the ID of the content type to associate with new items'
      },
      {
        option: '-f, --folder [folder]',
        description: 'List-relative URL of the folder where items should be created'
      },
      {
        option: '-b, --batchSize [batchSize]',
        description: 'Maximum number of records to send in a batch (default is 10)'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.webUrl) {
        return 'Required parameter webUrl missing';
      }

      const isValidSharePointUrl: boolean | string = SpoCommand.isValidSharePointUrl(args.options.webUrl);
      if (isValidSharePointUrl !== true) {
        return isValidSharePointUrl;
      }

      if (!args.options.listId && !args.options.listTitle) {
        return `Specify listId or listTitle`;
      }

      if (args.options.listId && args.options.listTitle) {
        return `Specify listId or listTitle but not both`;
      }

      if (!args.options.path) {
        return `Specify path`;
      }

      if (args.options.listId &&
        !Utils.isValidGuid(args.options.listId)) {
        return `${args.options.listId} in option listId is not a valid GUID`;
      }

      if (args.options.batchSize) {
        if (isNaN(args.options.batchSize)) {
          return `Specified batch size ${args.options.batchSize} is not a number`;
        }

        if (args.options.batchSize > 1000) {
          return `Batch size ${args.options.batchSize} exceeds the 1000 item limit`;
        }
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    The first row of the csv file must contain column headers. Column headers
    must match the internal name of the field in the list. If the column doesn't
    match a field in the list creating items will fail with an error.

    Rows in the csv file must contain column values based on the type of
    the field in the list:

    - Text: the text to add to the column
    - Number: the number to add to the column
    - Single-Select Metadata: the metadata name, followed by the pipe (|)
      character, followed by the metadata ID, followed by a semicolon, eg.
      TermLabel1|fa2f6bfd-1fad-4d18-9c89-289fe6941377;
    - Multi-Select Metadata: same format as single-select metadata, where
      multiple terms are separated with a semicolon, eg.
      TermLabel1|cf8c72a1-0207-40ee-aebd-fca67d20bc8a;
      TermLabel2|e5cc320f-8b65-4882-afd5-f24d88d52b75;
    - Single-Select Person: {'Key':'i:0#.f|membership|--UPN--'}
      where --UPN-- is the UPN of the person to add
    - Multi-Select Person: [{'Key':'i:0#.f|membership|--UPN1--'},
      {'Key':'i:0#.f|membership|--UPN2--'}] where --UPN1-- and --UPN2-- are
      UPNs of the persons to add
    - Hyperlink: the URL of the hyperlink followed by the text to be displayed
      for the hyperlink. The value must be enclosed in quotes, eg.
      "https://www.bing.com, Bing"
  
  Examples:
  
    Add items from file ${chalk.grey('data.csv')} to the specified list
      ${commands.LISTITEM_BATCH_ADD} --webUrl https://contoso.sharepoint.com/sites/project-x --listTitle "Events" --path test.csv
   `);
  }
}

module.exports = new SpoListItemBatchAddCommand();