import commands from '../commands';
import GlobalOptions from '../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../Command';
import request from '../../../request';
import AzmgmtCommand from '../../base/AzmgmtCommand';
import * as os from 'os';
import * as fs from 'fs';
import * as path from 'path';

const vorpal: Vorpal = require('../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  environment: string;
  path: string;
}

class FlowImportCommand extends AzmgmtCommand {
  public get name(): string {
    return commands.FLOW_IMPORT;
  }

  public get description(): string {
    return 'Gets information about the specified Microsoft Flow';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    if (this.verbose) {
      cmd.log(`Retrieving information about Microsoft Flow...`);
    }
    let blobUrl = '';

    const requestOptions1: any = {
      url: `${this.resource}providers/Microsoft.BusinessAppPlatform/environments/${encodeURIComponent(args.options.environment)}/generateResourceStorage?api-version=2016-11-01`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    request
      .post(requestOptions1)
      .then((res: any) => {


        console.log('HERHSEHSEHRHESRH--------------->')
        cmd.log('JSON.stringify(res)');
        cmd.log(JSON.stringify(res));

        // open file
        var fileName = path.basename(args.options.path);

        cmd.log('fileName');
        cmd.log(fileName);

        const fileBody: Buffer = fs.readFileSync(args.options.path);

        const urlDetails = res.sharedAccessSignature.split('?');
        blobUrl = `${urlDetails[0]}/${fileName}?${urlDetails[1]}`;

        const requestOptions: any = {
          url: blobUrl,
          body: fileBody,
          headers: {
            accept: 'application/json',
            'x-ms-blob-type':'BlockBlob',
            'x-anonymous': true
          }
        };

        return request.put(requestOptions)
       })
       .then((res: any) => {
        const requestOptions3: any = {
          url: `${this.resource}providers/Microsoft.BusinessAppPlatform/environments/${encodeURIComponent(args.options.environment)}/listImportParameters?api-version=2016-11-01`,
          headers: {
            body: { packageLink: { value: blobUrl } },
            accept: 'application/json'
          },
          json: true
        };

        request.post(requestOptions3);
       })
       .then((res:any) => {

        console.log('listImportOperations----listImportOperations');
        console.log(res);
        cmd.log(JSON.stringify(res));

        const requestOptions4: any = {
          url: `${this.resource}providers/Microsoft.BusinessAppPlatform/environments/${encodeURIComponent(args.options.environment)}/importPackage?api-version=2016-11-01`,
          headers: {
            accept: 'application/json',
            body: { packageLink: { value: blobUrl } }
          },
          json: true
        };

        request.post(requestOptions4); 
       })
      .then((res: any): void => {

        console.log('listImportOperations----listImportOperations');
        console.log(res);
        cmd.log(JSON.stringify(res));

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '--path <path>',
        description: 'The path of the zip'
      },
      {
        option: '-e, --environment <environment>',
        description: 'The name of the environment for which to retrieve available Flows'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.path) {
        return 'Required option name missing';
      }

      if (!args.options.environment) {
        return 'Required option environment missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.FLOW_IMPORT).helpInformation());
    log(
      `  Remarks:

    ${chalk.yellow('Attention:')} This command is based on an API that is currently
    in preview and is subject to change once the API reached general
    availability.
  
    By default, the command will try to retrieve Microsoft Flows you own.
    If you want to retrieve Flow owned by another user, use the ${chalk.blue('asAdmin')}
    flag.

    If the environment with the name you specified doesn't exist, you will get
    the ${chalk.grey('Access to the environment \'xyz\' is denied.')} error.

    If the Microsoft Flow with the name you specified doesn't exist, you will
    get the ${chalk.grey(`The caller with object id \'abc\' does not have permission${os.EOL}` +
        '    for connection \'xyz\' under Api \'shared_logicflows\'.')} error.
    If you try to retrieve a non-existing flow as admin, you will get the
    ${chalk.grey('Could not find flow \'xyz\'.')} error.
   
  Examples:
  
    Get information about the specified Microsoft Flow owned by the currently
    signed-in user
      ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d

    Get information about the specified Microsoft Flow owned by another user
      ${this.getCommandName()} --environment Default-d87a7535-dd31-4437-bfe1-95340acd55c5 --name 3989cb59-ce1a-4a5c-bb78-257c5c39381d --asAdmin
`);
  }
}

module.exports = new FlowImportCommand();