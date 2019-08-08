import request from '../../../../request';
import commands from '../../commands';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import GlobalOptions from '../../../../GlobalOptions';
import { ContextInfo, FormDigestInfo } from '../../spo';
const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
}

class SpoAppSyncToTeamsCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_SYNCTOTEAMS;
  }

  public get description(): string {
    return 'Syncs SharePoint Framework solutions package to Microsoft Teams';
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: (err?: any) => void): void {
    let appCatalogSiteUrl: string;

    this
      .getSpoUrl(cmd, this.debug)
      .then((spoUrl: string): Promise<string> => {

        return this.getAppCatalogSiteUrl(cmd, spoUrl, args);
      })
      .then((appCatalog: string): Promise<FormDigestInfo> => {
        appCatalogSiteUrl = appCatalog;

        return this.getRequestDigest(appCatalogSiteUrl);
      })
      .then((res: ContextInfo): Promise<string> => {
        if (this.verbose) {
          cmd.log(`Syncing the app...`);
        }

        const requestOptions: any = {
          url: `${appCatalogSiteUrl}/_api/web/tenantappcatalog/SyncSolutionToTeams(id=${args.options.id})`,
          headers: {
            'X-RequestDigest': res.FormDigestValue,
            accept: 'application/json;odata=nometadata'
          },
          json: true
        };

        return request.post(requestOptions);
      })
      .then((): void => {
        if (this.verbose) {
          cmd.log(vorpal.chalk.green('DONE'));
        }

        cb();
      }, (rawRes: any): void => { 
        cmd.log('ERROR:');
        cmd.log(JSON.stringify(rawRes));
        this.handleRejectedODataPromise(rawRes, cmd, cb)});
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'List item id of the solution package in the tenant app catalog site'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Examples:
  
    Syncs SPFx solution to Microsoft Teams that has list item id of '1' in the tenant app catalog
      ${commands.APP_SYNCTOTEAMS} --id 1
  `);
  }
}

module.exports = new SpoAppSyncToTeamsCommand();