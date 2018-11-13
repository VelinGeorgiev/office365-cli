import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import { AppMetadata } from './AppMetadata';
import Utils from '../../../../Utils';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  siteUrl?: string;
  scope?: string;
}

class AppListCommand extends SpoCommand {
  public get name(): string {
    return commands.APP_LIST;
  }

  public get description(): string {
    return 'Lists apps from the tenant app catalog';
  }

  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.siteUrl = (!(!args.options.siteUrl)).toString();
    telemetryProps.scope = (!(!args.options.scope)).toString();
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    auth
      .ensureAccessToken(auth.service.resource, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}. Loading apps from tenant app catalog...`);
        }

        if (this.verbose) {
          cmd.log(`Retrieving apps...`);
        }

        let siteUrl = auth.site.url;
        if(args.options.siteUrl) {
          siteUrl = args.options.siteUrl;
        }

        const requestOptions: any = {
          url: `${siteUrl}/_api/web/${scope}appcatalog/AvailableApps`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${accessToken}`,
            accept: 'application/json;odata=nometadata'
          })
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: string): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        const apps: { value: AppMetadata[] } = JSON.parse(res);

        if (apps.value && apps.value.length > 0) {
          if (args.options.output === 'json') {
            cmd.log(apps.value);
          }
          else {
            cmd.log(apps.value.map(a => {
              return {
                Title: a.Title,
                ID: a.ID,
                Deployed: a.Deployed,
                AppCatalogVersion: a.AppCatalogVersion
              };
            }));
          }
        }
        else {
          if (this.verbose) {
            cmd.log('No apps found');
          }
        }
        cb();
      }, (rawRes: any): void => this.handleRejectedODataPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-s, --scope [scope]',
        description: 'Specify the target app catalog: \'tenant\' or \'sitecollection\' (default = tenant)',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '-u, --siteUrl [siteUrl]',
        description: 'Absolute URL of the site to list the apps'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      // verify either 'tenant' or 'sitecollection' specified if scope provided
      if (args.options.scope) {
        const testScope: string = args.options.scope.toLowerCase();
        if (!(testScope === 'tenant' || testScope === 'sitecollection')) {
          return `Scope must be either 'tenant' or 'sitecollection' if specified`
        }

        if (testScope === 'tenant' && args.options.siteUrl){
          return `siteUrl must be used with scope 'sitecollection'`;
        }

        if (args.options.siteUrl) {
          return SpoCommand.isValidSharePointUrl(args.options.siteUrl);
        }
      } else if (args.options.siteUrl) {
          return `siteUrl must be used with scope 'sitecollection'`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_LIST).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site, using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:

    To list apps from the tenant app catalog, you have to first log in to a SharePoint site using
    the ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    When using the text output type (default), the command lists only the values of the ${chalk.grey('Title')},
    ${chalk.grey('ID')}, ${chalk.grey('Deployed')} and ${chalk.grey('AppCatalogVersion')} properties of the app. When setting the output
    type to JSON, all available properties are included in the command output.
   
  Examples:
  
    Return the list of available apps from the tenant app catalog. Show the installed version in the site if applicable.
      ${chalk.grey(config.delimiter)} ${commands.APP_LIST} --scope tenant

    Return the list of available apps from a site collection app catalog. Show the installed version in the site if applicable.
      ${chalk.grey(config.delimiter)} ${commands.APP_LIST} --scope sitecollection --siteUrl https://contoso.sharepoint.com/sites/foo

    Return the list of available apps from a site collection app catalog of the site you are currently logged in. 
    Show the installed version in the site if applicable.
      ${chalk.grey(config.delimiter)} ${commands.APP_LIST} --scope sitecollection

  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new AppListCommand();