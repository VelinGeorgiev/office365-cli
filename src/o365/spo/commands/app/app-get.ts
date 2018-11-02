import auth from '../../SpoAuth';
import { Auth } from '../../../../Auth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import { AppMetadata } from './AppMetadata';
import Utils from '../../../../Utils';
import { SpoAppBaseCommand } from './SpoAppBaseCommand';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appCatalogUrl?: string;
  id?: string;
  name?: string;
  scope?: string;
  siteUrl?: string;
}

class SpoAppGetCommand extends SpoAppBaseCommand {
  public get name(): string {
    return commands.APP_GET;
  }

  public get description(): string {
    return 'Gets information about the specific app from the specified app catalog';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.id = typeof args.options.id !== 'undefined';
    telemetryProps.name = typeof args.options.name !== 'undefined';
    telemetryProps.appCatalogUrl = typeof args.options.appCatalogUrl !== 'undefined';
    telemetryProps.scope = (!(!args.options.scope)).toString();
    telemetryProps.siteUrl = (!(!args.options.siteUrl)).toString();
    return telemetryProps;
  }

  protected requiresTenantAdmin(): boolean {
    return false;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const scope: string = (args.options.scope) ? args.options.scope.toLowerCase() : 'tenant';
    let siteAccessToken: string = '';
    let appCatalogSiteUrl: string = '';

    this.getAppCatalogSiteUrl(cmd, auth.site.url, auth.service.accessToken, args)
      .then((siteUrl: string): Promise<string> => {
        appCatalogSiteUrl = siteUrl;

        const resource: string = Auth.getResourceFromUrl(appCatalogSiteUrl);
        return auth.getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug);
      })
      .then((accessToken: string): Promise<{ UniqueId: string }> | request.RequestPromise => {
        siteAccessToken = accessToken;

        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}...`);
        }

        if (args.options.id) {
          return Promise.resolve({ UniqueId: args.options.id });
        }

        if (this.verbose) {
          cmd.log(`Looking up app id for app named ${args.options.name}...`);
        }

        const requestOptions: any = {
          url: `${appCatalogSiteUrl}/_api/web/getfolderbyserverrelativeurl('AppCatalog')/files('${args.options.name}')?$select=UniqueId`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            accept: 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: { UniqueId: string }): request.RequestPromise => {
        if (this.verbose) {
          cmd.log(`Retrieving information for app ${res}...`);
        }

        const requestOptions: any = {
          url: `${appCatalogSiteUrl}/_api/web/${scope}appcatalog/AvailableApps/GetById('${encodeURIComponent(res.UniqueId)}')`,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            accept: 'application/json;odata=nometadata'
          }),
          json: true
        };

        if (this.debug) {
          cmd.log('Executing web request...');
          cmd.log(requestOptions);
          cmd.log('');
        }

        return request.get(requestOptions);
      })
      .then((res: AppMetadata): void => {
        if (this.debug) {
          cmd.log('Response:');
          cmd.log(res);
          cmd.log('');
        }

        cmd.log(res);

        cb();
      }, (rawRes: any): void => this.handleRejectedODataJsonPromise(rawRes, cmd, cb));
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id [id]',
        description: 'ID of the app to retrieve information for. Specify the id or the name but not both'
      },
      {
        option: '-n, --name [name]',
        description: 'Name of the app to retrieve information for. Specify the id or the name but not both'
      },
      {
        option: '-u, --appCatalogUrl [appCatalogUrl]',
        description: 'URL of the tenant app catalog site. If not specified, the CLI will try to resolve it automatically'
      },
      {
        option: '-s, --scope [scope]',
        description: 'Scope of the app catalog: tenant|sitecollection. Default tenant',
        autocomplete: ['tenant', 'sitecollection']
      },
      {
        option: '--siteUrl [siteUrl]',
        description: 'The URL of the site collection with app catalog where the solution package is located. Must be specified when the scope is \'sitecollection\''
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
          return `Scope must be either 'tenant' or 'sitecollection'`
        }

        if (testScope === 'sitecollection' && !args.options.siteUrl) {
          if (args.options.appCatalogUrl) {
            return `You must specify siteUrl when the scope is sitecollection instead of appCatalogUrl`;
          }
          return `You must specify siteUrl when the scope is sitecollection`;
        }
        else if (testScope === 'tenant' && args.options.siteUrl) {
          return `The siteUrl option can only be used when the scope option is set to sitecollection`;
        }
      }

      if (!args.options.id && !args.options.name) {
        return 'Specify either the id or the name';
      }

      if (args.options.id && args.options.name) {
        return 'Specify either the id or the name but not both';
      }

      if (args.options.id && !Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      if (args.options.appCatalogUrl) {
        return SpoAppBaseCommand.isValidSharePointUrl(args.options.appCatalogUrl);
      }

      if (!args.options.scope && args.options.siteUrl) {
        return `The siteUrl option can only be used when the scope option is set to sitecollection`;
      }

      if (args.options.siteUrl) {
        return SpoAppBaseCommand.isValidSharePointUrl(args.options.siteUrl);
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APP_GET).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
    using the ${chalk.blue(commands.LOGIN)} command.

  Remarks:
  
    To get information about the specified app available in the tenant app catalog,
    you have to first log in to a SharePoint site using the ${chalk.blue(commands.LOGIN)} command,
    eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.
   
  Examples:
  
    Return details about the app with ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    available in the tenant app catalog.
      ${chalk.grey(config.delimiter)} ${commands.APP_GET} --id b2307a39-e878-458b-bc90-03bc578531d6

    Return details about the app with name ${chalk.grey('solution.sppkg')}
    available in the tenant app catalog. Will try to detect the app catalog URL
      ${chalk.grey(config.delimiter)} ${commands.APP_GET} --name solution.sppkg

    Return details about the app with name ${chalk.grey('solution.sppkg')}
    available in the tenant app catalog using the specified app catalog URL
      ${chalk.grey(config.delimiter)} ${commands.APP_GET} --name solution.sppkg --appCatalogUrl https://contoso.sharepoint.com/sites/apps

    Return details about the app with ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}
    available in the site collection app catalog
    of site ${chalk.grey('https://contoso.sharepoint.com/sites/site1')}.
      ${chalk.grey(config.delimiter)} ${commands.APP_GET} --id b2307a39-e878-458b-bc90-03bc578531d6 --scope sitecollection --siteUrl https://contoso.sharepoint.com/sites/site1

  More information:
  
    Application Lifecycle Management (ALM) APIs
      https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
`);
  }
}

module.exports = new SpoAppGetCommand();