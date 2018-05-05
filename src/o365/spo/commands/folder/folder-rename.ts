import auth from '../../SpoAuth';
import config from '../../../../config';
import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import * as request from 'request-promise-native';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { Auth } from '../../../../Auth';
import * as url from 'url';
import { ContextInfo, ClientSvcResponse, ClientSvcResponseContents } from '../../spo';
import { ClientSvcCommons, IdentityResponse } from '../../common/client-svc-commons';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  folderUrl: string;
  name: string;
}

class SpoFolderRenameCommand extends SpoCommand {

  public get name(): string {
    return commands.FOLDER_RENAME;
  }

  public get description(): string {
    return 'Renames a folder';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    const clientSvcCommons: ClientSvcCommons = new ClientSvcCommons(cmd, this.debug);
    let siteAccessToken: string = '';
    let formDigestValue: string = '';
    let serverRelativeUrl: string = '';

    auth
    .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
    .then((accessToken: string): request.RequestPromise => {
      siteAccessToken = accessToken;

      if (this.debug) {
        cmd.log(`Retrieved access token ${accessToken}. Retrieving request digest...`);
      }

      return this.getRequestDigestForSite(args.options.webUrl, siteAccessToken, cmd, this.debug);
    })
    .then((contextResponse: ContextInfo): Promise<IdentityResponse> => {
      formDigestValue = contextResponse.FormDigestValue;

      if (this.debug) {
        cmd.log('contextResponse:');
        cmd.log(JSON.stringify(contextResponse));
        cmd.log('');
      }

      return clientSvcCommons.requestObjectIdentity(args.options.webUrl,siteAccessToken, formDigestValue);
    })
    .then((webObjectIdentity: IdentityResponse): Promise<IdentityResponse> => {
      
      if (this.debug) {
        cmd.log('IdentityResponse:');
        cmd.log(JSON.stringify(webObjectIdentity));
        cmd.log('');
      }

      let webRelativeUrl = this.getWebRelativeUrlFromWebUrl(args.options.webUrl);
      serverRelativeUrl = `${webRelativeUrl}${this.formatRelativeUrl(args.options.folderUrl)}`;

      return clientSvcCommons.requestFolderObjectIdentity(webObjectIdentity, args.options.webUrl, siteAccessToken, formDigestValue);
    })
    .then((folderObjectIdentity: IdentityResponse): Promise<void> => {
      if (this.debug) {
        cmd.log('IdentityResponse:');
        cmd.log(JSON.stringify(folderObjectIdentity));
        cmd.log('');
      }

      if (this.verbose) {
        cmd.log(`Renaming folder ${args.options.folderUrl} to ${args.options.name}`);
      }

      let webRelativeUrl = this.getWebRelativeUrlFromWebUrl(args.options.webUrl);
      let serverRelativeUrl: string = `${webRelativeUrl}${this.formatRelativeUrl(args.options.folderUrl)}`;
      // remove last '/' of url
      if (serverRelativeUrl.lastIndexOf('/') === serverRelativeUrl.length - 1) {
        serverRelativeUrl = serverRelativeUrl.substring(0, serverRelativeUrl.length - 1);
      }
      const renamedServerRelativeUrl = `${serverRelativeUrl.substring(0, serverRelativeUrl.lastIndexOf('/'))}/${args.options.name}`;
      cmd.log(renamedServerRelativeUrl);

      const requestOptions: any = {
        url: `${args.options.webUrl}/_vti_bin/client.svc/ProcessQuery`,
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${siteAccessToken}`,
          'X-RequestDigest': formDigestValue
        }),
        body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Method Name="MoveTo" Id="32" ObjectPathId="26"><Parameters><Parameter Type="String">${renamedServerRelativeUrl}</Parameter></Parameters></Method></Actions><ObjectPaths><Identity Id="26" Name="${folderObjectIdentity.objectIdentity}" /></ObjectPaths></Request>`
      };

      if (this.debug) {
        cmd.log('Executing web request...');
        cmd.log(requestOptions);
        cmd.log('');
      }

      return new Promise<void>((resolve: any, reject: any): void => {
        request.post(requestOptions).then((res: any) => {
  
          const json: ClientSvcResponse = JSON.parse(res);
          const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
          if (contents && contents.ErrorInfo) {
            return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
          }
          return resolve();

        }, (err: any): void => { reject(err); });
      });
    })
    .then((resp: any): void => {
      if (this.verbose) {
        cmd.log('DONE');
      }
      cb();
    }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
}

  public formatRelativeUrl(relativeUrl: string): string {

    // add '/' at 0
    if (relativeUrl.charAt(0) !== '/') {
      relativeUrl = `/${relativeUrl}`;
    }

    // remove last '/' of url
    if (relativeUrl.lastIndexOf('/') === relativeUrl.length - 1) {
      relativeUrl = relativeUrl.substring(0, relativeUrl.length - 1);
    }

    return relativeUrl;
  }

  public getWebRelativeUrlFromWebUrl(webUrl: string): string {

    const tenantUrl = `${url.parse(webUrl).protocol}//${url.parse(webUrl).hostname}`;
    let webRelativeUrl = webUrl.replace(tenantUrl, '');

    return this.formatRelativeUrl(webRelativeUrl);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the folder to be renamed'
      },
      {
        option: '-f, --folderUrl <folderUrl>',
        description: 'Site-relative URL of the folder (including the folder)'
      },
      {
        option: '-n, --name <name>',
        description: 'New name for the target folder'
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

      if (!args.options.folderUrl) {
        return 'Required parameter sourceUrl missing';
      }

      if (!args.options.name) {
        return 'Required parameter name missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
    using the ${chalk.blue(commands.CONNECT)} command.
  
  Remarks:
  
    To rename a folder, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.
        
  Examples:
  
    Renames a folder with site relative url ${chalk.grey('/Shared Documents/My Folder 1')}
    located in site ${chalk.grey('https://contoso.sharepoint.com/sites/project-x')}
      ${chalk.grey(config.delimiter)} ${commands.FOLDER_RENAME} --webUrl https://contoso.sharepoint.com/sites/project-x --folderUrl '/Shared Documents/My Folder 1' --name 'My Folder 2'
    `);
  }
}

module.exports = new SpoFolderRenameCommand();