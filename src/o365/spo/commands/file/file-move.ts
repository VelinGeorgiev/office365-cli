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
import { ContextInfo } from '../../spo';
import * as url from 'url';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetUrl: string;
  deleteIfAlreadyExists?: boolean;
}

interface JobProgressOptions {
  webUrl: string;
  accessToken: string;
  /**
   * Response object retrieved from /_api/site/CreateCopyJobs
   */
  copyJopInfo: any;
  /**
   * Poll interval to call /_api/site/GetCopyJobProgress
   */
  progressPollInterval: number;
  /**
   * Max poll intervals to call /_api/site/GetCopyJobProgress
   * after which to give up
   */
  progressMaxPollAttempts: number;
  /**
   * Retry attempts before give up.
   * Give up if /_api/site/GetCopyJobProgress returns 
   * X reject promises in a row
   */
  progressRetryAttempts: number;
}

class SpoFilemoveCommand extends SpoCommand {
  public get name(): string {
    return commands.FILE_COPY;
  }

  public get description(): string {
    return 'Copies a file to another location';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.deleteIfAlreadyExists = args.options.deleteIfAlreadyExists || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const resource: string = Auth.getResourceFromUrl(args.options.webUrl);
    let siteAccessToken = '';
    const webUrl = args.options.webUrl;
    const tenantUrl = `${url.parse(webUrl).protocol}//${url.parse(webUrl).hostname}`;

    if (this.debug) {
      cmd.log(`Retrieving access token for ${resource}...`);
    }

    auth
      .getAccessToken(resource, auth.service.refreshToken as string, cmd, this.debug)
      .then((accessToken: string): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`Retrieved access token ${accessToken}.`);
        }

        siteAccessToken = accessToken;

        // Check if the source file exists.
        // Called on purpose, we explicitly check if user specified file
        // in the sourceUrl option. 
        // The CreateCopyJobs endpoint accepts file, folder or batch from both.
        // A user might enter folder instead of file as source url by mistake
        // then there are edge cases when deleteIfAlreadyExists flag is set
        // the user can receive misleading error message.
        return this.fileExists(tenantUrl, webUrl, args.options.sourceUrl, siteAccessToken, cmd);
      })
      .then((resp: any): Promise<void> => {
        if (this.debug) {
          cmd.log(`fileExists response...`);
          cmd.log(resp);
        }

        if (args.options.deleteIfAlreadyExists) {
          // try delete target file, if deleteIfAlreadyExists flag is set
          const filename = args.options.sourceUrl.replace(/^.*[\\\/]/, '');
          return this.recycleFile(tenantUrl, args.options.targetUrl, filename, siteAccessToken, cmd);
        }

        return Promise.resolve();
      })
      .then((resp: any): request.RequestPromise => {
        if (this.debug) {
          cmd.log(`deleteIfAlreadyExists response`);
          cmd.log(resp);
        }

        // all preconditions met, now create copy job
        const sourceAbsoluteUrl = this.urlCombine(webUrl, args.options.sourceUrl);
        const requestUrl: string = `${webUrl}/_api/site/CreateCopyJobs`;
        const requestOptions: any = {
          url: requestUrl,
          headers: Utils.getRequestHeaders({
            authorization: `Bearer ${siteAccessToken}`,
            'accept': 'application/json;odata=nometadata'
          }),
          body: {
            exportObjectUris: [sourceAbsoluteUrl],
            destinationUri: this.urlCombine(tenantUrl, args.options.targetUrl),
            options: { "IgnoreVersionHistory": true }
          },
          json: true
        };

        if (this.debug) {
          cmd.log('CreateCopyJobs request...');
          cmd.log(requestOptions);
        }

        return request.post(requestOptions);
      })
      .then((jobInfo: any): Promise<any> => {

        if (this.debug) {
          cmd.log('CreateCopyJobs response...');
          cmd.log(jobInfo);
        }

        const jobProgressOptions: JobProgressOptions = {
          webUrl: webUrl,
          accessToken: siteAccessToken,
          copyJopInfo: jobInfo.value[0],
          progressMaxPollAttempts: 1000, // 1 sec.
          progressPollInterval: 30 * 60, // approx. 30 mins. if interval is 1000
          progressRetryAttempts: 5
        }

        return this.getCopyJobProgress(jobProgressOptions, cmd);
      })
      .then((resp: any): void => {
        if (this.verbose) {
          cmd.log('DONE');
        }
        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  /**
   * Checks if a file exists on the server relative url
   */
  private fileExists(tenantUrl: string, webUrl: string, sourceUrl: string, siteAccessToken: string, cmd: any): request.RequestPromise {
    const webServerRelativeUrl: string = webUrl.replace(tenantUrl, '');
    const fileServerRelativeUrl: string = `${webServerRelativeUrl}${sourceUrl}`;

    const requestUrl = `${webUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(fileServerRelativeUrl)}')/`;
    const requestOptions: any = {
      url: requestUrl,
      method: 'GET',
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${siteAccessToken}`,
        'accept': 'application/json;odata=nometadata'
      }),
      json: true
    };

    if (this.debug) {
      cmd.log(`fileExists request...`);
      cmd.log(requestOptions);
    }

    return request.get(requestOptions);
  }

  /**
   * A polling function that awaits the 
   * queued copy job to return JobStatus = 0 meaning it is done with the task.
   */
  private getCopyJobProgress(opts: JobProgressOptions, cmd: CommandInstance): Promise<void> {
    let pollCount: number = 0;
    let retryAttemptsCount: number = 0;

    const checkCondition = (resolve: () => void, reject: (error: any) => void): void => {
      pollCount++;
      const requestUrl: string = `${opts.webUrl}/_api/site/GetCopyJobProgress`;
      const requestOptions: any = {
        url: requestUrl,
        headers: Utils.getRequestHeaders({
          authorization: `Bearer ${opts.accessToken}`,
          'accept': 'application/json;odata=nometadata'
        }),
        body: { "copyJobInfo": opts.copyJopInfo },
        json: true
      };

      if (this.debug) {
        cmd.log('getCopyJobProgress request...');
        cmd.log(requestOptions);
      }

      request.post(requestOptions).then((resp: any): void => {
        retryAttemptsCount = 0; // clear retry on promise success 

        if (this.debug) {
          cmd.log('getCopyJobProgress response...');
          cmd.log(resp);
        }

        if (this.verbose) {
          if (resp.JobState && resp.JobState === 4) {
            cmd.log(`Check #${pollCount}. Copy job in progress... JobState: ${resp.JobState}`);
          }
          else {
            cmd.log(`Check #${pollCount}. JobState: ${resp.JobState}`);
          }
        }

        for (const item of resp.Logs) {
          const log = JSON.parse(item);

          // reject if progress error 
          if (log.Event === "JobError" || log.Event === "JobFatalError") {
            return reject(log.Message);
          }
        }

        // three possible scenarios
        // job done = success promise returned
        // job in progress = recursive call using setTimeout returned
        // max poll attempts flag raised = reject promise returned
        if (resp.JobState === 0) {
          // job done
          resolve();
          return;
        }

        if (pollCount < opts.progressMaxPollAttempts) {
          // if the condition isn't met but the timeout hasn't elapsed, go again
          setTimeout(checkCondition, opts.progressPollInterval, resolve, reject);
        }
        else {
          reject(new Error('getCopyJobProgress timed out'));
        }
      },
        (error: any): void => {
          retryAttemptsCount++;

          // let's retry x times in row before we give up since
          // this is progress check and even if rejects a promise
          // the actual copy process can success.
          if (retryAttemptsCount <= opts.progressRetryAttempts) {
            setTimeout(checkCondition, opts.progressPollInterval, resolve, reject);
          } else {
            reject(error);
          }
        });
    }

    return new Promise<void>(checkCondition);
  }

  /**
   * Moves file in the site recycle bin
   */
  private recycleFile(tenantUrl: string, targetUrl: string, filename: string, siteAccessToken: string, cmd: CommandInstance): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const targetFolderAbsoluteUrl: string = this.urlCombine(tenantUrl, targetUrl);

      // since the target WebFullUrl is unknown we can use getRequestDigestForSite
      // to get it from target folder absolute url.
      // Similar approach used here Microsoft.SharePoint.Client.Web.WebUrlFromFolderUrlDirect
      this.getRequestDigestForSite(targetFolderAbsoluteUrl, siteAccessToken, cmd, this.debug)
        .then((contextResponse: ContextInfo): void => {
          if (this.debug) {
            cmd.log(`contextResponse.WebFullUrl: ${contextResponse.WebFullUrl}`);
          }

          if (targetUrl.charAt(0) !== '/') {
            targetUrl = `/${targetUrl}`;
          }
          if (targetUrl.lastIndexOf('/') !== targetUrl.length - 1) {
            targetUrl = `${targetUrl}/`;
          }

          const requestUrl: string = `${contextResponse.WebFullUrl}/_api/web/GetFileByServerRelativeUrl('${encodeURIComponent(`${targetUrl}${filename}`)}')/recycle()`;
          const requestOptions: any = {
            url: requestUrl,
            method: 'POST',
            headers: Utils.getRequestHeaders({
              authorization: `Bearer ${siteAccessToken}`,
              'X-HTTP-Method': 'DELETE',
              'If-Match': '*',
              'accept': 'application/json;odata=nometadata'
            }),
            json: true
          };

          if (this.debug) {
            cmd.log(`recycleFile request...`);
            cmd.log(requestOptions);
          }

          request.post(requestOptions)
            .then((resp: any): void => {
              resolve();
            })
            .catch((err: any): any => {
              if (err.statusCode === 404) {
                // file does not exist so can proceed
                return resolve();
              }

              if (this.debug) {
                cmd.log(`recycleFile error...`);
                cmd.log(err);
              }

              reject(err);
            });
        }, (e: any) => reject(e));
    });
  }

  /**
   * Combines base and relative url considering any missing slashes
   * @param baseUrl https://contoso.com
   * @param relativeUrl sites/abc
   */
  private urlCombine(baseUrl: string, relativeUrl: string): string {
    // remove last '/' of base if exists
    if (baseUrl.lastIndexOf('/') === baseUrl.length - 1) {
      baseUrl = baseUrl.substring(0, baseUrl.length - 1);
    }

    // remove '/' at 0
    if (relativeUrl.charAt(0) === '/') {
      relativeUrl = relativeUrl.substring(1, relativeUrl.length);
    }

    // remove last '/' of next if exists
    if (relativeUrl.lastIndexOf('/') === relativeUrl.length - 1) {
      relativeUrl = relativeUrl.substring(0, relativeUrl.length - 1);
    }

    return `${baseUrl}/${relativeUrl}`;
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-u, --webUrl <webUrl>',
        description: 'The URL of the site where the file is located'
      },
      {
        option: '-s, --sourceUrl <sourceUrl>',
        description: 'Site-relative URL of the file to copy'
      },
      {
        option: '-t, --targetUrl <targetUrl>',
        description: 'Server-relative URL where to copy the file'
      },
      {
        option: '--deleteIfAlreadyExists',
        description: 'If a file already exists at the targetUrl, it will be moved to the recycle bin. If omitted, the copy operation will be canceled if the file already exists at the targetUrl location'
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

      if (!args.options.sourceUrl) {
        return 'Required parameter sourceUrl missing';
      }

      if (!args.options.targetUrl) {
        return 'Required parameter targetUrl missing';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} before using this command, log in to a SharePoint Online site,
    using the ${chalk.blue(commands.LOGIN)} command.
  
  Remarks:
  
    To copy a file, you have to first log in to SharePoint using the
    ${chalk.blue(commands.LOGIN)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.LOGIN} https://contoso.sharepoint.com`)}.

    When you copy a file using the ${chalk.grey(this.name)} command,
    only the latest version of the file is copied.
        
  Examples:
  
    Copy file to a document library in another site collection
      ${chalk.grey(config.delimiter)} ${commands.FILE_COPY} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test2/Shared%20Documents/

    Copy file to a document library in the same site collection
        ${chalk.grey(config.delimiter)} ${commands.FILE_COPY} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test1/HRDocuments/

    Copy file to a document library in another site collection. If a file with
    the same name already exists in the target document library, move it
    to the recycle bin
        ${chalk.grey(config.delimiter)} ${commands.FILE_COPY} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/sp1.pdf --targetUrl /sites/test2/Shared%20Documents/ --deleteIfAlreadyExists

  More information:

    Copy items from a SharePoint document library
      https://support.office.com/en-us/article/move-or-copy-items-from-a-sharepoint-document-library-00e2f483-4df3-46be-a861-1f5f0c1a87bc
    `);
  }
}

module.exports = new SpoFilemoveCommand();