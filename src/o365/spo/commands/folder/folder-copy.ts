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

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  webUrl: string;
  sourceUrl: string;
  targetUrl: string;
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

class SpoFolderCopyCommand extends SpoCommand {

  public get name(): string {
    return commands.FOLDER_COPY;
  }

  public get description(): string {
    return 'Copies a folder to another location';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
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
   * A polling function that awaits the 
   * Azure queued copy job to return JobStatus = 0 meaning it is done with the task.
   */
  public getCopyJobProgress(opts: JobProgressOptions, cmd: any): 
     Promise<any> {
       
    let pollCount = 0;
    let retryAttemptsCount = 0;

    const checkCondition = (resolve: any, reject: any) => {
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
      
      request.post(requestOptions).then((resp: any) => {
        retryAttemptsCount = 0; // clear retry on promise success 

        if (this.debug) {
          cmd.log('getCopyJobProgress response...');
          cmd.log(resp);
        }

        if(this.verbose) {
          if(resp.JobState &&  resp.JobState === 4) {
            cmd.log(`Check #${pollCount}. Copy job in progress... JobState: ${resp.JobState}`);
          } else {
            cmd.log(`Check #${pollCount}. JobState: ${resp.JobState}`);
          }
        }

        for(const item of resp.Logs) {
          const log = JSON.parse(item);
          
          // reject if progress error 
          if(log.Event === "JobError" || log.Event === "JobFatalError") {
            return reject(log.Message);
          }
        }

        // three possible scenarios
        // job done = success promise returned
        // job in progress = recursive call using setTimeout returned
        // max poll attempts flag rised = reject promise returned
        if (resp.JobState === 0) {
          // job done
          resolve();
        }
        else if (pollCount < opts.progressMaxPollAttempts) {
          // if the condition isn't met but the timeout hasn't elapsed, go again
          setTimeout(checkCondition, opts.progressPollInterval, resolve, reject);
        }
        else {
          reject(new Error('getCopyJobProgress timed out'));
        }
      }, 
      (error: any) => {
        retryAttemptsCount++;
        
        // let's retry x times in row before we give up since
        // this is progress check and even if rejects a promise
        // the actual copy process can success.
        if(retryAttemptsCount <= opts.progressRetryAttempts) {
          setTimeout(checkCondition, opts.progressPollInterval, resolve, reject);
        }else{
          reject(error);
        }
      });
    }
    return new Promise<any>(checkCondition);
  }

  /**
   * Combines base and relative url 
   * considering any missing slashes
   * @param baseUrl https://contoso.com
   * @param relativeUrl sites/abc
   */
  public urlCombine(baseUrl: string, relativeUrl: string): string {

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
        option: '-w, --webUrl <webUrl>',
        description: 'The URL of the site where the folder is located'
      },
      {
        option: '-u, --sourceUrl <sourceUrl>',
        description: 'Site-relative URL of the folder to copy'
      },
      {
        option: '-t, --targetUrl <targetUrl>',
        description: 'Server-relative URL where to copy the folder'
      },
      {
        option: '-d, --deleteIfAlreadyExists [deleteIfAlreadyExists]',
        description: 'If a folder already exists at the targetUrl, it will be moved to the recycle bin. If ommitted, the copy operation will be canceled if the folder already exists at the targetUrl location'
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
      `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
    using the ${chalk.blue(commands.CONNECT)} command.
  
  Remarks:
  
    To copy a folder, you have to first connect to SharePoint using the
    ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    When you use Copy to with documents that have version history, only the latest version is copied.
        
  Examples:
  
    Performs folder copy between two site collections for folder with name ${chalk.grey('MyFolder')}
    located in site document library ${chalk.grey('https://contoso.sharepoint.com/sites/test1/Shared%20Documents')}
      ${chalk.grey(config.delimiter)} ${commands.FOLDER_COPY} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test2/Shared%20Documents/

    Performs folder copy between two document libraries in the same site collection for folder with name ${chalk.grey('MyFolder')}
    located in site document library ${chalk.grey('https://contoso.sharepoint.com/sites/test1/Shared%20Documents')}
        ${chalk.grey(config.delimiter)} ${commands.FOLDER_COPY} --webUrl https://contoso.sharepoint.com/sites/test1 --sourceUrl /Shared%20Documents/MyFolder --targetUrl /sites/test1/HRDocuments/
    `);
  }
}

module.exports = new SpoFolderCopyCommand();