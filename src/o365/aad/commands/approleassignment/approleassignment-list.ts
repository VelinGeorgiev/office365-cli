import commands from '../../commands';
import GlobalOptions from '../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../Command';
import Utils from '../../../../Utils';
import AadCommand from '../../../base/AadCommand';
import request from '../../../../request';
import { AppRoleAssignment } from './AppRoleAssignment';
import { ServicePrincipal } from './ServicePrincipal';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  appId?: string;
  displayName?: string;
}

class AadAppRoleAssignmentListCommand extends AadCommand {
  public get name(): string {
    return commands.APPROLEASSIGNMENT_LIST;
  }

  public get description(): string {
    return 'Lists AppRoleAssignments for the specified application registration';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.appId = typeof args.options.appId !== 'undefined';
    telemetryProps.displayName = typeof args.options.displayName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void  {
      let sp: ServicePrincipal;

      // get the service principal associated with the appId
      const spMatchQuery: string = args.options.appId ?
        `appId eq '${encodeURIComponent(args.options.appId)}'` :
        `displayName eq '${encodeURIComponent(args.options.displayName as string)}'`;

      this.GetServicePrincipalForApp(spMatchQuery)
      .then((resp: { value: ServicePrincipal[] }): Promise<ServicePrincipal>[] | Promise<any> => {

        if (!resp.value.length) {
          return Promise.reject('app registration not found');
        }

        sp = resp.value[0];

        // the role assignment has an appRoleId but no name. To get the name, we need to get all the roles from the resource.
        // the resource is a service principal. Multiple roles may have same resource id.
        const resourceIds = sp.appRoleAssignments.map((item: AppRoleAssignment) => item.resourceId);
        
        const tasks: Promise<ServicePrincipal>[] = [];
        for (let i: number = 0; i < resourceIds.length; i++) {
          tasks.push(this.GetServicePrincipal(resourceIds[i]));
        }

        return Promise.all(tasks);
      })
      .then((resources: ServicePrincipal[]) => {
        // loop through all appRoleAssignments for the servicePrincipal and lookup the appRole.Id in the resources[resourceId].appRoles array...
        const results: any[] = [];
        sp.appRoleAssignments.map((appRoleAssignment: AppRoleAssignment) => {
          const resource = resources.find((r: any) => r.objectId === appRoleAssignment.resourceId);
          if (resource) {
            const appRole = resource.appRoles.find((r: any) => r.id === appRoleAssignment.id);
            if (appRole) {
              results.push({
                appRoleId: appRoleAssignment.id,
                resourceDisplayName: appRoleAssignment.resourceDisplayName,
                resourceId: appRoleAssignment.resourceId,
                roleId: appRole.id,
                roleName: appRole.value
              });
            }
          }
        });

        if (args.options.output === 'json') {
          cmd.log(results);
        }
        else {
          cmd.log(results.map((r: any) => {
            return {
              resourceDisplayName: r.resourceDisplayName,
              roleName: r.roleName
            }
          }));
        }

        cb();
      }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
  }

  private GetServicePrincipalForApp(filterParam: string): Promise<{ value: ServicePrincipal[]}> {

    const spRequestOptions: any = {
      url: `${this.resource}/myorganization/servicePrincipals?api-version=1.6&$expand=appRoleAssignments&$filter=${filterParam}`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    return request.get<{ value: ServicePrincipal[]}>(spRequestOptions);
  }

  private GetServicePrincipal(spId: string): Promise<ServicePrincipal> {
    const spRequestOptions: any = {
      url: `${this.resource}/myorganization/servicePrincipals/${spId}?api-version=1.6`,
      headers: {
        accept: 'application/json'
      },
      json: true
    };

    return request.get<ServicePrincipal>(spRequestOptions);
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --appId [appId]',
        description: 'Application (client) Id of the App Registration for which the configured appRoles should be retrieved'
      },
      {
        option: '-n, --displayName [displayName]',
        description: 'Display name of the application for which the configured appRoles should be retrieved'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.appId && !args.options.displayName) {
        return 'Specify either appId or displayName';
      }

      if (args.options.appId && !Utils.isValidGuid(args.options.appId)) {
        return `${args.options.appId} is not a valid GUID`;
      }

      if (args.options.appId && args.options.displayName) {
        return 'Specify either appId or displayName but not both';
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.APPROLEASSIGNMENT_LIST).helpInformation());
    log(
      `  Remarks:
  
    Specify either the ${chalk.grey('appId')} or ${chalk.grey('displayName')} but not both. 
    If you specify both values, the command will fail with an error.
   
  Examples:
  
    List AppRoles assigned to service principal with Application (client) ID ${chalk.grey('b2307a39-e878-458b-bc90-03bc578531d6')}.
      ${commands.APPROLEASSIGNMENT_LIST} --appId b2307a39-e878-458b-bc90-03bc578531d6

    List AppRoles assigned to service principal with Application display name ${chalk.grey('MyAppName')}.
      ${commands.APPROLEASSIGNMENT_LIST} --displayName 'MyAppName'

  More information:
  
    Application and service principal objects in Azure Active Directory (Azure AD): 
    https://docs.microsoft.com/en-us/azure/active-directory/develop/active-directory-application-objects
`);
  }
}

module.exports = new AadAppRoleAssignmentListCommand();