import * as request from 'request-promise-native';
import Utils from '../../../Utils';
import config from '../../../config';
import { ClientSvcResponse, ClientSvcResponseContents } from '../spo';
import { BasePermissions } from './base-permissions';

export interface IdentityResponse {
  objectIdentity: string;
  serverRelativeUrl: string;
};

export class ClientSvcCommons {

  public readonly cmd: CommandInstance;
  public readonly debug: boolean;

  constructor(cmd: CommandInstance, debug: boolean) {
    this.cmd = cmd;
    this.debug = debug;
  }

  /**
   * Requests web object itentity for the current web.
   * This request has to be send before we can construct the property bag request.
   * The response data looks like:
   * _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
   * _ObjectType_=SP.Web
   * ServerRelativeUrl=/sites/contoso
   * The ObjectIdentity is needed to create another request to retrieve the property bag or set property.
   * @param webUrl web url
   * @param siteAccessToken site access token
   * @param formDigestValue formDigestValue
   */
  public requestObjectIdentity(webUrl: string, siteAccessToken: string, formDigestValue: string): Promise<IdentityResponse> {
    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${siteAccessToken}`,
        'X-RequestDigest': formDigestValue
      }),
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="1" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="ServerRelativeUrl" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Property Id="5" ParentId="3" Name="Web" /><StaticProperty Id="3" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /></ObjectPaths></Request>`
    };

    return new Promise<IdentityResponse>((resolve: any, reject: any): void => {
      request.post(requestOptions).then((res: any) => {

        const json: ClientSvcResponse = JSON.parse(res);

        const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
        if (contents && contents.ErrorInfo) {
          return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }

        const identityObject = json.find(x => { return x['_ObjectIdentity_'] });
        if (identityObject) {
          return resolve(
            {
              objectIdentity: identityObject['_ObjectIdentity_'],
              serverRelativeUrl: identityObject['ServerRelativeUrl']
            });
        }

        reject('Cannot proceed. _ObjectIdentity_ not found'); // this is not supposed to happen
      }, (err: any): void => { reject(err); });
    });
  }

  /**
   * Gets EffectiveBasePermissions for web return type is "_ObjectType_\":\"SP.Web\".
   * @param webObjectIdentity ObjectIdentity. Looks like _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>
   * @param webUrl web url
   * @param siteAccessToken site access token
   * @param formDigestValue formDigestValue
   */
  public getEffectiveBasePermissions(webObjectIdentity: string, webUrl: string, siteAccessToken: string, formDigestValue: string): Promise<BasePermissions> {

    let basePermissionsResult = new BasePermissions();

    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${siteAccessToken}`,
        'X-RequestDigest': formDigestValue
      }),
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><Query Id="11" ObjectPathId="5"><Query SelectAllProperties="false"><Properties><Property Name="EffectiveBasePermissions" ScalarProperty="true" /></Properties></Query></Query></Actions><ObjectPaths><Identity Id="5" Name="${webObjectIdentity}" /></ObjectPaths></Request>`
    };

    if (this.debug) {
      this.cmd.log('Request:');
      this.cmd.log(JSON.stringify(requestOptions));
      this.cmd.log('');
    }

    return new Promise<BasePermissions>((resolve: any, reject: any): void => {
      request.post(requestOptions).then((res: any) => {

        if (this.debug) {
          this.cmd.log('Response:');
          this.cmd.log(JSON.stringify(res));
          this.cmd.log('');

          this.cmd.log('Attempt to get the web EffectiveBasePermissions');
        }

        const json: ClientSvcResponse = JSON.parse(res);
        const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
        if (contents && contents.ErrorInfo) {
          return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }

        const permissionsObj = json.find(x => { return x['EffectiveBasePermissions'] });
        if (permissionsObj) {
          basePermissionsResult.high = permissionsObj['EffectiveBasePermissions']['High'];
          basePermissionsResult.low = permissionsObj['EffectiveBasePermissions']['Low'];
          return resolve(basePermissionsResult);
        }

        reject('Cannot proceed. EffectiveBasePermissions not found'); // this is not supposed to happen
      }, (err: any): void => { reject(err); })
    });
  }

  /**
    * Gets folder by server relative url (GetFolderByServerRelativeUrl is REST)
    * The response data looks like:
    * _ObjectIdentity_=<GUID>|<GUID>:site:<GUID>:web:<GUID>:folder:<GUID>
    * _ObjectType_=SP.Folder
    * @param identityResp IdentityResponse
    * @param webUrl web url
    * @param siteAccessToken site access token
    * @param formDigestValue formDigestValue
    */
  public requestFolderObjectIdentity(identityResp: IdentityResponse, webUrl: string, folder: string, siteAccessToken: string, formDigestValue: string): Promise<IdentityResponse> {
    let serverRelativeUrl: string = folder;
    if (identityResp.serverRelativeUrl !== '/') {
      serverRelativeUrl = `${identityResp.serverRelativeUrl}${serverRelativeUrl}`
    }

    const requestOptions: any = {
      url: `${webUrl}/_vti_bin/client.svc/ProcessQuery`,
      headers: Utils.getRequestHeaders({
        authorization: `Bearer ${siteAccessToken}`,
        'X-RequestDigest': formDigestValue
      }),
      body: `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><ObjectIdentityQuery Id="11" ObjectPathId="9" /><Query Id="12" ObjectPathId="9"><Query SelectAllProperties="false"><Properties><Property Name="Properties" SelectAll="true"><Query SelectAllProperties="false"><Properties /></Query></Property></Properties></Query></Query></Actions><ObjectPaths><Method Id="9" ParentId="5" Name="GetFolderByServerRelativeUrl"><Parameters><Parameter Type="String">${serverRelativeUrl}</Parameter></Parameters></Method><Identity Id="5" Name="${identityResp.objectIdentity}" /></ObjectPaths></Request>`
    };

    if (this.debug) {
      this.cmd.log('Request:');
      this.cmd.log(JSON.stringify(requestOptions));
      this.cmd.log('');
    }

    return new Promise<IdentityResponse>((resolve: any, reject: any) => {

      return request.post(requestOptions).then((res: any) => {
        if (this.debug) {
          this.cmd.log('Response:');
          this.cmd.log(JSON.stringify(res));
          this.cmd.log('');
        }

        const json: ClientSvcResponse = JSON.parse(res);

        const contents: ClientSvcResponseContents = json.find(x => { return x['ErrorInfo']; });
        if (contents && contents.ErrorInfo) {
          return reject(contents.ErrorInfo.ErrorMessage || 'ClientSvc unknown error');
        }
        const objectIdentity = json.find(x => { return x['_ObjectIdentity_'] });
        if (objectIdentity) {
          return resolve({
            objectIdentity: objectIdentity['_ObjectIdentity_'],
            serverRelativeUrl: serverRelativeUrl
          });
        }

        reject('Cannot proceed. Folder _ObjectIdentity_ not found'); // this is not suppose to happen

      }, (err: any): void => { reject(err); })
    });
  }
}