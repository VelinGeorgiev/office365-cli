import SpoCommand from "./SpoCommand";
import { IdentityResponse } from "./common/IdentityResponse";
import Utils from "../../Utils";
import { ClientSvcResponseContents, ClientSvcResponse } from "./spo";
import * as request from 'request-promise-native';
import config from "../../config";

export default abstract class SpoClientSvcCommand extends SpoCommand {
  
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
   * @param cmd command cmd
   */
  protected requestObjectIdentity(webUrl: string, siteAccessToken: string, formDigestValue: string, cmd: CommandInstance): Promise<IdentityResponse> {
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

}