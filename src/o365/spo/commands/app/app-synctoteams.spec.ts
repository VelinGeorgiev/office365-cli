import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./app-synctoteams');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';

describe(commands.APP_SYNCTOTEAMS, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => {});
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
      commandWrapper: {
        command: command.name
      },
      action: command.action(),
      log: (msg: string) => {
        log.push(msg);
      }
    };
    sinon.stub(request, 'get').resolves({ "CorporateCatalogUrl": "https://contoso.sharepoint.com/sites/apps" });
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.get,
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.APP_SYNCTOTEAMS), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('syncs app from tenant app catalog to Microsoft Teams (debug)', (done) => {
    const postRequestStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/tenantappcatalog/SyncSolutionToTeams`) > -1) {
          return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: true, id: 1, confirm: true } }, () => {
      try {
        assert.notEqual(postRequestStub.lastCall.args[0].url.indexOf('/_api/web/tenantappcatalog/SyncSolutionToTeams(id=1)'), -1);
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('syncs app from tenant app catalog to Microsoft Teams', (done) => {
    const postRequestStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/tenantappcatalog/SyncSolutionToTeams`) > -1) {
          return Promise.resolve();
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: 1, confirm: true } }, () => {
      try {
        assert.notEqual(postRequestStub.lastCall.args[0].url.indexOf('/_api/web/tenantappcatalog/SyncSolutionToTeams(id=1)'), -1);
        done();
      } catch (e) {
        done(e);
      }
    });
  });

  it('correctly handles failure when canot sync the app', (done) => {
      sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_api/web/tenantappcatalog`) > -1) {
          return Promise.reject({
            error: JSON.stringify({
              'odata.error': {
                code: '-1, Microsoft.SharePoint.Client.ResourceNotFoundException',
                message: {
                  lang: "en-US",
                  value: "Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown."
                }
              }
            })
          });
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({ options: { debug: false, id: 1 } }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError("Exception of type 'Microsoft.SharePoint.Client.ResourceNotFoundException' was thrown.")));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('fails validation if the id option not specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: {} });
    assert.notEqual(actual, true);
  });

  it('passes validation when the id is specified', () => {
    const actual = (command.validate() as CommandValidate)({ options: { id: 1 } });
    assert.equal(actual, true);
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsdebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsdebugOption = true;
      }
    });
    assert(containsdebugOption);
  });

  it('has help referring to the right command', () => {
    const cmd: any = {
      log: (msg: string) => { },
      prompt: () => { },
      helpInformation: () => { }
    };
    const find = sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    assert(find.calledWith(commands.APP_SYNCTOTEAMS));
  });

  it('has help with examples', () => {
    const _log: string[] = [];
    const cmd: any = {
      log: (msg: string) => {
        _log.push(msg);
      },
      prompt: () => { },
      helpInformation: () => { }
    };
    sinon.stub(vorpal, 'find').callsFake(() => cmd);
    cmd.help = command.help();
    cmd.help({}, () => { });
    let containsExamples: boolean = false;
    _log.forEach(l => {
      if (l && l.indexOf('Examples:') > -1) {
        containsExamples = true;
      }
    });
    Utils.restore(vorpal.find);
    assert(containsExamples);
  });
});