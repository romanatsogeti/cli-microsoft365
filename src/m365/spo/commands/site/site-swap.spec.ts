import commands from '../../commands';
import Command from '../../../../Command';
import config from '../../../../config';
import * as sinon from 'sinon';
import appInsights from '../../../../appInsights';
import auth from '../../../../Auth';
const command: Command = require('./site-swap');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';


describe(commands.SITE_SWAP, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(appInsights, 'trackEvent').callsFake(() => { });
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
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

    cmdInstanceLogSpy = sinon.spy(cmdInstance, 'log');
  });

  afterEach(() => {
    Utils.restore([
      vorpal.find,
      request.post,
      request.get,
      Utils.executeCommand,
      (command as any).getSpoAdminUrl
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest,
      appInsights.trackEvent
    ]);
    auth.service.connected = false;
  });

  it('has correct name', () => {
    assert.strictEqual(command.name.startsWith(commands.SITE_SWAP), true);
  });

  it('has a description', () => {
    assert.notStrictEqual(command.description, null);
  });

  it('cannot swap sites if target url is not root site or the search center', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="SwapSiteWithSmartGestureOptionForce"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sourcesite</Parameter><Parameter Type="String">https://contoso.sharepoint.com/sites/not-a-root-site</Parameter><Parameter Type="String">https://contoso.sharepoint.com/sites/root-archive</Parameter><Parameter Type="Boolean">true</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.20502.1",
            "ErrorInfo": {
              "ErrorMessage" : "The target URL must be the root site or the search center site",
              "ErrorValue": null,
              "ErrorTypeName": "Microsoft.Online.SharePoint.Common.SpoException",
              "TraceCorrelationId": "89d9799f-300c-0000-54e8-79e858d1224b"
            },
            "TraceCorrelationId": "89d9799f-300c-0000-54e8-79e858d1224b"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false, 
        sourceUrl: 'https://contoso.sharepoint.com/sites/sourcesite',
        targetUrl: 'https://contoso.sharepoint.com/not-a-root-site',
        archiveUrl: 'https://contoso.sharepoint.com/sites/root-archive'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.calledWith('target url is not root site or search center.'));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('swaps the sites while archiving the original site to the archival url', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if ((opts.url === 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery') &&
        opts.body === `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="4" ObjectPathId="3" /><ObjectPath Id="6" ObjectPathId="5" /><Query Id="7" ObjectPathId="5"><Query SelectAllProperties="true"><Properties /></Query></Query></Actions><ObjectPaths><Constructor Id="3" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /><Method Id="5" ParentId="3" Name="SwapSiteWithSmartGestureOptionForce"><Parameters><Parameter Type="String">https://contoso.sharepoint.com/sites/sourcesite</Parameter><Parameter Type="String">https://contoso.sharepoint.com/</Parameter><Parameter Type="String">https://contoso.sharepoint.com/sites/root-archive</Parameter><Parameter Type="Boolean">true</Parameter><Parameter Type="Boolean">false</Parameter></Parameters></Method></ObjectPaths></Request>`) {
        return Promise.resolve(JSON.stringify([
          {
            "SchemaVersion": "15.0.0.0",
            "LibraryVersion": "16.0.20502.1",
            "ErrorInfo": null,
            "TraceCorrelationId": "f10a459e-409f-4000-c5b4-09fb5e795218"
          }
        ]));
      }

      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false, 
        sourceUrl: 'https://contoso.sharepoint.com/sites/sourcesite',
        targetUrl: 'https://contoso.sharepoint.com/',
        archiveUrl: 'https://contoso.sharepoint.com/sites/root-archive'
      }
    }, () => {
      try {
        assert(cmdInstanceLogSpy.notCalled);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });
});
