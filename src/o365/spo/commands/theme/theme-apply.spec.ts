import commands from '../../commands';
import Command, { CommandOption, CommandValidate, CommandError } from '../../../../Command';
import * as sinon from 'sinon';
const command: Command = require('./theme-apply');
import * as assert from 'assert';
import request from '../../../../request';
import Utils from '../../../../Utils';
import auth from '../../../../Auth';
import config from '../../../../config';

describe(commands.THEME_APPLY, () => {
  let vorpal: Vorpal;
  let log: string[];
  let cmdInstance: any;
  let cmdInstanceLogSpy: sinon.SinonSpy;

  before(() => {
    sinon.stub(auth, 'restoreAuth').callsFake(() => Promise.resolve());
    sinon.stub(command as any, 'getRequestDigest').callsFake(() => Promise.resolve({ FormDigestValue: 'ABC' }));
    auth.service.connected = true;
    auth.service.spoUrl = 'https://contoso.sharepoint.com';
  });

  beforeEach(() => {
    vorpal = require('../../../../vorpal-init');
    log = [];
    cmdInstance = {
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
      request.post
    ]);
  });

  after(() => {
    Utils.restore([
      auth.restoreAuth,
      (command as any).getRequestDigest
    ]);
    auth.service.connected = false;
    auth.service.spoUrl = undefined;
  });

  it('has correct name', () => {
    assert.equal(command.name.startsWith(commands.THEME_APPLY), true);
  });

  it('has a description', () => {
    assert.notEqual(command.description, null);
  });

  it('applies theme when correct parameters are passed', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: false,
        name: 'Contoso',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    }, () => {
      try {
        assert.equal(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.equal(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.equal(postStub.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="SetWebTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">Contoso</Parameter><Parameter Type="String">https://contoso.sharepoint.com/sites/project-x</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`);
        assert.equal(cmdInstanceLogSpy.notCalled, true);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('applies theme when correct parameters are passed (debug)', (done) => {
    const postStub: sinon.SinonStub = sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "SchemaVersion": "15.0.0.0", "LibraryVersion": "16.0.7025.1207", "ErrorInfo": null, "TraceCorrelationId": "3d92299e-e019-4000-c866-de7d45aa9628" }, 12, true]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    }, () => {
      try {
        assert.equal(postStub.lastCall.args[0].url, 'https://contoso-admin.sharepoint.com/_vti_bin/client.svc/ProcessQuery');
        assert.equal(postStub.lastCall.args[0].headers['X-RequestDigest'], 'ABC');
        assert.equal(postStub.lastCall.args[0].body, `<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="${config.applicationName}" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009"><Actions><ObjectPath Id="10" ObjectPathId="9" /><Method Name="SetWebTheme" Id="11" ObjectPathId="9"><Parameters><Parameter Type="String">Contoso</Parameter><Parameter Type="String">https://contoso.sharepoint.com/sites/project-x</Parameter></Parameters></Method></Actions><ObjectPaths><Constructor Id="9" TypeId="{268004ae-ef6b-4e9b-8425-127220d84719}" /></ObjectPaths></Request>`);
        assert.notEqual(cmdInstanceLogSpy.lastCall.args[0].indexOf('DONE'), -1);
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles error command error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "requestObjectIdentity ClientSvc error" } }]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        webUrl: 'https://contoso.sharepoint.com/sites/project-x'
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('requestObjectIdentity ClientSvc error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('handles unknown error command error correctly', (done) => {
    sinon.stub(request, 'post').callsFake((opts) => {
      if (opts.url.indexOf(`/_vti_bin/client.svc/ProcessQuery`) > -1) {
        return Promise.resolve(JSON.stringify([{ "ErrorInfo": { "ErrorMessage": "" } }]));
      }
      return Promise.reject('Invalid request');
    });

    cmdInstance.action({
      options: {
        debug: true,
        name: 'Contoso',
        filePath: 'theme.json',
        inverted: false,
      }
    }, (err?: any) => {
      try {
        assert.equal(JSON.stringify(err), JSON.stringify(new CommandError('ClientSvc unknown error')));
        done();
      }
      catch (e) {
        done(e);
      }
    });
  });

  it('supports debug mode', () => {
    const options = (command.options() as CommandOption[]);
    let containsDebugOption = false;
    options.forEach(o => {
      if (o.option === '--debug') {
        containsDebugOption = true;
      }
    });
    assert(containsDebugOption);
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
    assert(find.calledWith(commands.THEME_APPLY));
  });

  it('fails validation if name is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: '', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when name is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Contoso-Blue', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.equal(actual, true);
  });

  it('fails validation if webUrl is not passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Contoso-Blue', webUrl: '' } });
    assert.notEqual(actual, true);
  });

  it('fails validation if webUrl is not a valid SharePoint URL', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Contoso-Blue', webUrl: 'invalid' } });
    assert.notEqual(actual, true);
  });

  it('passes validation when webUrl is passed', () => {
    const actual = (command.validate() as CommandValidate)({ options: { name: 'Contoso-Blue', webUrl: 'https://contoso.sharepoint.com/sites/project-x' } });
    assert.equal(actual, true);
  });
});