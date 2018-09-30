// #!/usr/bin/env node
import { performance } from 'perf_hooks';
performance.mark('ttc_start');

import * as fs from 'fs';
import * as path from 'path';
import * as updateNotifier from 'update-notifier';
import config from './config';
// import Command from './Command';
import appInsights from './appInsights';
import Utils from './Utils';
import { autocomplete } from './autocomplete';

import cmd from './commands_index';

const packageJSON = require('../package.json');
const vorpal: Vorpal = require('./vorpal-init'),
  chalk = vorpal.chalk;

const readdirR = (dir: string): string | string[] => {
  return fs.statSync(dir).isDirectory()
    ? Array.prototype.concat(...fs.readdirSync(dir).map(f => readdirR(path.join(dir, f))))
    : dir;
}

appInsights.trackEvent({
  name: 'started'
});

updateNotifier({ pkg: packageJSON }).notify({ defer: false });

fs.realpath(__dirname, (err: NodeJS.ErrnoException, resolvedPath: string): void => {
  // const commandsDir: string = path.join(resolvedPath, './o365');
  performance.mark('start');
  // const files: string[] = readdirR(commandsDir) as string[];
  performance.mark('end');
  // performance.measure('discovering files', 'start', 'end');

  performance.mark('start');
  (cmd as any).forEach((c: any) => {
    c.init(vorpal);
  });
  // files.forEach(file => {
  //   if (file.indexOf(`${path.sep}commands${path.sep}`) > -1 &&
  //     file.indexOf('.spec.js') === -1 &&
  //     file.indexOf('.js.map') === -1) {
  //     try {
  //       const cmd: any = require(file);
  //       if (cmd instanceof Command) {
  //         cmd.init(vorpal);
  //       }
  //     }
  //     catch { }
  //   }
  // });

  performance.mark('end');
  performance.measure('loading commands', 'start', 'end')

  if (process.argv.indexOf('--completion:clink:generate') > -1) {
    console.log(autocomplete.getClinkCompletion(vorpal));
    process.exit();
  }
  if (process.argv.indexOf('--completion:sh:generate') > -1) {
    autocomplete.generateShCompletion(vorpal);
    process.exit();
  }
  if (process.argv.indexOf('--completion:sh:setup') > -1) {
    autocomplete.generateShCompletion(vorpal);
    autocomplete.setupShCompletion();
    process.exit();
  }
  if (process.argv.indexOf('--reconsent') > -1) {
    console.log(`To reconsent the PnP Office 365 Management Shell Azure AD application navigate in your web browser to https://login.microsoftonline.com/common/oauth2/authorize?client_id=${config.cliAadAppId}&response_type=code&prompt=admin_consent`);
    process.exit();
  }

  // disable linux-normalizing args to support JSON and XML values
  vorpal.isCommandArgKeyPairNormalized = false;

  vorpal
    .title('Office 365 CLI')
    .description(packageJSON.description)
    .version(packageJSON.version);

  vorpal
    .command('version', 'Shows the current version of the CLI')
    .action(function (this: CommandInstance, args: any, cb: () => void) {
      this.log(packageJSON.version);
      cb();
    });

  vorpal.pipe((stdout: any): any => {
    return Utils.logOutput(stdout);
  });

  let v: Vorpal | null = null;
  try {
    if (process.argv.length > 2) {
      vorpal.delimiter('');
      vorpal.on('client_command_error', (err?: any): void => {
        if (v) {
          process.exit(1);
        }
      });
    }
    performance.mark('ttc_end');
  performance.measure('ttc', 'ttc_start', 'ttc_end');
    v = vorpal.parse(process.argv);

    // if no command has been passed/match, run immersive mode
    if (!v._command) {
      vorpal
        .delimiter(chalk.red(config.delimiter + ' '))
        .show();
    }

console.log(performance.getEntriesByType('measure'));

  }
  catch (e) {
    appInsights.trackException({
      exception: e
    });
    appInsights.flush();
    process.exit(1);
  }
});