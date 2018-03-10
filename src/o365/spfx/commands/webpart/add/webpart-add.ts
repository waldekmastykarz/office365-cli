import commands from '../../../commands';
import GlobalOptions from '../../../../../GlobalOptions';
import {
  CommandOption,
  CommandValidate
} from '../../../../../Command';
import SchematicsCommand from '../../../../../schematics/SchematicsCommand';
import * as fs from 'fs';
import * as path from 'path';
import Utils from '../../../../../Utils';

const vorpal: Vorpal = require('../../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  name: string;
}

class SpfxWebPartAddCommand extends SchematicsCommand {
  public get name(): string {
    return `${commands.WEBPART_ADD}`;
  }

  public get description(): string {
    return 'Adds web part to an existing project';
  }

  public get schematic(): string {
    return 'webpart';
  }

  // public getTelemetryProperties(args: CommandArgs): any {
  //   const telemetryProps: any = super.getTelemetryProperties(args);
  //   telemetryProps.overwrite = args.options.overwrite || false;
  //   return telemetryProps;
  // }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-n, --name <name>',
        description: 'Web part name'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.name) {
        return 'Missing required option name';
      }

      return true;
    };
  }

  public commandHelp(args: CommandArgs, log: (message: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.WEBPART_ADD).helpInformation());
//     log(
//       `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
//   using the ${chalk.blue(commands.CONNECT)} command.
                
//   Remarks:

//     To add an app to the tenant app catalog, you have to first connect to a SharePoint site using the
//     ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

//     When specifying the path to the app package file you can use both relative and absolute paths.
//     Note, that ~ in the path, will not be resolved and will most likely result in an error.

//     If you try to upload a package that already exists in the tenant app catalog without specifying
//     the ${chalk.blue('--overwrite')} option, the command will fail with an error stating that the
//     specified package already exists.

//   Examples:
  
//     Add the ${chalk.grey('spfx.sppkg')} package to the tenant app catalog
//       ${chalk.grey(config.delimiter)} ${commands.APP_ADD} -p /Users/pnp/spfx/sharepoint/solution/spfx.sppkg

//     Overwrite the ${chalk.grey('spfx.sppkg')} package in the tenant app catalog with the newer version
//       ${chalk.grey(config.delimiter)} ${commands.APP_ADD} -p sharepoint/solution/spfx.sppkg --overwrite

//   More information:

//     Application Lifecycle Management (ALM) APIs
//       https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
// `);
  }
}

module.exports = new SpfxWebPartAddCommand();