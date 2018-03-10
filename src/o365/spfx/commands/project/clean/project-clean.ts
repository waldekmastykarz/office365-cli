import commands from '../../../commands';
import GlobalOptions from '../../../../../GlobalOptions';
import SchematicsCommand from '../../../../../schematics/SchematicsCommand';

const vorpal: Vorpal = require('../../../../../vorpal-init');

interface CommandArgs {
  options: GlobalOptions;
}

class SpfxProjectCleanCommand extends SchematicsCommand {
  public get name(): string {
    return `${commands.PROJECT_CLEAN}`;
  }

  public get description(): string {
    return 'Removes unnecessary files from the project';
  }

  public get schematic(): string {
    return 'spfx-project-clean';
  }

  public commandHelp(args: CommandArgs, log: (message: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.getCommandName()).helpInformation());
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

module.exports = new SpfxProjectCleanCommand();