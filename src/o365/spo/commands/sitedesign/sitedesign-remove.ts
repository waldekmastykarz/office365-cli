import config from '../../../../config';
import request from '../../../../request';
import commands from '../../commands';
import {
  CommandOption, CommandValidate
} from '../../../../Command';
import SpoCommand from '../../SpoCommand';
import Utils from '../../../../Utils';
import { ContextInfo } from '../../spo';
import GlobalOptions from '../../../../GlobalOptions';

const vorpal: Vorpal = require('../../../../vorpal-init');

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  id: string;
  confirm?: boolean;
}

class SpoSiteDesignRemoveCommand extends SpoCommand {
  public get name(): string {
    return `${commands.SITEDESIGN_REMOVE}`;
  }

  public get description(): string {
    return 'Removes the specified site design';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.confirm = args.options.confirm || false;
    return telemetryProps;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const removeSiteDesign: () => void = (): void => {
      let spoAdminUrl: string = '';

      this
        .getSpoAdminUrl(cmd, this.debug)
        .then((_spoAdminUrl: string): Promise<ContextInfo> => {
          spoAdminUrl = _spoAdminUrl;
          return this.getRequestDigest(spoAdminUrl);
        })
        .then((res: ContextInfo): Promise<void> => {
          const requestOptions: any = {
            url: `${spoAdminUrl}/_api/Microsoft.Sharepoint.Utilities.WebTemplateExtensions.SiteScriptUtility.DeleteSiteDesign`,
            headers: {
              'X-RequestDigest': res.FormDigestValue,
              'content-type': 'application/json;charset=utf-8',
              accept: 'application/json;odata=nometadata'
            },
            body: { id: args.options.id },
            json: true
          };

          return request.post(requestOptions);
        })
        .then((): void => {
          if (this.verbose) {
            cmd.log(vorpal.chalk.green('DONE'));
          }

          cb();
        }, (err: any): void => this.handleRejectedODataJsonPromise(err, cmd, cb));
    };

    if (args.options.confirm) {
      removeSiteDesign();
    }
    else {
      cmd.prompt({
        type: 'confirm',
        name: 'continue',
        default: false,
        message: `Are you sure you want to remove the site design ${args.options.id}?`,
      }, (result: { continue: boolean }): void => {
        if (!result.continue) {
          cb();
        }
        else {
          removeSiteDesign();
        }
      });
    }
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-i, --id <id>',
        description: 'Site design ID'
      },
      {
        option: '--confirm',
        description: 'Don\'t prompt for confirming removing the site design'
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(): CommandValidate {
    return (args: CommandArgs): boolean | string => {
      if (!args.options.id) {
        return 'Required parameter id missing';
      }

      if (!Utils.isValidGuid(args.options.id)) {
        return `${args.options.id} is not a valid GUID`;
      }

      return true;
    };
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  Remarks:

    If the specified ${chalk.grey('id')} doesn't refer to an existing site design, you will get
    a ${chalk.grey('File not found')} error.

  Examples:
  
    Remove site design with ID ${chalk.grey('2c1ba4c4-cd9b-4417-832f-92a34bc34b2a')}. Will prompt
    for confirmation before removing the design
      ${chalk.grey(config.delimiter)} ${this.name} --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a

    Remove site design with ID ${chalk.grey('2c1ba4c4-cd9b-4417-832f-92a34bc34b2a')} without prompting
    for confirmation
      ${chalk.grey(config.delimiter)} ${this.name} --id 2c1ba4c4-cd9b-4417-832f-92a34bc34b2a --confirm

  More information:

    SharePoint site design and site script overview
      https://docs.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-overview
`);
  }
}

module.exports = new SpoSiteDesignRemoveCommand();