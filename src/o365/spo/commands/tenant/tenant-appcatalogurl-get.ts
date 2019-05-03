import request from '../../../../request';
import config from '../../../../config';
import commands from '../../commands';
import SpoCommand from '../../SpoCommand';
const vorpal: Vorpal = require('../../../../vorpal-init');

class SpoTenantAppCatalogUrlGetCommand extends SpoCommand {
  public get name(): string {
    return commands.TENANT_APPCATALOGURL_GET;
  }

  public get description(): string {
    return 'Gets the URL of the tenant app catalog';
  }

  public commandAction(cmd: CommandInstance, args: any, cb: (err?: any) => void): void {
    this
      .getSpoAdminUrl(cmd, this.debug)
      .then((spoAdminUrl: string): Promise<string> => {
        const requestOptions: any = {
          url: `${spoAdminUrl}/_api/SP_TenantSettings_Current`,
          headers: {
            accept: 'application/json;odata=nometadata'
          }
        };
    
        return request.get(requestOptions);
      })
      .then((res: string): void => {
        const json = JSON.parse(res);

        if (json.CorporateCatalogUrl) {
          cmd.log(json.CorporateCatalogUrl);
        }
        else {
          if (this.verbose) {
            cmd.log("Tenant app catalog is not configured.");
          }
        }
        cb();
      }, (err: any): void => this.handleRejectedPromise(err, cmd, cb));
  }

  public commandHelp(args: {}, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.name).helpInformation());
    log(
      `  ${chalk.yellow('Important:')} to use this command you have to have permissions to access
    the tenant admin site.
    
  Examples:
  
    Get the URL of the tenant app catalog
      ${chalk.grey(config.delimiter)} ${commands.TENANT_APPCATALOGURL_GET}
  ` );
  }
}

module.exports = new SpoTenantAppCatalogUrlGetCommand();