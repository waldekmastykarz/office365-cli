import * as chalk from 'chalk';
import * as fs from 'fs';
import auth, { AuthType, CloudType } from '../../Auth';
import { Logger } from '../../cli';
import Command, {
  CommandError, CommandOption
} from '../../Command';
import GlobalOptions from '../../GlobalOptions';
import Utils from '../../Utils';
import commands from './commands';

interface CommandArgs {
  options: Options;
}

interface Options extends GlobalOptions {
  authType?: string;
  certificateFile?: string;
  cloud?: string;
  password?: string;
  thumbprint?: string;
  userName?: string;
}

class LoginCommand extends Command {
  public get name(): string {
    return `${commands.LOGIN}`;
  }

  public get description(): string {
    return 'Log in to Microsoft 365';
  }

  public getTelemetryProperties(args: CommandArgs): any {
    const telemetryProps: any = super.getTelemetryProperties(args);
    telemetryProps.authType = args.options.authType || 'deviceCode';
    telemetryProps.certificateFile = typeof args.options.certificateFile !== 'undefined';
    telemetryProps.cloud = typeof args.options.cloud !== 'undefined';
    telemetryProps.password = typeof args.options.password !== 'undefined';
    telemetryProps.thumbprint = typeof args.options.thumbprint !== 'undefined';
    telemetryProps.userName = typeof args.options.userName !== 'undefined';
    return telemetryProps;
  }

  public commandAction(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    // disconnect before re-connecting
    if (this.debug) {
      logger.log(`Logging out from Microsoft 365...`);
    }

    const logout: () => void = (): void => {
      auth.service.logout();
      if (this.verbose) {
        logger.log(chalk.green('DONE'));
      }
    }

    const login: () => void = (): void => {
      if (this.verbose) {
        logger.log(`Signing in to Microsoft 365...`);
      }

      switch (args.options.authType) {
        case 'password':
          auth.service.authType = AuthType.Password;
          auth.service.userName = args.options.userName;
          auth.service.password = args.options.password;
          break;
        case 'certificate':
          auth.service.authType = AuthType.Certificate;
          auth.service.certificate = fs.readFileSync(args.options.certificateFile as string, 'base64');
          auth.service.thumbprint = args.options.thumbprint;
          auth.service.password = args.options.password;
          break;
        case 'identity':
          auth.service.authType = AuthType.Identity;
          auth.service.userName = args.options.userName;
          break;
      }

      if (args.options.cloud) {
        auth.service.cloudType = CloudType[args.options.cloud as keyof typeof CloudType];
      }
      else {
        auth.service.cloudType = CloudType.AzureCloud;
      }
      auth.setAuthContext();

      auth
        .ensureAccessToken(auth.defaultResource, logger, this.debug)
        .then((): void => {
          if (this.verbose) {
            logger.log(chalk.green('DONE'));
          }

          auth.service.connected = true;
          cb();
        }, (rej: string): void => {
          if (this.debug) {
            logger.log('Error:');
            logger.log(rej);
            logger.log('');
          }

          if (rej !== 'Polling_Request_Cancelled') {
            cb(new CommandError(rej));
            return;
          }
          cb();
        });
    }

    auth
      .clearConnectionInfo()
      .then((): void => {
        logout();
        login();
      }, (error: any): void => {
        if (this.debug) {
          logger.log(new CommandError(error));
        }

        logout();
        login();
      });
  }

  public action(logger: Logger, args: CommandArgs, cb: (err?: any) => void): void {
    auth
      .restoreAuth()
      .then((): void => {
        this.initAction(args, logger);
        this.commandAction(logger, args, cb);
      }, (error: any): void => {
        cb(new CommandError(error));
      });
  }

  public options(): CommandOption[] {
    const options: CommandOption[] = [
      {
        option: '-t, --authType [authType]',
        description: 'The type of authentication to use. Allowed values certificate|deviceCode|password|identity. Default deviceCode',
        autocomplete: ['certificate', 'deviceCode', 'password', 'identity']
      },
      {
        option: '-u, --userName [userName]',
        description: 'Name of the user to authenticate. Required when authType is set to password'
      },
      {
        option: '-p, --password [password]',
        description: 'Password for the user. Required when authType is set to password'
      },
      {
        option: '-c, --certificateFile [certificateFile]',
        description: 'Path to the file with certificate private key. Required when authType is set to certificate'
      },
      {
        option: '--thumbprint [thumbprint]',
        description: 'Certificate thumbprint. Required when authType is set to certificate'
      },
      {
        option: '--cloud [cloud]',
        description: `Type of cloud to connect to. Default AzureCloud. Allowed values ${Utils.getEnums(CloudType).join(',')}`,
        autocomplete: Utils.getEnums(CloudType)
      }
    ];

    const parentOptions: CommandOption[] = super.options();
    return options.concat(parentOptions);
  }

  public validate(args: CommandArgs): boolean | string {
    if (args.options.authType === 'password') {
      if (!args.options.userName) {
        return 'Required option userName missing';
      }

      if (!args.options.password) {
        return 'Required option password missing';
      }
    }

    if (args.options.authType === 'certificate') {
      if (!args.options.certificateFile) {
        return 'Required option certificateFile missing';
      }

      if (!fs.existsSync(args.options.certificateFile)) {
        return `File '${args.options.certificateFile}' does not exist`;
      }

      if (!args.options.thumbprint) {
        return 'Required option thumbprint missing';
      }
    }

    if (args.options.cloud &&
      typeof CloudType[args.options.cloud as keyof typeof CloudType] === 'undefined') {
      return `${args.options.cloud} is not a valid value for cloud. Valid options are ${Utils.getEnums(CloudType).join(',')}`;
    }

    return true;
  }
}

module.exports = new LoginCommand();