import Command, { CommandAction, CommandError } from '../Command';
import appInsights from '../appInsights';
import * as path from 'path';
import * as fs from 'fs';
import * as os from 'os';
import {
  normalize,
  virtualFs,
} from '@angular-devkit/core';
import { NodeJsSyncHost } from '@angular-devkit/core/node';
import { DryRunEvent, UnsuccessfulWorkflowExecution } from '@angular-devkit/schematics';
import { NodeWorkflow } from '@angular-devkit/schematics/tools';

const vorpal: Vorpal = require('../vorpal-init');

export default abstract class SchematicsCommand extends Command {
  private _projectRootPath: string | null = null;

  protected abstract get schematic(): string;

  protected overwrite(): boolean {
    return false;
  }

  protected get projectRootPath(): string | null {
    return this._projectRootPath;
  }

  public action(): CommandAction {
    const cmd: SchematicsCommand = this;

    return function (this: CommandInstance, args: any, cb: () => void) {
      cmd._debug = args.options.debug || false;
      cmd._verbose = cmd._debug || args.options.verbose || false;

      appInsights.trackEvent({
        name: cmd.getCommandName(),
        properties: cmd.getTelemetryProperties(args)
      });
      appInsights.flush();

      const commandInstance = this;

      cmd._projectRootPath = cmd.getProjectRoot(process.cwd());
      if (cmd.projectRootPath === null) {
        this.log(new CommandError(`Couldn't find project root folder`));
        cb();
        return;
      }

      const fsHost = new virtualFs.ScopedHost(new NodeJsSyncHost(), normalize(cmd.projectRootPath));
      const workflow = new NodeWorkflow(fsHost, { force: cmd.overwrite() });

      try {
        cmd.commandAction(this, args, () => {
          workflow.reporter.subscribe((event: DryRunEvent) => {
            switch (event.kind) {
              case 'error':
                const desc = event.description == 'alreadyExist' ? 'already exists' : 'does not exist.';
                this.log(new CommandError(`${event.path} ${desc}.`));
                break;
              case 'update':
                this.log(`${vorpal.chalk.white('UPDATE')} ${event.path} (${event.content.length} bytes)`);
                break;
              case 'create':
                this.log(`${vorpal.chalk.green('CREATE')} ${event.path} (${event.content.length} bytes)`);
                break;
              case 'delete':
                this.log(`${vorpal.chalk.yellow('DELETE')} ${event.path}`);
                break;
              case 'rename':
                this.log(`${vorpal.chalk.blue('RENAME')} ${event.path} => ${event.to}`);
                break;
            }
          });

          try {
            workflow
              .execute({
                collection: path.resolve(__dirname, '../../'),
                schematic: cmd.schematic,
                options: args.options,
                debug: cmd._debug
              })
              .subscribe({
                error(err: Error) {
                  if (err instanceof UnsuccessfulWorkflowExecution) {
                    commandInstance.log(new CommandError('The Schematic workflow failed. See above.'));
                  }
                  else {
                    commandInstance.log(new CommandError(err.message));
                  }

                  cb();
                },
                complete() {
                  cmd.postSchematicAction(commandInstance, args, (): void => {
                    commandInstance.log(vorpal.chalk.green('DONE'));
                    cb();
                  });
                },
              });
          }
          catch (ex) {
            this.log(new CommandError(`An exception has occurred while executing the command: ${os.EOL}${ex}`));
            cb();
          }
        });
      }
      catch (ex) {
        this.log(new CommandError(ex));
        cb();
      }
    }
  }

  private getProjectRoot(folderPath: string): string | null {
    const packageJsonPath: string = path.resolve(folderPath, 'package.json');
    if (fs.existsSync(packageJsonPath)) {
      return folderPath;
    }
    else {
      const parentPath: string = path.resolve(folderPath, `..${path.sep}`);
      if (parentPath !== folderPath) {
        return this.getProjectRoot(parentPath);
      }
      else {
        return null;
      }
    }
  }

  public commandAction(cmd: CommandInstance, args: any, cb: () => void): void {
    cb();
  }

  protected postSchematicAction(cmd: CommandInstance, args: any, cb: () => void): void {
    cb();
  }
}