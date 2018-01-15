import config from '../../../config';
import commands from '../../commands';
import Command, {
  CommandError
} from '../../../Command';
import * as fs from 'fs';
import * as path from 'path';

const vorpal: Vorpal = require('../../../vorpal-init');

class ProvisionCommand extends Command {
  private static mappings: any;
  private static mappingsKeys: string[];
  private numErrors: number;
  private errors: any[];
  private queue: any[];
  private cb: () => void;
  private runFirst: boolean;

  public get name(): string {
    return commands.PROVISION;
  }

  public get description(): string {
    return 'Provision Office 365 configuration from the specified template';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    if (!ProvisionCommand.mappings) {
      ProvisionCommand.mappings = JSON.parse(fs.readFileSync(path.join(__dirname, '../../../../src/o365/commands/provision/mappings.json'), 'utf8'));
      ProvisionCommand.mappingsKeys = Object.keys(ProvisionCommand.mappings).sort((a, b) => b.length - a.length);
    }

    const templatePath: string = path.join(__dirname, '../../../../src/o365/commands/provision/test.json');
    const templateRaw: string = fs.readFileSync(templatePath, 'utf8');
    const template: any = JSON.parse(templateRaw);
    this.queue = [];
    this.errors = [];
    this.numErrors = 0;
    this.cb = cb;
    this.runFirst = false;

    try {
      this.provision(template.Provisioning, { _path: 'Provisioning' }, cmd);
    }
    catch (e) {
      cmd.log(vorpal.chalk.red(e));
    }
  }

  private provision(obj: any, context: any, cmd: CommandInstance): void {
    const mappingKey: string | undefined = ProvisionCommand.getMappingKey(context._path);
    if (mappingKey) {
      const mapping: any = ProvisionCommand.mappings[mappingKey];

      if (Array.isArray(obj)) {
        obj.forEach((o, i) => {
          const ctx = ProvisionCommand.clone(context);
          ctx._path += `[${i}]`;
          this.provisionObject(o, ctx, mapping, cmd);
        });
      }
      else {
        this.provisionObject(obj, ProvisionCommand.clone(context), mapping, cmd);
      }
    }

    if (!Array.isArray(obj) && typeof obj === 'object') {
      const keys: string[] = Object.keys(obj);
      keys.forEach(k => {
        const childContext = ProvisionCommand.clone(context);
        childContext._path += `.${k}`;
        this.provision(obj[k], childContext, cmd);
      });
    }
  }

  private processQueue(): void {
    const commandInfo: any = this.queue.shift();
    const log: any[] = [];
    const cmd = vorpal.find(commandInfo.command);
    const cmdInstance = {
      log: (msg: any): void => {
        log.push(msg);
      }
    };
    (cmd as any)._fn.call(cmdInstance, { options: commandInfo.options }, () => {
      const result = log.pop();
      if (typeof result === 'undefined') {
        console.log(`${vorpal.chalk.green('✓')} ${commandInfo.context._path}`);
      }
      else {
        if (result instanceof CommandError) {
          this.errors.push({
            message: result.message,
            command: commandInfo.command,
            options: commandInfo.options,
            context: commandInfo.context
          });
          console.log(vorpal.chalk.red(`${++this.numErrors}) ${commandInfo.context._path}`));
        }
        else {
          Object.assign(commandInfo.context, result);
          console.log(`${vorpal.chalk.green('✓')} ${commandInfo.context._path}`);
        }
      }

      if (this.queue.length > 0) {
        this.processQueue();
      }
      else {
        if (this.errors.length > 0) {
          console.log('');
          console.log(vorpal.chalk.red('Errors:'));
          console.log('');
          this.errors.forEach((e, i) => {
            console.log(`${i + 1}) ${e.context._path}`);
            console.log('');
            console.log(vorpal.chalk.red(`  ${e.message}`));
            console.log(`  ${e.command}`);
            console.log(`  ${JSON.stringify(e.options)}`);
            console.log(`  ${JSON.stringify(e.context)}`);
            console.log('');
          });
        }
        this.cb();
      }
    });
  }

  private provisionObject(obj: any, context: any, mapping: any, cmd: CommandInstance): void {
    let command = ProvisionCommand.getCommandForObject(obj, mapping);
    if (!command) {
      cmd.log(`${vorpal.chalk.yellow('!')} No command found`);
      return;
    }

    let commands: any[] = Array.isArray(command) ? command : [command];

    commands.forEach(c => {
      const commandName: string = ProvisionCommand.getCommandName(c.command);
      if (!c.options) {
        c.options = {};
      }
      c.options.output = 'json';
      Object.assign(c.options, ProvisionCommand.mapObjectToCommand(obj, commandName, mapping));

      this.queue.push({
        command: c.command,
        options: ProvisionCommand.clone(c.options),
        context: ProvisionCommand.clone(context)
      });

      if (!this.runFirst) {
        this.runFirst = true;
        this.processQueue();
      }
    });
  }

  private static mapObjectToCommand(obj: any, commandName: string, mapping: any): any {
    const command = vorpal.find(commandName);
    if (!command) {
      return {};
    }

    const options: any = {};
    const commandOptions: string[] = (command as any).options.map(o => o.long);
    const objProperties: string[] = Object.keys(obj);
    if (mapping.mappings && mapping.mappings._value) {
      options[mapping.mappings._value] = obj;
    }

    objProperties.forEach(p => {
      const propertyName: string = mapping.mappings && mapping.mappings[p] ?
        mapping.mappings[p] :
        `${p[0].toLowerCase()}${p.substr(1)}`;
      const optionName: string = `--${propertyName}`;
      const option = commandOptions.find(o => o === optionName);
      if (option) {
        options[propertyName] = '' + obj[p];
      }
    });

    return options;
  }

  private static getCommandName(command: string): string {
    let i: number;
    if ((i = command.indexOf('-')) > -1) {
      return command.substr(0, i).trimRight();
    }
    else {
      return command;
    }
  }

  private static getCommandForObject(obj: any, mapping: any): any | undefined {
    if (mapping.command) {
      return mapping.command;
    }

    if (obj.Action && mapping.actions) {
      return mapping.actions[obj.Action];
    }

    return;
  }

  private static getMappingKey(objectType: string): string | undefined {
    return ProvisionCommand.mappingsKeys.find(v => objectType.endsWith(v));
  }

  private static clone(obj: any): any {
    return JSON.parse(JSON.stringify(obj));
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.PROVISION).helpInformation());
    log(
      `  Remarks:

    Before running the ${chalk.grey(this.name)} command, ensure that you're connected
    to all the Office 365 services that are referenced in the specified provisioning
    template.

  Examples:
  
    Show the information about the current connection to SharePoint Online
      ${chalk.grey(config.delimiter)} ${this.name}
`);
  }
}

module.exports = new ProvisionCommand();