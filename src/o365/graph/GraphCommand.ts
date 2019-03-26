import Command, { CommandAction, CommandError } from '../../Command';
import auth from '../../Auth';

export default abstract class GraphCommand extends Command {
  protected get resource(): string {
    return 'https://graph.microsoft.com';
  }

  public action(): CommandAction {
    const cmd: GraphCommand = this;

    return function (this: CommandInstance, args: any, cb: (err?: any) => void) {
      auth
        .restoreAuth()
        .then((): void => {
          cmd.initAction(args, this);

          if (!auth.service.connected) {
            cb(new CommandError('Log in to the Microsoft Graph first'));
            return;
          }

          cmd.commandAction(this, args, cb);
        }, (error: any): void => {
          cb(new CommandError(error));
        });
    }
  }
}