import Command, { CommandAction, CommandError } from '../../Command';
import auth from '../../Auth';

export default abstract class GraphCommand extends Command {
  protected get resource(): string {
    return 'https://graph.microsoft.com';
  }
}