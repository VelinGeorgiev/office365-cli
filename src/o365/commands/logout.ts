import auth from '../../Auth';
import config from '../../config';
import commands from './commands';
import Command, {
  CommandError,
} from '../../Command';
import appInsights from '../../appInsights';

const vorpal: Vorpal = require('../../vorpal-init');

class LogoutCommand extends Command {
  public get name(): string {
    return commands.LOGOUT;
  }

  public get description(): string {
    return 'Log out from Office 365';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: () => void): void {
    const chalk = vorpal.chalk;

    appInsights.trackEvent({
      name: this.getUsedCommandName(cmd)
    });

    if (this.verbose) {
      cmd.log('Logging out from Microsoft Graph...');
    }

    const logout: () => void = (): void => {
      auth.service.logout();
      if (this.verbose) {
        cmd.log(chalk.green('DONE'));
      }
      cb();
    }

    auth
      .clearConnectionInfo()
      .then((): void => {
        logout();
      }, (error: any): void => {
        if (this.debug) {
          cmd.log(new CommandError(error));
        }

        logout();
      });
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.LOGOUT).helpInformation());
    log(
      `  Remarks:

    The ${chalk.blue(commands.LOGOUT)} command logs out from the Microsoft Graph
    and removes any access and refresh tokens from memory.

  Examples:
  
    Log out from Microsoft Graph
      ${chalk.grey(config.delimiter)} ${commands.LOGOUT}

    Log out from Microsoft Graph in debug mode including detailed debug
    information in the console output
      ${chalk.grey(config.delimiter)} ${commands.LOGOUT} --debug
`);
  }
}

module.exports = new LogoutCommand();