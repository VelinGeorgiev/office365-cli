import auth from '../../Auth';
import config from '../../config';
import commands from './commands';
import Command, {
  CommandError
} from '../../Command';
import Utils from '../../Utils';
import { AuthType } from '../../Auth';

const vorpal: Vorpal = require('../../vorpal-init');

class StatusCommand extends Command {
  public get name(): string {
    return commands.STATUS;
  }

  public get description(): string {
    return 'Shows Office 365 login status';
  }

  public commandAction(cmd: CommandInstance, args: {}, cb: (err?: any) => void): void {
    auth
      .restoreAuth()
      .then((): void => {
        if (auth.service.connected) {
          if (this.debug) {
            cmd.log({
              connectedAs: Utils.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].value),
              authType: AuthType[auth.service.authType],
              accessTokens: JSON.stringify(auth.service.accessTokens, null, 2),
              refreshToken: auth.service.refreshToken
            });
          }
          else {
            cmd.log({
              connectedAs: Utils.getUserNameFromAccessToken(auth.service.accessTokens[auth.defaultResource].value)
            });
          }
        }
        else {
          if (this.verbose) {
            cmd.log('Logged out from Office 365');
          }
          else {
            cmd.log('Logged out');
          }
        }
        cb();
      }, (error: any): void => {
        cb(new CommandError(error));
      });
  }

  public commandHelp(args: any, log: (help: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(commands.STATUS).helpInformation());
    log(
      `  Remarks:

    If you are logged in to Office 365, the ${chalk.blue(commands.STATUS)} command
    will show you information about the currently stored refresh and access
    token and the expiration date and time of the access token when run in debug
    mode.

  Examples:
  
    Show the information about the current login to Office 365
      ${chalk.grey(config.delimiter)} ${commands.STATUS}
`);
  }
}

module.exports = new StatusCommand();