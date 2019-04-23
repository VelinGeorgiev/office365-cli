import * as fs from 'fs';
import * as path from 'path';

class ListCommands {

  public generateYamlList(vorpal: Vorpal): void {
    const commandsInfo: any = this.getCommandsInfo(vorpal);
    fs.writeFileSync(path.join(__dirname, `..${path.sep}commands.yaml`), commandsInfo);
  }

  public generateShortcutsYamlList(vorpal: Vorpal): void {
    const commandsShortcutsInfo: any = this.getShortcutsInfo(vorpal, 'yaml');
    fs.writeFileSync(path.join(__dirname, `..${path.sep}shortcuts.yaml`), commandsShortcutsInfo);
  }

  public generateShortcutsMarkdownList(vorpal: Vorpal): void {
    const commandsShortcutsInfo: any = this.getShortcutsInfo(vorpal, 'markdown');
    fs.writeFileSync(path.join(__dirname, `..${path.sep}shortcuts.md`), commandsShortcutsInfo);
  }

  private getCommandsInfo(vorpal: Vorpal): any {
    let commandsInfo: string = '';
    const commands: CommandInfo[] = vorpal.commands;
    const visibleCommands: CommandInfo[] = commands.filter(c => !c._hidden);
    visibleCommands.forEach(c => {
      commandsInfo += ListCommands.processCommand(c._name, c);
      c._aliases.forEach(a => {
       commandsInfo += ListCommands.processCommand(a, c);
      });
    });
    return commandsInfo;
  }

  private static processCommand(commandName: string, commandInfo: CommandInfo): string {
    let result: string = `${commandName}:\n`;

    // for(const opt of commandInfo.options) {

    //   const flags: string = (opt as any).flags;
    //   if(['-o, --output [output]','--verbose','--debug'].indexOf(flags) === -1) {
    //     result += `  - '${flags}'\n`;
    //   }
    // }

    return result;
  }

  private getShortcutsInfo(vorpal: Vorpal, format: string = 'yaml'): any {
    let commandsShortcutsInfo: string = '';
    const commandsShortcuts: any = {};
    const commands: CommandInfo[] = vorpal.commands;
    const visibleCommands: CommandInfo[] = commands.filter(c => !c._hidden);
    visibleCommands.forEach(c => {
      ListCommands.processShortcut(c, commandsShortcuts);
      c._aliases.forEach(a => {
        ListCommands.processShortcut(c, commandsShortcuts);
      });
    });

    if(format === 'markdown') {
      commandsShortcutsInfo = 'Short option|Asossiated long option\n|---|---\n';

      for(const key of Object.keys(commandsShortcuts)) {
        commandsShortcutsInfo += `${key}|${commandsShortcuts[key].join(', ')}\n`;
      }
    } else {

      commandsShortcutsInfo = 'short options:\n';

      for(const key of Object.keys(commandsShortcuts)) {
        commandsShortcutsInfo += `'${key}': '${commandsShortcuts[key].join(', ')}'\n`;
      }
    }
    
    return commandsShortcutsInfo;
  }

  private static processShortcut(commandInfo: CommandInfo, commandsShortcuts: any): void {

    for(const opt of commandInfo.options) {

      const flags: string = (opt as any).flags;
      if(['-o, --output [output]','--verbose','--debug'].indexOf(flags) === -1) {
        
        const short: string = (opt as any).short;
        if(!short) {
          continue;
        }

        const long: string = (opt as any).long;

        if(!commandsShortcuts[short]) {
          commandsShortcuts[short] = [long];
        } else {
          if(commandsShortcuts[short].indexOf(long) === -1) {
            commandsShortcuts[short].push(long);
          }
        }
      }
    }
  }
}

export const listCommands = new ListCommands();