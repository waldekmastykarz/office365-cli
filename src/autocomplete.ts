const omelette: (template: string) => Omelette = require('omelette');
import * as fs from 'fs';
import * as path from 'path';

class Autocomplete {
  private static autocompleteFilePath: string = path.join(__dirname, `..${path.sep}commands.json`);
  private omelette: Omelette;
  private commands: any = {};

  constructor() {
    this.init();
  }

  private init(): void {
    if (fs.existsSync(Autocomplete.autocompleteFilePath)) {
      try {
        const data: string = fs.readFileSync(Autocomplete.autocompleteFilePath, 'utf-8');
        this.commands = JSON.parse(data);
      }
      catch { }
    }

    const _this = this;

    function handleAutocomplete(this: any, fragment: string, data: any): void {
      let replies: Object | string[] = {};
      let allWords: string[] = [];

      if (data.fragment === 1) {
        replies = Object.keys(_this.commands);
      }
      else {
        allWords = data.line.split(/\s+/).slice(1, -1);
        // build array of words to use as a path to retrieve completion
        // options from the commands tree
        const words: string[] = allWords
          .filter((e: string, i: number): boolean => {
            if (e.indexOf('-') !== 0) {
              // if the word is not an option check if it's not
              // option's value, eg. --output json, in which case
              // the suggestion should be command options
              return i === 0 || allWords[i - 1].indexOf('-') !== 0;
            }
            else {
              // remove all options but last one
              return i === allWords.length - 1;
            }
          });
        let accessor: Function = new Function('_', "return _['" + (words.join("']['")) + "']");

        replies = accessor(_this.commands);
        // if the last word is an option without autocomplete
        // suggest other options from the same command
        if (words[words.length - 1].indexOf('-') === 0 &&
          !Array.isArray(replies)) {
          accessor = new Function('_', "return _['" + (words.filter(w => w.indexOf('-') !== 0).join("']['")) + "']");
          replies = accessor(_this.commands);

          if (!Array.isArray(replies)) {
            replies = Object.keys(replies);
          }
        }
      }

      if (!Array.isArray(replies)) {
        replies = Object.keys(replies);
      }

      // remove options that already have been used
      replies = (replies as string[]).filter(r => r.indexOf('-') !== 0 || allWords.indexOf(r) === -1);

      this.reply(replies);
    }

    this.omelette = omelette('o365|office365');
    this.omelette.on('complete', handleAutocomplete);
    this.omelette.init();
  }

  public generateAutocomplete(vorpal: Vorpal): void {
    const autocomplete: any = {};
    const commands: CommandInfo[] = vorpal.commands;
    const visibleCommands: CommandInfo[] = commands.filter(c => !c._hidden);
    visibleCommands.forEach(c => {
      Autocomplete.processCommand(c._name, c, autocomplete);
      c._aliases.forEach(a => Autocomplete.processCommand(a, c, autocomplete));
    });

    fs.writeFileSync(Autocomplete.autocompleteFilePath, JSON.stringify(autocomplete, null, 2));
  }

  private static processCommand(commandName: string, commandInfo: CommandInfo, autocomplete: any) {
    const chunks: string[] = commandName.split(' ');
    let parent: any = autocomplete;
    for (let i: number = 0; i < chunks.length; i++) {
      const current: any = chunks[i];
      if (current === 'exit' || current === 'quit') {
        continue;
      }

      if (!parent[current]) {
        if (i < chunks.length - 1) {
          parent[current] = {};
        }
        else {
          // last chunk, add options
          const optionsArr: string[] = commandInfo.options.map(o => o.short)
            .concat(commandInfo.options.map(o => o.long)).filter(o => o != null);
          const optionsObj: any = {};
          optionsArr.forEach(o => {
            const option: CommandOption = commandInfo.options.filter(opt => opt.long === o || opt.short === o)[0];
            if (option.autocomplete) {
              optionsObj[o] = option.autocomplete;
            }
            else {
              optionsObj[o] = {};
            }
          });
          parent[current] = optionsObj;
        }
      }

      parent = parent[current];
    }
  }

  public setupAutocomplete(): void {
    this.omelette.setupShellInitFile();
  }
}

export const autocomplete = new Autocomplete();