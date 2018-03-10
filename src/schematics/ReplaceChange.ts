import { Change, Host } from '../schematics';

export class ReplaceChange implements Change {
  order: number;
  description: string;

  constructor(public path: string, private pos: number, private oldText: string,
              private newText: string) {
    if (pos < 0) {
      throw new Error('Negative positions are invalid');
    }
    this.description = `Replaced ${oldText} into position ${pos} of ${path} with ${newText}`;
    this.order = pos;
  }

  apply(host: Host): Promise<void> {
    return host.read(this.path).then(content => {
      const prefix = content.substring(0, this.pos);
      const suffix = content.substring(this.pos + this.oldText.length);
      const text = content.substring(this.pos, this.pos + this.oldText.length);

      if (text !== this.oldText) {
        return Promise.reject(new Error(`Invalid replace: "${text}" != "${this.oldText}".`));
      }

      return host.write(this.path, `${prefix}${this.newText}${suffix}`);
    });
  }
}