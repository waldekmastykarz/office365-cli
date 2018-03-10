import { Tree } from "@angular-devkit/schematics";
import { spawn } from 'child_process';

export class SchematicsTools {
  public static updateYoRc(tree: Tree, version: string): string | undefined {
    const yoRcFile = tree.read('.yo-rc.json');
    if (yoRcFile === null) {
      return;
    }

    const yoRc: any = JSON.parse(yoRcFile.toString('utf-8'));
    if (yoRc &&
      yoRc['@microsoft/generator-sharepoint'] &&
      yoRc['@microsoft/generator-sharepoint'].version) {
      yoRc['@microsoft/generator-sharepoint'].version = version;
    }

    return JSON.stringify(yoRc, null, 2);
  }

  public static runNpm(args: string[]): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (err: any) => void): void => {
      spawn('npm', args, {
        stdio: 'inherit',
        shell: true
      })
        .on('close', (code: number) => {
          if (code === 0) {
            resolve();
          }
          else {
            reject('Installing dependencies failed');
          }
        });
    });
  }
}