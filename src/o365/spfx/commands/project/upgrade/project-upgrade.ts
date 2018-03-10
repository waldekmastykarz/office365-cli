import commands from '../../../commands';
import GlobalOptions from '../../../../../GlobalOptions';
import SchematicsCommand from '../../../../../schematics/SchematicsCommand';
import * as path from 'path';
import * as fs from 'fs';
import { CommandError } from '../../../../../Command';
import { Dependencies, SchematicsTools } from '../../../../../schematics';

const vorpal: Vorpal = require('../../../../../vorpal-init');

interface Options extends GlobalOptions {
  toVersion: string;
}

interface CommandArgs {
  options: Options;
}

class SpfxProjectUpgradeCommand extends SchematicsCommand {
  private schematicName: string = '';
  private projectVersion: string | undefined;
  private toVersion: string = '';

  public get name(): string {
    return `${commands.PROJECT_UPGRADE}`;
  }

  public get description(): string {
    return 'Upgrades project to the specified version';
  }

  public get schematic(): string {
    return this.schematicName;
  }

  public commandAction(cmd: CommandInstance, args: CommandArgs, cb: () => void): void {
    const supportedVersions: string[] = [
      '1.3.4',
      '1.4.0',
      '1.4.1'
    ];

    this.toVersion = args.options.toVersion ? args.options.toVersion : supportedVersions[supportedVersions.length - 1];

    if (supportedVersions.indexOf(this.toVersion) < 0) {
      throw `Office 365 CLI doesn't support upgrading SharePoint Framework projects to version ${this.toVersion}`;
    }

    // get project version
    this.projectVersion = this.getProjectVersion();
    if (!this.projectVersion) {
      throw `Unable to determine the version of the current SharePoint Framework project`;
    }

    const pos: number = supportedVersions.indexOf(this.projectVersion);
    if (pos < 0) {
      throw `Office 365 CLI doesn't support upgrading projects build on SharePoint Framework v${this.projectVersion}`;
    }

    if (pos === supportedVersions.indexOf(this.toVersion)) {
      throw 'Project doesn\'t need to be upgraded';
    }

    if (pos === supportedVersions.length - 1) {
      throw `${this.projectVersion} is the latest version supported by the Office 365 CLI`;
    }

    this.schematicName = `spfx-project-upgrade-${this.projectVersion.replace(/\./g, '-')}`;

    cmd.log(`Upgrading project to v${supportedVersions[pos+1]}...`);

    const isReactProject: boolean = this.isReactProject();
    const dependencies = this.getDependenciesForVersion(this.projectVersion, isReactProject);
    this.updateDependencies(cmd, dependencies, cb);
  }

  protected postSchematicAction(cmd: CommandInstance, args: any, cb: () => void): void {
    if (this.projectVersion !== this.toVersion) {
      const action = this.action().bind(cmd, args, cb);
      action();
    }
    else {
      cb();
    }
  }

  private getProjectVersion(): string | undefined {
    const yoRcPath: string = path.resolve(this.projectRootPath as string, '.yo-rc.json');
    if (fs.existsSync(yoRcPath)) {
      try {
        const yoRc: any = JSON.parse(fs.readFileSync(yoRcPath, 'utf-8'));
        if (yoRc && yoRc['@microsoft/generator-sharepoint'] &&
          yoRc['@microsoft/generator-sharepoint'].version) {
          return yoRc['@microsoft/generator-sharepoint'].version;
        }
      }
      catch { }
    }

    const packageJsonPath: string = path.resolve(this.projectRootPath as string, 'package.json');
    if (fs.existsSync(packageJsonPath)) {
      try {
        const packageJson: any = JSON.parse(fs.readFileSync(packageJsonPath, 'utf-8'));
        if (packageJson &&
          packageJson.dependencies &&
          packageJson.dependencies['@microsoft/sp-core-library']) {
          const coreLibVersion: string = packageJson.dependencies['@microsoft/sp-core-library'];
          return coreLibVersion.replace(/[^0-9\.]/g, '');
        }
      }
      catch { }
    }

    return undefined;
  }

  private isReactProject(): boolean {
    let isReactProject: boolean = false;

    const packageJsonPath: string = path.resolve(this.projectRootPath as string, 'package.json');
    if (fs.existsSync(packageJsonPath)) {
      try {
        const packageJson: any = JSON.parse(fs.readFileSync(packageJsonPath, 'utf-8'));
        isReactProject = typeof packageJson.dependencies.react !== 'undefined';
      }
      catch { }
    }

    return isReactProject;
  }

  private getDependenciesForVersion(projectVersion: string, isReactProject: boolean): Dependencies {
    const dep: Dependencies = {
      install: {
        dependencies: [],
        devDependencies: []
      }
    };

    switch (projectVersion) {
      case '1.3.4':
        dep.install.dependencies = [
          '@microsoft/sp-core-library@1.4.0',
          '@microsoft/sp-webpart-base@1.4.0',
          '@microsoft/sp-lodash-subset@1.4.0',
          '@microsoft/sp-office-ui-fabric-core@1.4.0',
          `@types/webpack-env@'>=1.12.1 <1.14.0'`
        ];
        if (isReactProject) {
          dep.install.dependencies.push(
            'react@15.6.2',
            'react-dom@15.6.2',
            '@types/react@15.6.6',
            '@types/react-dom@15.5.6'
          );
        }
        dep.install.devDependencies = [
          '@microsoft/sp-build-web@1.4.0',
          '@microsoft/sp-module-interfaces@1.4.0',
          '@microsoft/sp-webpart-workbench@1.4.0',
          'gulp@3.9.1',
          `@types/chai@'>=3.4.34 <3.6.0'`,
          `@types/mocha@'>=2.2.33 <2.6.0'`,
          'ajv@5.2.2'
        ];
        dep.uninstall = {
          dependencies: [
            "@types/react-addons-shallow-compare",
            "@types/react-addons-test-utils",
            "@types/react-addons-update"
          ]
        }
        break;
      case '1.4.0':
        dep.install.dependencies = [
          '@microsoft/sp-core-library@1.4.1',
          '@microsoft/sp-webpart-base@1.4.1',
          '@microsoft/sp-lodash-subset@1.4.1',
          '@microsoft/sp-office-ui-fabric-core@1.4.1',
          `@types/webpack-env@'>=1.12.1 <1.14.0'`
        ];
        if (isReactProject) {
          dep.install.dependencies.push(
            'react@15.6.2',
            'react-dom@15.6.2',
            '@types/react@15.6.6',
            '@types/react-dom@15.5.6'
          );
        }
        dep.install.devDependencies = [
          '@microsoft/sp-build-web@1.4.1',
          '@microsoft/sp-module-interfaces@1.4.1',
          '@microsoft/sp-webpart-workbench@1.4.1',
          'gulp@3.9.1',
          `@types/chai@'>=3.4.34 <3.6.0'`,
          `@types/mocha@'>=2.2.33 <2.6.0'`,
          'ajv@5.2.2'
        ];
        break;
    }

    return dep;
  }

  private updateDependencies(cmd: CommandInstance, dep: Dependencies, cb: () => void): void {
    cmd.log(`Updating dependencies...`);

    SchematicsTools
      .runNpm(['install', '--quiet'].concat(dep.install.dependencies, '--save'))
      .then((): Promise<void> => {
        return SchematicsTools.runNpm(['install', '--quiet'].concat(dep.install.devDependencies, '--save-dev'));
      })
      .then((): Promise<void> => {
        if (dep.uninstall && dep.uninstall.dependencies) {
          return SchematicsTools.runNpm(['uninstall', '--quiet'].concat(dep.uninstall.dependencies, '--save'));
        }
        else {
          return Promise.resolve();
        }
      })
      .then((): Promise<void> => {
        if (dep.uninstall && dep.uninstall.devDependencies) {
          return SchematicsTools.runNpm(['uninstall', '--quiet'].concat(dep.uninstall.devDependencies, '--save-dev'));
        }
        else {
          return Promise.resolve();
        }
      })
      .then((): void => {
        cb();
      }, (err: any): void => {
        throw err;
      });
  }

  public commandHelp(args: CommandArgs, log: (message: string) => void): void {
    const chalk = vorpal.chalk;
    log(vorpal.find(this.getCommandName()).helpInformation());
    //     log(
    //       `  ${chalk.yellow('Important:')} before using this command, connect to a SharePoint Online site,
    //   using the ${chalk.blue(commands.CONNECT)} command.

    //   Remarks:

    //     To add an app to the tenant app catalog, you have to first connect to a SharePoint site using the
    //     ${chalk.blue(commands.CONNECT)} command, eg. ${chalk.grey(`${config.delimiter} ${commands.CONNECT} https://contoso.sharepoint.com`)}.

    //     When specifying the path to the app package file you can use both relative and absolute paths.
    //     Note, that ~ in the path, will not be resolved and will most likely result in an error.

    //     If you try to upload a package that already exists in the tenant app catalog without specifying
    //     the ${chalk.blue('--overwrite')} option, the command will fail with an error stating that the
    //     specified package already exists.

    //   Examples:

    //     Add the ${chalk.grey('spfx.sppkg')} package to the tenant app catalog
    //       ${chalk.grey(config.delimiter)} ${commands.APP_ADD} -p /Users/pnp/spfx/sharepoint/solution/spfx.sppkg

    //     Overwrite the ${chalk.grey('spfx.sppkg')} package in the tenant app catalog with the newer version
    //       ${chalk.grey(config.delimiter)} ${commands.APP_ADD} -p sharepoint/solution/spfx.sppkg --overwrite

    //   More information:

    //     Application Lifecycle Management (ALM) APIs
    //       https://docs.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins
    // `);
  }
}

module.exports = new SpfxProjectUpgradeCommand();