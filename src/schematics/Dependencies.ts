export interface Dependencies {
  install: {
    dependencies: string[];
    devDependencies: string[];
  },
  uninstall?: {
    dependencies?: string[];
    devDependencies?: string[];
  }
}