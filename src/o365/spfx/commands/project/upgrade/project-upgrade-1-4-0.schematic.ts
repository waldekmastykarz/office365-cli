import {
  Rule,
  SchematicContext,
  Tree
} from '@angular-devkit/schematics';
import { SchematicsTools } from '../../../../../schematics';

export default function (options: any): Rule {
  return (tree: Tree, _context: SchematicContext) => {
    const update = SchematicsTools.updateYoRc(tree, '1.4.1');
    if (update) {
      tree.overwrite('.yo-rc.json', update);
    }

    return tree;
  };
}