import {
  Rule,
  SchematicContext,
  Tree
} from '@angular-devkit/schematics';

export default function (options: any): Rule {
  return (tree: Tree, _context: SchematicContext) => {
    if (tree.exists('.npmignore')) {
      tree.delete('.npmignore');
    }
    
    return tree;
  };
}
