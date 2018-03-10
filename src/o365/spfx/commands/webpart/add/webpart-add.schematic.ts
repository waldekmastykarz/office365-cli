// import { normalize, strings } from '@angular-devkit/core';
// import {
//   Rule,
//   SchematicsException,
//   apply,
//   branchAndMerge,
//   chain,
//   filter,
//   mergeWith,
//   move,
//   noop,
//   template,
//   url,
// } from '@angular-devkit/schematics';
// import { Schema as ClassOptions } from './schema';

// export default function (options: ClassOptions): Rule {
//   // options.type = !!options.type ? `.${options.type}` : '';
//   // options.path = options.path ? normalize(options.path) : options.path;

//   const sourceDir = options.sourceDir;
//   if (!sourceDir) {
//     throw new SchematicsException(`sourceDir option is required.`);
//   }

//   const templateSource = apply(url('./files'), [
//     template({
//       ...strings,
//       ...options,
//     }),
//     move(sourceDir),
//   ]);

//   return chain([
//     branchAndMerge(chain([
//       mergeWith(templateSource),
//     ])),
//   ]);
// }