import { VbsEngine } from '../src/index.ts';
import * as path from 'path';
import * as fs from 'fs';

globalThis.nodePath = path;
globalThis.nodeFs = fs;

const engine = new VbsEngine();

engine.addCode(`
  ' Node.js modules are automatically available from globalThis
  Function GetCurrentDir()
      GetCurrentDir = nodePath.resolve(".")
  End Function
  
  Function ReadPackageJson()
      ReadPackageJson = nodeFs.readFileSync("package.json", "utf8")
  End Function
`);

const currentDir = engine.run('GetCurrentDir');
const content = engine.run('ReadPackageJson');

console.log('Current Directory:', currentDir);
console.log('Package.json length:', typeof content === 'string' ? content.length : 0);
