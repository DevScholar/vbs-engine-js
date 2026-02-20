import { VbsEngine } from '../src/index.ts';
import * as path from 'path';
import * as fs from 'fs';

globalThis.nodePath = path;
globalThis.nodeFs = fs;

const engine = new VbsEngine();

const currentDir = process.cwd();

engine.run(`
  ' Node.js modules are automatically available from globalThis
  currentDir = nodePath.resolve(".")
  
  ' Read package.json
  content = nodeFs.readFileSync("package.json", "utf8")
  
  ' Print result
  MsgBox "Current Directory: " & currentDir
  MsgBox "Package.json length: " & Len(content)
`);

console.log('VBScript executed successfully!');
console.log('currentDir (from VBS):', engine.getVariableAsJs('currentDir'));
console.log('content length:', engine.getVariableAsJs('content'));
