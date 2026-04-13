import { VbsEngine } from '../src/index.ts';

// Optional dependency: @devscholar/node-ps1-dotnet provides GetObject for
// COM moniker binding in Node.js. Install it separately before running this example:
//   npm install @devscholar/node-ps1-dotnet
const activexModuleName = '@devscholar/node-ps1-dotnet/activex';
const { GetObject } = await import(activexModuleName).catch(() => {
  throw new Error(
    'This example requires the optional dependency @devscholar/node-ps1-dotnet. ' +
    'Install it first with: npm install @devscholar/node-ps1-dotnet'
  );
});

// Expose GetObject to globalThis so the VBS engine's GetObject() can use it —
// just like WSH provided GetObject() natively.
(globalThis as any).GetObject = GetObject;

const engine = new VbsEngine();

engine.addCode(`
  Dim wmi, results, item

  Set wmi = GetObject("winmgmts:")

  ' Query OS info
  Set results = wmi.ExecQuery("SELECT Caption, Version, OSArchitecture FROM Win32_OperatingSystem")
  For Each item In results
    Print "OS:   " & item.Caption
    Print "Ver:  " & item.Version
    Print "Arch: " & item.OSArchitecture
  Next

  ' Query CPU info
  Set results = wmi.ExecQuery("SELECT Name, NumberOfCores FROM Win32_Processor")
  For Each item In results
    Print "CPU:   " & item.Name
    Print "Cores: " & item.NumberOfCores
  Next
`);

if (engine.error) {
  console.error('VBS Error:', JSON.stringify(engine.error));
}
