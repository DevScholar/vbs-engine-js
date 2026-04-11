import { VbsEngine } from '../src/index.ts';

// Optional dependency: @devscholar/node-ps1-dotnet provides ActiveXObject for
// COM object support in Node.js. Install it separately before running this example:
//   npm install @devscholar/node-ps1-dotnet
const activexModuleName = '@devscholar/node-ps1-dotnet/activex';
const { ActiveXObject } = await import(activexModuleName).catch(() => {
  throw new Error(
    'This example requires the optional dependency @devscholar/node-ps1-dotnet. ' +
    'Install it first with: npm install @devscholar/node-ps1-dotnet'
  );
});

// Expose ActiveXObject to globalThis so the VBS engine's CreateObject() can use
// it — just like Internet Explorer provided window.ActiveXObject natively.
(globalThis as any).ActiveXObject = ActiveXObject;

const engine = new VbsEngine();

engine.addCode(`
  Dim WshShell, BtnCode
  Set WshShell = CreateObject("WScript.Shell")

  BtnCode = WshShell.Popup("Do you feel alright?", 7, "Answer This Question:", 4 + 32)

  Select Case BtnCode
    Case 6
      MsgBox "Glad to hear you feel alright."
    Case 7
      MsgBox "Hope you're feeling better soon."
    Case -1
      MsgBox "Is there anybody out there?"
  End Select
`);

if (engine.error) {
  console.error('VBS Error:', JSON.stringify(engine.error));
}
