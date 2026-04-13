import dotnet from '@devscholar/node-ps1-dotnet';
import { VbsEngine } from '../src/index.ts';

dotnet.load('System.Windows.Forms');
dotnet.load('System.Drawing');

const System = (dotnet as any).System;

// Expose .NET namespaces so VBScript can use `New Forms.Form`, `New Drawing.Font`, etc.
// VBScript's New keyword cannot call arbitrary JS constructors directly, but with the
// globalThis extension in this engine it resolves dotted names like `New Forms.Form`
// by looking up `globalThis.Forms.Form` and calling Reflect.construct on it.
(globalThis as any).Forms = System.Windows.Forms;
(globalThis as any).Drawing = System.Drawing;

System.Windows.Forms.Application.EnableVisualStyles();
System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);

const Engine = new VbsEngine();

Engine.addCode(`
  Dim Form, Label, Button
  Dim ClickCount
  ClickCount = 0

  Set Form = New Forms.Form
  Form.Text = "Counter App"
  Form.Width = 640
  Form.Height = 480
  Form.StartPosition = 1

  Set Label = New Forms.Label
  Label.Text = "Clicks: 0"
  Label.Font = New Drawing.Font("Arial", 24)
  Label.AutoSize = True
  Label.Location = New Drawing.Point(90, 30)
  Call Form.Controls.Add(Label)

  Set Button = New Forms.Button
  Button.Text = "Click to Add"
  Button.Font = New Drawing.Font("Arial", 14)
  Button.AutoSize = True
  Button.Location = New Drawing.Point(100, 90)
  Button.add_Click GetRef("OnButtonClick")
  Call Form.Controls.Add(Button)

  Call Forms.Application.Run(Form)

  Sub OnButtonClick(Sender, E)
    ClickCount = ClickCount + 1
    Dim Message
    Message = "Clicked " & ClickCount & " times"
    Label.Text = Message
    console.log Message
  End Sub
`);

if (Engine.error) {
  console.error('VBS Error:', JSON.stringify(Engine.error));
}
