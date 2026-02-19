import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';
import type { VbValue } from '../runtime/values.ts';

describe('Interpreter', () => {
  describe('basic expressions', () => {
    it('should evaluate arithmetic expressions', () => {
      const engine = new VbsEngine();
      engine.run('x = 1 + 2 * 3');
      const x = engine.getVariable('x');
      expect(x.value).toBe(7);
    });

    it('should evaluate string concatenation', () => {
      const engine = new VbsEngine();
      engine.run('x = "hello" & " " & "world"');
      const x = engine.getVariable('x');
      expect(x.value).toBe('hello world');
    });

    it('should evaluate comparison expressions', () => {
      const engine = new VbsEngine();
      engine.run('x = 5 > 3');
      const x = engine.getVariable('x');
      expect(x.value).toBe(true);
    });

    it('should evaluate logical expressions', () => {
      const engine = new VbsEngine();
      engine.run('x = True And False');
      const x = engine.getVariable('x');
      expect(x.value).toBe(false);
    });

    it('should evaluate integer division', () => {
      const engine = new VbsEngine();
      engine.run('x = 7 \\ 2');
      const x = engine.getVariable('x');
      expect(x.value).toBe(3);
    });

    it('should evaluate modulo', () => {
      const engine = new VbsEngine();
      engine.run('x = 7 Mod 3');
      const x = engine.getVariable('x');
      expect(x.value).toBe(1);
    });

    it('should evaluate power', () => {
      const engine = new VbsEngine();
      engine.run('x = 2 ^ 3');
      const x = engine.getVariable('x');
      expect(x.value).toBe(8);
    });

    it('should evaluate string comparison case-insensitively', () => {
      const engine = new VbsEngine();
      engine.run('x = "HELLO" = "hello"');
      const x = engine.getVariable('x');
      expect(x.value).toBe(true);
    });

    it('should evaluate Not operator', () => {
      const engine = new VbsEngine();
      engine.run('x = Not True');
      const x = engine.getVariable('x');
      expect(x.value).toBe(false);
    });
  });

  describe('built-in functions', () => {
    it('should use Len function', () => {
      const engine = new VbsEngine();
      engine.run('x = Len("hello")');
      const x = engine.getVariable('x');
      expect(x.value).toBe(5);
    });

    it('should use UCase function', () => {
      const engine = new VbsEngine();
      engine.run('x = UCase("hello")');
      const x = engine.getVariable('x');
      expect(x.value).toBe('HELLO');
    });

    it('should use LCase function', () => {
      const engine = new VbsEngine();
      engine.run('x = LCase("HELLO")');
      const x = engine.getVariable('x');
      expect(x.value).toBe('hello');
    });

    it('should use Abs function', () => {
      const engine = new VbsEngine();
      engine.run('x = Abs(-5)');
      const x = engine.getVariable('x');
      expect(x.value).toBe(5);
    });

    it('should use Left function', () => {
      const engine = new VbsEngine();
      engine.run('x = Left("hello", 2)');
      const x = engine.getVariable('x');
      expect(x.value).toBe('he');
    });

    it('should use Right function', () => {
      const engine = new VbsEngine();
      engine.run('x = Right("hello", 2)');
      const x = engine.getVariable('x');
      expect(x.value).toBe('lo');
    });

    it('should use Mid function', () => {
      const engine = new VbsEngine();
      engine.run('x = Mid("hello", 2, 3)');
      const x = engine.getVariable('x');
      expect(x.value).toBe('ell');
    });

    it('should use Trim function', () => {
      const engine = new VbsEngine();
      engine.run('x = Trim("  hello  ")');
      const x = engine.getVariable('x');
      expect(x.value).toBe('hello');
    });

    it('should use CInt function', () => {
      const engine = new VbsEngine();
      engine.run('x = CInt("42")');
      const x = engine.getVariable('x');
      expect(x.value).toBe(42);
    });

    it('should use CLng function', () => {
      const engine = new VbsEngine();
      engine.run('x = CLng("123456")');
      const x = engine.getVariable('x');
      expect(x.value).toBe(123456);
    });

    it('should use CDbl function', () => {
      const engine = new VbsEngine();
      engine.run('x = CDbl("3.14")');
      const x = engine.getVariable('x');
      expect(x.value).toBe(3.14);
    });

    it('should use CStr function', () => {
      const engine = new VbsEngine();
      engine.run('x = CStr(42)');
      const x = engine.getVariable('x');
      expect(x.value).toBe('42');
    });

    it('should use TypeName function', () => {
      const engine = new VbsEngine();
      engine.run('x = TypeName("hello")');
      const x = engine.getVariable('x');
      expect(x.value).toBe('String');
    });

    it('should use IsNull function', () => {
      const engine = new VbsEngine();
      engine.run('x = IsNull(Null)');
      const x = engine.getVariable('x');
      expect(x.value).toBe(true);
    });

    it('should use IsEmpty function', () => {
      const engine = new VbsEngine();
      engine.run('Dim v\nx = IsEmpty(v)');
      const x = engine.getVariable('x');
      expect(x.value).toBe(true);
    });

    it('should use IsNumeric function', () => {
      const engine = new VbsEngine();
      engine.run('x = IsNumeric("42")');
      const x = engine.getVariable('x');
      expect(x.value).toBe(true);
    });

    it('should use Array function', () => {
      const engine = new VbsEngine();
      engine.run('x = Array(1, 2, 3)');
      const x = engine.getVariable('x');
      expect(x.type).toBe('Array');
    });

    it('should use UBound function', () => {
      const engine = new VbsEngine();
      engine.run('arr = Array(1, 2, 3)\nx = UBound(arr)');
      const x = engine.getVariable('x');
      expect(x.value).toBe(2);
    });

    it('should use LBound function', () => {
      const engine = new VbsEngine();
      engine.run('arr = Array(1, 2, 3)\nx = LBound(arr)');
      const x = engine.getVariable('x');
      expect(x.value).toBe(0);
    });
  });

  describe('control flow', () => {
    it('should execute If-Then-Else', () => {
      const engine = new VbsEngine();
      engine.run(`
If True Then
    x = 1
Else
    x = 2
End If
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(1);
    });

    it('should execute ElseIf', () => {
      const engine = new VbsEngine();
      engine.run(`
x = 0
If False Then
    x = 1
ElseIf True Then
    x = 2
Else
    x = 3
End If
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(2);
    });

    it('should execute For-To loop', () => {
      const engine = new VbsEngine();
      engine.run(`
x = 0
For i = 1 To 5
    x = x + i
Next
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(15);
    });

    it('should execute For-To-Step loop', () => {
      const engine = new VbsEngine();
      engine.run(`
x = 0
For i = 1 To 5 Step 2
    x = x + i
Next
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(9);
    });

    it('should execute For-Each loop', () => {
      const engine = new VbsEngine();
      engine.run(`
arr = Array(1, 2, 3)
x = 0
For Each item In arr
    x = x + item
Next
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(6);
    });

    it('should execute Do-While loop', () => {
      const engine = new VbsEngine();
      engine.run(`
x = 0
Do While x < 5
    x = x + 1
Loop
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(5);
    });

    it('should execute Do-Until loop', () => {
      const engine = new VbsEngine();
      engine.run(`
x = 0
Do Until x = 5
    x = x + 1
Loop
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(5);
    });

    it('should execute Select-Case', () => {
      const engine = new VbsEngine();
      engine.run(`
x = 0
Select Case 2
    Case 1
        x = 1
    Case 2
        x = 2
    Case Else
        x = 3
End Select
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(2);
    });

    it('should execute Exit For', () => {
      const engine = new VbsEngine();
      engine.run(`
x = 0
For i = 1 To 10
    If i = 5 Then Exit For
    x = x + 1
Next
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(4);
    });
  });

  describe('procedures', () => {
    it('should define and call Sub', () => {
      const engine = new VbsEngine();
      engine.run(`
Sub Test
    x = 42
End Sub

Call Test
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(42);
    });

    it('should define and call Function', () => {
      const engine = new VbsEngine();
      engine.run(`
Function Add(a, b)
    Add = a + b
End Function

x = Add(2, 3)
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(5);
    });

    it('should handle ByVal parameters', () => {
      const engine = new VbsEngine();
      engine.run(`
Sub Increment(ByVal n)
    n = n + 1
End Sub

x = 5
Call Increment(x)
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(5);
    });

    it('should handle ByRef parameters (default)', () => {
      const engine = new VbsEngine();
      engine.run(`
Sub Increment(n)
    n = n + 1
End Sub

x = 5
Call Increment(x)
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(6);
    });

    it('should handle explicit ByRef parameters', () => {
      const engine = new VbsEngine();
      engine.run(`
Sub Increment(ByRef n)
    n = n + 1
End Sub

x = 5
Call Increment(x)
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(6);
    });

    it('should handle ByRef in Function', () => {
      const engine = new VbsEngine();
      engine.run(`
Function Swap(a, b)
    Dim temp
    temp = a
    a = b
    b = temp
    Swap = True
End Function

x = 1
y = 2
result = Swap(x, y)
`);
      const x = engine.getVariable('x');
      const y = engine.getVariable('y');
      expect(x.value).toBe(2);
      expect(y.value).toBe(1);
    });
  });

  describe('classes', () => {
    it('should create class instance with New', () => {
      const engine = new VbsEngine();
      engine.run(`
Class Person
    Public Name
End Class
Dim p
Set p = New Person
result = TypeName(p)
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe('Person');
    });

    it('should access class property', () => {
      const engine = new VbsEngine();
      engine.run(`
Class Person
    Public Name
End Class
Dim p
Set p = New Person
p.Name = "John"
result = p.Name
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe('John');
    });

    it('should call class method', () => {
      const engine = new VbsEngine();
      engine.run(`
Class Calculator
    Public Function Add(a, b)
        Add = a + b
    End Function
End Class
Dim calc
Set calc = New Calculator
result = calc.Add(5, 3)
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(8);
    });

    it('should use Me keyword', () => {
      const engine = new VbsEngine();
      engine.run(`
Class Counter
    Private count
    
    Public Sub Increment()
        count = count + 1
    End Sub
    
    Public Function GetCount()
        GetCount = count
    End Function
End Class
Dim c
Set c = New Counter
c.Increment()
c.Increment()
result = c.GetCount()
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(2);
    });

    it('should use With statement', () => {
      const engine = new VbsEngine();
      engine.run(`
Class Point
    Public X
    Public Y
End Class
Dim pt
Set pt = New Point
With pt
    .X = 10
    .Y = 20
End With
result = pt.X + pt.Y
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(30);
    });
  });

  describe('error handling', () => {
    it('should handle On Error Resume Next', () => {
      const engine = new VbsEngine();
      engine.run(`
On Error Resume Next
Dim x
x = 1 / 0
result = Err.Number
`);
      const result = engine.getVariable('result');
      expect(result.value).not.toBe(0);
    });

    it('should clear error with Err.Clear', () => {
      const engine = new VbsEngine();
      engine.run(`
On Error Resume Next
Dim x
x = 1 / 0
Call Err.Clear()
result = Err.Number
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(0);
    });
  });

  describe('arrays', () => {
    it('should use Array function', () => {
      const engine = new VbsEngine();
      engine.run('arr = Array(1, 2, 3)\nresult = arr(1)');
      const result = engine.getVariable('result');
      expect(result.value).toBe(2);
    });
  });

  describe('variables and scope', () => {
    it('should handle Dim declaration', () => {
      const engine = new VbsEngine();
      engine.run(`
Dim x
x = 5
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(5);
    });

    it('should handle Const declaration', () => {
      const engine = new VbsEngine();
      engine.run(`
Const PI = 3.14159
result = PI
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(3.14159);
    });

    it('should handle Option Explicit', () => {
      const engine = new VbsEngine();
      engine.run(`
Option Explicit
Dim x
x = 5
`);
      const x = engine.getVariable('x');
      expect(x.value).toBe(5);
    });
  });

  describe('IE compatibility - string functions', () => {
    it('should handle Mid with start beyond string length', () => {
      const engine = new VbsEngine();
      engine.run('x = Mid("hello", 10)');
      const x = engine.getVariable('x');
      expect(x.value).toBe('');
    });

    it('should handle Mid with 1-based index', () => {
      const engine = new VbsEngine();
      engine.run('x = Mid("hello", 1, 2)');
      const x = engine.getVariable('x');
      expect(x.value).toBe('he');
    });

    it('should handle InStr returning 0 when not found', () => {
      const engine = new VbsEngine();
      engine.run('x = InStr("hello", "x")');
      const x = engine.getVariable('x');
      expect(x.value).toBe(0);
    });

    it('should handle InStr with 1-based return', () => {
      const engine = new VbsEngine();
      engine.run('x = InStr("hello", "l")');
      const x = engine.getVariable('x');
      expect(x.value).toBe(3);
    });

    it('should handle InStr with start position', () => {
      const engine = new VbsEngine();
      engine.run('x = InStr(4, "hello", "l")');
      const x = engine.getVariable('x');
      expect(x.value).toBe(4);
    });

    it('should handle Left with length > string', () => {
      const engine = new VbsEngine();
      engine.run('x = Left("hello", 100)');
      const x = engine.getVariable('x');
      expect(x.value).toBe('hello');
    });

    it('should handle Right with length > string', () => {
      const engine = new VbsEngine();
      engine.run('x = Right("hello", 100)');
      const x = engine.getVariable('x');
      expect(x.value).toBe('hello');
    });

    it('should handle Replace function', () => {
      const engine = new VbsEngine();
      engine.run('x = Replace("hello world", "world", "VBScript")');
      const x = engine.getVariable('x');
      expect(x.value).toBe('hello VBScript');
    });

    it('should handle Split function', () => {
      const engine = new VbsEngine();
      engine.run('arr = Split("a,b,c", ",")\nx = arr(1)');
      const x = engine.getVariable('x');
      expect(x.value).toBe('b');
    });

    it('should handle Join function', () => {
      const engine = new VbsEngine();
      engine.run('arr = Array("a", "b", "c")\nx = Join(arr, "-")');
      const x = engine.getVariable('x');
      expect(x.value).toBe('a-b-c');
    });
  });

  describe('IE compatibility - type conversion', () => {
    it('should handle CInt rounding', () => {
      const engine = new VbsEngine();
      engine.run('x = CInt(3.5)\ny = CInt(3.4)');
      const x = engine.getVariable('x');
      const y = engine.getVariable('y');
      expect(x.value).toBe(4);
      expect(y.value).toBe(3);
    });

    it('should handle CBool with numbers', () => {
      const engine = new VbsEngine();
      engine.run('x = CBool(0)\ny = CBool(1)');
      const x = engine.getVariable('x');
      const y = engine.getVariable('y');
      expect(x.value).toBe(false);
      expect(y.value).toBe(true);
    });

    it('should handle CBool with strings', () => {
      const engine = new VbsEngine();
      engine.run('x = CBool("True")\ny = CBool("False")');
      const x = engine.getVariable('x');
      const y = engine.getVariable('y');
      expect(x.value).toBe(true);
      expect(y.value).toBe(false);
    });

    it('should handle CDate with string', () => {
      const engine = new VbsEngine();
      engine.run('x = TypeName(CDate("2024-01-01"))');
      const x = engine.getVariable('x');
      expect(x.value).toBe('Date');
    });
  });

  describe('IE compatibility - Empty/Null/Nothing', () => {
    it('should handle Empty as uninitialized variable', () => {
      const engine = new VbsEngine();
      engine.run('Dim x\nresult = IsEmpty(x)');
      const result = engine.getVariable('result');
      expect(result.value).toBe(true);
    });

    it('should handle Empty in numeric context as 0', () => {
      const engine = new VbsEngine();
      engine.run('Dim x\nresult = x + 5');
      const result = engine.getVariable('result');
      expect(result.value).toBe(5);
    });

    it('should handle Empty in string context as empty string', () => {
      const engine = new VbsEngine();
      engine.run('Dim x\nresult = x & "hello"');
      const result = engine.getVariable('result');
      expect(result.value).toBe('hello');
    });

    it('should handle Null propagation', () => {
      const engine = new VbsEngine();
      engine.run('x = Null\nresult = x + 5');
      const result = engine.getVariable('result');
      expect(result.type).toBe('Null');
    });

    it('should handle IsNull function', () => {
      const engine = new VbsEngine();
      engine.run('x = Null\nresult = IsNull(x)');
      const result = engine.getVariable('result');
      expect(result.value).toBe(true);
    });

    it('should handle IsObject function', () => {
      const engine = new VbsEngine();
      engine.run(`
Class Test
End Class
Dim obj
Set obj = New Test
result = IsObject(obj)
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(true);
    });
  });

  describe('IE compatibility - string comparison', () => {
    it('should compare strings case-insensitively', () => {
      const engine = new VbsEngine();
      engine.run('result = ("HELLO" = "hello")');
      const result = engine.getVariable('result');
      expect(result.value).toBe(true);
    });

    it('should compare strings with < operator', () => {
      const engine = new VbsEngine();
      engine.run('result = ("apple" < "banana")');
      const result = engine.getVariable('result');
      expect(result.value).toBe(true);
    });

    it('should compare strings with > operator', () => {
      const engine = new VbsEngine();
      engine.run('result = ("zebra" > "apple")');
      const result = engine.getVariable('result');
      expect(result.value).toBe(true);
    });

    it('should use StrComp for binary comparison', () => {
      const engine = new VbsEngine();
      engine.run('result = StrComp("HELLO", "hello", 0)');
      const result = engine.getVariable('result');
      expect(result.value).not.toBe(0);
    });
  });

  describe('IE compatibility - error handling', () => {
    it('should handle division by zero', () => {
      const engine = new VbsEngine();
      engine.run(`
On Error Resume Next
x = 1 / 0
result = Err.Number <> 0
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(true);
    });

    it('should handle type mismatch error', () => {
      const engine = new VbsEngine();
      engine.run(`
On Error Resume Next
x = CInt("abc")
result = Err.Number <> 0
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(true);
    });

    it('should handle subscript out of range', () => {
      const engine = new VbsEngine();
      engine.run(`
On Error Resume Next
arr = Array(1, 2, 3)
x = arr(10)
result = Err.Number <> 0
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(true);
    });
  });

  describe('IE compatibility - Set vs Let', () => {
    it('should use Set for object assignment', () => {
      const engine = new VbsEngine();
      engine.run(`
Class Test
    Public Value
End Class

Dim obj1, obj2
Set obj1 = New Test
obj1.Value = 42
Set obj2 = obj1
result = obj2.Value
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(42);
    });

    it('should handle Set with Nothing', () => {
      const engine = new VbsEngine();
      engine.run(`
Class Test
End Class

Dim obj
Set obj = New Test
Set obj = Nothing
result = IsObject(obj)
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(true);
    });

    it('should use Let (implicit) for value assignment', () => {
      const engine = new VbsEngine();
      engine.run(`
x = 42
y = x
result = y
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(42);
    });

    it('should handle property Let vs Set', () => {
      const engine = new VbsEngine();
      engine.run(`
Class Container
    Private strVal
    
    Public Property Let StringValue(v)
        strVal = v
    End Property
    
    Public Property Get StringValue
        StringValue = strVal
    End Property
End Class

Dim c
Set c = New Container
c.StringValue = "hello"
result = c.StringValue
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe('hello');
    });

    it('should handle Private variables in class', () => {
      const engine = new VbsEngine();
      engine.run(`
Class Test
    Private value
    
    Public Sub SetValue(v)
        value = v
    End Sub
    
    Public Function GetValue()
        GetValue = value
    End Function
End Class

Dim t
Set t = New Test
Call t.SetValue(42)
result = t.GetValue()
`);
      const result = engine.getVariable('result');
      expect(result.value).toBe(42);
    });
  });

  describe('browser integration', () => {
    it('should handle CreateObject without ActiveXObject', () => {
      const engine = new VbsEngine();
      
      engine.registerFunction('CreateObject', (cls: VbValue): VbValue => {
        const className = String(cls.value ?? cls);
        if (typeof (globalThis as unknown as { ActiveXObject?: unknown }).ActiveXObject === 'undefined') {
          throw new Error(`ActiveXObject is not supported in this browser environment. Cannot create: '${className}'`);
        }
        return { type: 'Object', value: null };
      });

      expect(() => {
        engine.run(`Set obj = CreateObject("Scripting.FileSystemObject")`);
      }).toThrow('ActiveXObject is not supported');
    });

    it('should handle GetObject without ActiveXObject', () => {
      const engine = new VbsEngine();
      
      engine.registerFunction('GetObject', (pathname?: VbValue, cls?: VbValue): VbValue => {
        const className = cls ? String(cls.value ?? cls) : '';
        if (typeof (globalThis as unknown as { ActiveXObject?: unknown }).ActiveXObject === 'undefined') {
          throw new Error(`ActiveXObject is not supported in this browser environment. Cannot get: '${className}'`);
        }
        return { type: 'Object', value: null };
      });

      expect(() => {
        engine.run(`Set obj = GetObject("test.txt", "Scripting.FileSystemObject")`);
      }).toThrow('ActiveXObject is not supported');
    });
  });
});
