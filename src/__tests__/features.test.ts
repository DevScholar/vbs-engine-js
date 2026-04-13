import { describe, it, expect } from 'vitest';
import { VbsEngine } from '../index.ts';

// ---------------------------------------------------------------------------
// VB6 Type Syntax (As Type in declarations and parameters)
// ---------------------------------------------------------------------------
describe('VB6 Type Syntax', () => {
  it('Dim with As Integer initializes to 0', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim x As Integer');
    expect(engine._getVariable('x').type).toBe('Integer');
    expect(engine._getVariable('x').value).toBe(0);
  });

  it('Dim with As Long initializes to 0', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim x As Long');
    expect(engine._getVariable('x').type).toBe('Long');
    expect(engine._getVariable('x').value).toBe(0);
  });

  it('Dim with As String initializes to empty string', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim s As String');
    expect(engine._getVariable('s').type).toBe('String');
    expect(engine._getVariable('s').value).toBe('');
  });

  it('Dim with As Boolean initializes to False', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim b As Boolean');
    expect(engine._getVariable('b').type).toBe('Boolean');
    expect(engine._getVariable('b').value).toBe(false);
  });

  it('Dim with As Double initializes to 0', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim d As Double');
    expect(engine._getVariable('d').type).toBe('Double');
    expect(engine._getVariable('d').value).toBe(0);
  });

  it('Dim with As Byte initializes to 0', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim b As Byte');
    expect(engine._getVariable('b').type).toBe('Byte');
    expect(engine._getVariable('b').value).toBe(0);
  });

  it('Dim with As Object initializes to Nothing', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Set obj = Nothing');
    const v = engine._getVariable('obj');
    expect(v.type).toBe('Object');
    expect(v.value).toBeNull();
  });

  it('Function with typed parameter and return type parses correctly', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Function Add(ByVal a As Integer, ByVal b As Integer) As Long
        Add = a + b
      End Function
      result = Add(3, 4)
    `);
    expect(engine._getVariable('result').value).toBe(7);
  });

  it('Sub with typed parameters works', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Sub Greet(ByVal name As String)
        greeting = "Hello " & name
      End Sub
      Greet("World")
    `);
    expect(engine._getVariable('greeting').value).toBe('Hello World');
  });

  it('Property Get with As Type parses correctly', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Class Counter
        Private m_count As Integer
        Property Get Count() As Integer
          Count = m_count
        End Property
        Property Let Count(ByVal v As Integer)
          m_count = v
        End Property
      End Class
      Set c = New Counter
      c.Count = 5
      result = c.Count
    `);
    expect(engine._getVariable('result').value).toBe(5);
  });
});

// ---------------------------------------------------------------------------
// VB6 Enums
// ---------------------------------------------------------------------------
describe('VB6 Enums', () => {
  it('Enum members are declared as constants with auto-increment from 0', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Enum Color
        Red
        Green
        Blue
      End Enum
    `);
    expect(engine._getVariable('Red').value).toBe(0);
    expect(engine._getVariable('Green').value).toBe(1);
    expect(engine._getVariable('Blue').value).toBe(2);
  });

  it('Enum members with explicit values', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Enum Status
        Pending = 10
        Active = 20
        Inactive = 30
      End Enum
    `);
    expect(engine._getVariable('Pending').value).toBe(10);
    expect(engine._getVariable('Active').value).toBe(20);
    expect(engine._getVariable('Inactive').value).toBe(30);
  });

  it('Enum auto-increment after explicit value', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Enum Priority
        Low
        Medium = 10
        High
        Critical
      End Enum
    `);
    expect(engine._getVariable('Low').value).toBe(0);
    expect(engine._getVariable('Medium').value).toBe(10);
    expect(engine._getVariable('High').value).toBe(11);
    expect(engine._getVariable('Critical').value).toBe(12);
  });

  it('Enum values can be used in expressions', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Enum Direction
        North = 0
        East = 90
        South = 180
        West = 270
      End Enum
      heading = East + 45
    `);
    expect(engine._getVariable('heading').value).toBe(135);
  });

  it('Enum values can be used in Select Case', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Enum Fruit
        Apple
        Banana
        Cherry
      End Enum
      pick = Banana
      Select Case pick
        Case Apple
          result = "apple"
        Case Banana
          result = "banana"
        Case Cherry
          result = "cherry"
      End Select
    `);
    expect(engine._getVariable('result').value).toBe('banana');
  });

  it('Public Enum is accessible globally', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Public Enum Weekday2
        Mon = 1
        Tue = 2
        Wed = 3
      End Enum
      day = Wed
    `);
    expect(engine._getVariable('day').value).toBe(3);
  });
});

// ---------------------------------------------------------------------------
// VB6 Collection
// ---------------------------------------------------------------------------
describe('VB6 Collection', () => {
  it('New Collection creates an empty collection', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      result = col.Count
    `);
    expect(engine._getVariable('result').value).toBe(0);
  });

  it('Collection.Add increases Count', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add "Alpha"
      col.Add "Beta"
      result = col.Count
    `);
    expect(engine._getVariable('result').value).toBe(2);
  });

  it('Collection.Item with 1-based index', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add "First"
      col.Add "Second"
      col.Add "Third"
      r1 = col.Item(1)
      r2 = col.Item(2)
      r3 = col.Item(3)
    `);
    expect(engine._getVariable('r1').value).toBe('First');
    expect(engine._getVariable('r2').value).toBe('Second');
    expect(engine._getVariable('r3').value).toBe('Third');
  });

  it('Collection.Add with string key allows Item lookup by key', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add "Paris", "FR"
      col.Add "London", "UK"
      col.Add "Berlin", "DE"
      r1 = col.Item("FR")
      r2 = col.Item("UK")
      r3 = col.Item("DE")
    `);
    expect(engine._getVariable('r1').value).toBe('Paris');
    expect(engine._getVariable('r2').value).toBe('London');
    expect(engine._getVariable('r3').value).toBe('Berlin');
  });

  it('Collection.Remove by index', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add "A"
      col.Add "B"
      col.Add "C"
      col.Remove 2
      result = col.Count
      r1 = col.Item(1)
      r2 = col.Item(2)
    `);
    expect(engine._getVariable('result').value).toBe(2);
    expect(engine._getVariable('r1').value).toBe('A');
    expect(engine._getVariable('r2').value).toBe('C');
  });

  it('Collection.Remove by key', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add "Paris", "FR"
      col.Add "London", "UK"
      col.Remove "FR"
      result = col.Count
      r1 = col.Item(1)
    `);
    expect(engine._getVariable('result').value).toBe(1);
    expect(engine._getVariable('r1').value).toBe('London');
  });

  it('For Each iterates all Collection items', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add 10
      col.Add 20
      col.Add 30
      total = 0
      For Each item In col
        total = total + item
      Next
    `);
    expect(engine._getVariable('total').value).toBe(60);
  });

  it('Collection stores mixed types', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      col.Add 42
      col.Add "hello"
      col.Add True
      n = col.Item(1)
      s = col.Item(2)
      b = col.Item(3)
    `);
    expect(engine._getVariable('n').value).toBe(42);
    expect(engine._getVariable('s').value).toBe('hello');
    expect(engine._getVariable('b').value).toBe(true);
  });

  it('TypeName of Collection is "Collection"', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Set col = New Collection
      result = TypeName(col)
    `);
    expect(engine._getVariable('result').value).toBe('Collection');
  });
});

// ---------------------------------------------------------------------------
// VBA7 LongLong type and 64-bit integer support
// ---------------------------------------------------------------------------
describe('VBA7 LongLong type', () => {
  it('Dim with As LongLong initializes to BigInt(0)', () => {
    const engine = new VbsEngine();
    engine.executeStatement('Dim x As LongLong');
    const v = engine._getVariable('x');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(0));
  });

  it('CLngLng converts integer to LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = CLngLng(42)');
    const v = engine._getVariable('x');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(42));
  });

  it('CLngLng converts string to LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = CLngLng("9223372036854775807")');
    const v = engine._getVariable('x');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt('9223372036854775807'));
  });

  it('CLngLng converts negative value', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = CLngLng(-100)');
    const v = engine._getVariable('x');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(-100));
  });

  it('LongLong + LongLong = LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(1000000000)
      b = CLngLng(2000000000)
      c = a + b
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(3000000000));
  });

  it('LongLong - LongLong = LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(5000000000)
      b = CLngLng(3000000000)
      c = a - b
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(2000000000));
  });

  it('LongLong * LongLong = LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(1000000)
      b = CLngLng(1000000)
      c = a * b
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(1000000000000));
  });

  it('LongLong + Integer = LongLong', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(9000000000)
      c = a + 1
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(9000000001));
  });

  it('LongLong + Double promotes to Double', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(100)
      c = a + 1.5
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('Double');
    expect(v.value).toBeCloseTo(101.5);
  });

  it('LongLong comparison works', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(9000000000)
      b = CLngLng(9000000000)
      If a = b Then eq = True Else eq = False
      If CLngLng(1) < CLngLng(2) Then lt = True Else lt = False
    `);
    expect(engine._getVariable('eq').value).toBe(true);
    expect(engine._getVariable('lt').value).toBe(true);
  });

  it('VarType of LongLong is 20', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = VarType(CLngLng(0))');
    expect(engine._getVariable('x').value).toBe(20);
  });

  it('TypeName of LongLong is "LongLong"', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = TypeName(CLngLng(0))');
    expect(engine._getVariable('x').value).toBe('LongLong');
  });

  it('CStr converts LongLong to string', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = CStr(CLngLng("9223372036854775807"))');
    expect(engine._getVariable('x').value).toBe('9223372036854775807');
  });

  it('LongLong integer division (\\)', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      a = CLngLng(10000000000)
      b = CLngLng(3)
      c = a \\ b
    `);
    const v = engine._getVariable('c');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt(3333333333));
  });

  it('64-bit range: max LongLong via string', () => {
    const engine = new VbsEngine();
    engine.executeStatement('x = CLngLng("9223372036854775807")');
    const v = engine._getVariable('x');
    expect(v.type).toBe('LongLong');
    expect(v.value).toBe(BigInt('9223372036854775807'));
  });
});

// ---------------------------------------------------------------------------
// VB6 User-Defined Types (Type...End Type)
// ---------------------------------------------------------------------------
describe('VB6 User-Defined Types', () => {
  it('Type declaration and Dim creates instance with default-initialized fields', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Point
        X As Integer
        Y As Integer
      End Type
      Dim p As Point
    `);
    const p = engine._getVariable('p');
    expect(p.type).toBe('Object');
    const instance = p.value as import('../runtime/class-registry.ts').VbObjectInstance;
    expect(instance.getProperty('X').type).toBe('Integer');
    expect(instance.getProperty('X').value).toBe(0);
    expect(instance.getProperty('Y').type).toBe('Integer');
    expect(instance.getProperty('Y').value).toBe(0);
  });

  it('Field access and assignment', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Point
        X As Integer
        Y As Integer
      End Type
      Dim p As Point
      p.X = 10
      p.Y = 20
      result = p.X + p.Y
    `);
    expect(engine._getVariable('result').value).toBe(30);
  });

  it('String and Boolean fields', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Person
        Name As String
        Active As Boolean
        Age As Integer
      End Type
      Dim per As Person
      per.Name = "Alice"
      per.Active = True
      per.Age = 30
    `);
    const per = engine._getVariable('per');
    const instance = per.value as import('../runtime/class-registry.ts').VbObjectInstance;
    expect(instance.getProperty('Name').value).toBe('Alice');
    expect(instance.getProperty('Active').value).toBe(true);
    expect(instance.getProperty('Age').value).toBe(30);
  });

  it('Multiple independent instances', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Point
        X As Integer
        Y As Integer
      End Type
      Dim a As Point
      Dim b As Point
      a.X = 1
      b.X = 99
    `);
    const a = engine._getVariable('a').value as import('../runtime/class-registry.ts').VbObjectInstance;
    const b = engine._getVariable('b').value as import('../runtime/class-registry.ts').VbObjectInstance;
    expect(a.getProperty('X').value).toBe(1);
    expect(b.getProperty('X').value).toBe(99);
  });

  it('Fields default to type-appropriate zero values', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Defaults
        N As Long
        S As String
        B As Boolean
        D As Double
      End Type
      Dim d As Defaults
    `);
    const instance = engine._getVariable('d').value as import('../runtime/class-registry.ts').VbObjectInstance;
    expect(instance.getProperty('N').value).toBe(0);
    expect(instance.getProperty('S').value).toBe('');
    expect(instance.getProperty('B').value).toBe(false);
    expect(instance.getProperty('D').value).toBe(0);
  });

  it('UDT used in Sub parameter (ByRef)', () => {
    const engine = new VbsEngine();
    engine.executeStatement(`
      Type Vec
        X As Double
        Y As Double
      End Type
      Sub Scale(v As Vec, factor As Double)
        v.X = v.X * factor
        v.Y = v.Y * factor
      End Sub
      Dim v As Vec
      v.X = 3
      v.Y = 4
      Scale v, 2
      rx = v.X
      ry = v.Y
    `);
    expect(engine._getVariable('rx').value).toBe(6);
    expect(engine._getVariable('ry').value).toBe(8);
  });
});
