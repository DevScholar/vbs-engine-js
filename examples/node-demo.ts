import { VbsEngine } from '../src/index.ts';

const engine = new VbsEngine();

engine.run(`
  ' Variables
  name = "World"
  
  ' Function
  Function Greet(n)
      Greet = "Hello, " & n & "!"
  End Function
  
  ' Array
  arr = Array(1, 2, 3, 4, 5)
  
  ' Loop
  sum = 0
  For i = 0 To 4
      sum = sum + arr(i)
  Next
`);

console.log('name:', engine.getVariableAsJs('name'));
console.log('sum:', engine.getVariableAsJs('sum'));
console.log('arr:', engine.getVariableAsJs('arr'));
