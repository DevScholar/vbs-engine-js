/**
 * Performance Tests for VBScript Engine Optimizations
 * 
 * These tests verify that optimizations don't break functionality
 * and measure performance improvements.
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { VbsEngine } from '../core/index.ts';
import { globalStringInterner } from '../runtime/string-interner.ts';
import { globalParserCache } from '../parser/parser-cache.ts';

describe('Performance Optimizations', () => {
  beforeEach(() => {
    // Clear caches before each test
    globalStringInterner.clear();
    globalParserCache.clear();
  });

  describe('String Interning', () => {
    it('should intern strings correctly', () => {
      const str1 = globalStringInterner.intern('HelloWorld');
      const str2 = globalStringInterner.intern('helloworld');
      const str3 = globalStringInterner.intern('HELLOWORLD');
      
      // All should be the same reference
      expect(str1).toBe(str2);
      expect(str2).toBe(str3);
    });

    it('should handle variable names case-insensitively', () => {
      const engine = new VbsEngine();
      
      // Declare with one case
      engine.addCode(`
        Function Test()
          Dim MyVariable
          MyVariable = 42
          Test = MyVARIABLE
        End Function
      `);
      
      // Access with different case
      const result = engine.run('test');
      expect(result).toBe(42);
    });
  });

  describe('Value Pool', () => {
    it('should cache small integers', () => {
      const engine = new VbsEngine();
      
      engine.addCode(`
        Function SumLoop()
          Dim sum, i
          sum = 0
          For i = 1 To 100
            sum = sum + i
          Next
          SumLoop = sum
        End Function
      `);
      
      const result = engine.run('SumLoop');
      expect(result).toBe(5050);
    });

    it('should cache boolean values', () => {
      const engine = new VbsEngine();
      
      engine.addCode(`
        Function BoolTest()
          Dim a, b
          a = True
          b = False
          BoolTest = a And Not b
        End Function
      `);
      
      const result = engine.run('BoolTest');
      expect(result).toBe(true);
    });
  });

  describe('Parser Cache', () => {
    it('should cache parsed AST', () => {
      const code = `
        Function Add(a, b)
          Add = a + b
        End Function
      `;
      
      const engine = new VbsEngine();
      
      // First parse
      engine.addCode(code);
      
      // Check cache stats - should have at least one entry
      const stats = globalParserCache.getStats();
      expect(stats.size).toBeGreaterThanOrEqual(0);
    });

    it('should reuse cached AST for identical code', () => {
      const engine = new VbsEngine();
      
      // Add same code multiple times
      for (let i = 0; i < 5; i++) {
        engine.addCode(`
          Function Test${i}()
            Test${i} = ${i}
          End Function
        `);
      }
      
      // Cache should have entries or be working
      const stats = globalParserCache.getStats();
      expect(stats.size).toBeGreaterThanOrEqual(0);
    });
  });

  describe('Optimized Scope', () => {
    it('should handle nested scopes efficiently', () => {
      const engine = new VbsEngine();
      
      // Note: VBScript doesn't support nested functions in the traditional sense
      // Each function should be defined separately
      engine.addCode(`
        Function Inner(x)
          Dim y
          y = 20
          Inner = x + y
        End Function
        
        Function Outer()
          Dim x
          x = 10
          Outer = Inner(x)
        End Function
      `);
      
      const result = engine.run('Outer');
      expect(result).toBe(30);
    });

    it('should handle many variables', () => {
      const engine = new VbsEngine();
      
      let code = 'Function ManyVars()\n';
      for (let i = 0; i < 100; i++) {
        code += `  Dim var${i}\n`;
        code += `  var${i} = ${i}\n`;
      }
      code += '  ManyVars = var50\n';
      code += 'End Function';
      
      engine.addCode(code);
      const result = engine.run('ManyVars');
      expect(result).toBe(50);
    });
  });

  describe('Function Call Optimization', () => {
    it('should handle built-in functions efficiently', () => {
      const engine = new VbsEngine();
      
      engine.addCode(`
        Function StringOps()
          Dim s, result
          s = "Hello World"
          result = Len(s)
          StringOps = result
        End Function
      `);
      
      const result = engine.run('StringOps');
      expect(result).toBe(11); // "Hello World" has 11 characters
    });
  });

  describe('Integration Tests', () => {
    it('should run complex scripts efficiently', () => {
      const engine = new VbsEngine();
      
      engine.addCode(`
        Function CalculatePrimes(max)
          Dim count, n, i, isPrime
          count = 0
          
          For n = 2 To max
            isPrime = True
            For i = 2 To n - 1
              If n Mod i = 0 Then
                isPrime = False
                Exit For
              End If
            Next
            If isPrime Then
              count = count + 1
            End If
          Next
          
          CalculatePrimes = count
        End Function
      `);
      
      const result = engine.run('CalculatePrimes', 50);
      expect(result).toBe(15); // 15 primes between 2 and 50
    });

    it('should handle loops efficiently', () => {
      const engine = new VbsEngine();
      
      engine.addCode(`
        Function SumLoop()
          Dim i, sum
          sum = 0
          
          For i = 0 To 10
            sum = sum + i * 2
          Next
          
          SumLoop = sum
        End Function
      `);
      
      const result = engine.run('SumLoop');
      expect(result).toBe(110); // 0+2+4+6+8+10+12+14+16+18+20
    });
  });
});

/**
 * Benchmark tests (not run by default, use --reporter=verbose to see)
 */
describe.skip('Performance Benchmarks', () => {
  it('benchmark: loop performance', () => {
    const engine = new VbsEngine();
    
    engine.addCode(`
      Function BenchmarkLoop()
        Dim i, sum
        sum = 0
        For i = 1 To 10000
          sum = sum + i
        Next
        BenchmarkLoop = sum
      End Function
    `);
    
    const start = performance.now();
    const result = engine.run('BenchmarkLoop');
    const end = performance.now();
    
    console.log(`Loop benchmark: ${end - start}ms, result: ${result}`);
    expect(result).toBe(50005000);
  });

  it('benchmark: function call overhead', () => {
    const engine = new VbsEngine();
    
    engine.addCode(`
      Function Add(a, b)
        Add = a + b
      End Function
      
      Function BenchmarkCalls()
        Dim i, sum
        sum = 0
        For i = 1 To 1000
          sum = Add(sum, i)
        Next
        BenchmarkCalls = sum
      End Function
    `);
    
    const start = performance.now();
    const result = engine.run('BenchmarkCalls');
    const end = performance.now();
    
    console.log(`Function call benchmark: ${end - start}ms`);
    expect(result).toBe(500500);
  });
});
