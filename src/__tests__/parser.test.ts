import { describe, it, expect } from 'vitest';
import { parse } from '../parser/index.ts';

describe('Parser', () => {
  describe('expressions', () => {
    it('should parse simple expressions', () => {
      const ast = parse('x = 42');
      expect(ast.body[0].type).toBe('ExpressionStatement');
    });

    it('should parse arithmetic expressions', () => {
      const ast = parse('x = 1 + 2 * 3');
      expect(ast.body[0].type).toBe('ExpressionStatement');
    });

    it('should parse string concatenation', () => {
      const ast = parse('x = "hello" & " world"');
      expect(ast.body[0].type).toBe('ExpressionStatement');
      const stmt = ast.body[0] as { type: string; expression: { type: string; right: { type: string; operator: string } } };
      expect(stmt.expression.type).toBe('AssignmentExpression');
      expect(stmt.expression.right.type).toBe('BinaryExpression');
      expect(stmt.expression.right.operator).toBe('&');
    });

    it('should parse function calls', () => {
      const ast = parse('result = MyFunction(1, 2, 3)');
      expect(ast.body[0].type).toBe('ExpressionStatement');
    });

    it('should parse member expressions', () => {
      const ast = parse('x = obj.Property');
      expect(ast.body[0].type).toBe('ExpressionStatement');
    });

    it('should parse comparison expressions', () => {
      const ast = parse('If x > 10 Then y = 1 End If');
      expect(ast.body[0].type).toBe('IfStatement');
    });
  });

  describe('statements', () => {
    it('should parse Dim statement', () => {
      const ast = parse('Dim x');
      expect(ast.body[0].type).toBe('VbDimStatement');
    });

    it('should parse Dim with multiple variables', () => {
      const ast = parse('Dim x, y, z');
      const dimStmt = ast.body[0] as { type: string; declarations: unknown[] };
      expect(dimStmt.type).toBe('VbDimStatement');
      expect(dimStmt.declarations).toHaveLength(3);
    });

    it('should parse If-Then-Else', () => {
      const ast = parse('If x > 0 Then y = 1 Else y = 2 End If');
      expect(ast.body[0].type).toBe('IfStatement');
    });

    it('should parse For-Next loop', () => {
      const ast = parse('For i = 1 To 10 Step 2\nx = x + i\nNext');
      expect(ast.body[0].type).toBe('VbForToStatement');
    });

    it('should parse Do-Loop', () => {
      const ast = parse('Do While x < 10\nx = x + 1\nLoop');
      expect(ast.body[0].type).toBe('VbDoLoopStatement');
    });

    it('should parse Sub definition', () => {
      const ast = parse('Sub MySub(x, y)\nEnd Sub');
      expect(ast.body[0].type).toBe('VbSubStatement');
    });

    it('should parse Function definition', () => {
      const ast = parse('Function MyFunc(x)\nMyFunc = x * 2\nEnd Function');
      expect(ast.body[0].type).toBe('VbFunctionStatement');
    });

    it('should parse Class definition', () => {
      const ast = parse('Class MyClass\nEnd Class');
      expect(ast.body[0].type).toBe('VbClassStatement');
    });
  });
});
