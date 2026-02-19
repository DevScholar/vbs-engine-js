import { describe, it, expect } from 'vitest';
import { Lexer, TokenType } from '../lexer/index.ts';

describe('Lexer', () => {
  describe('basic tokens', () => {
    it('should tokenize identifiers', () => {
      const lexer = new Lexer('foo bar _test');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.Identifier);
      expect(tokens[0].value).toBe('foo');
      expect(tokens[1].type).toBe(TokenType.Identifier);
      expect(tokens[1].value).toBe('bar');
      expect(tokens[2].type).toBe(TokenType.Identifier);
      expect(tokens[2].value).toBe('_test');
    });

    it('should tokenize keywords case-insensitively', () => {
      const lexer = new Lexer('DIM Dim dim IF If if');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.Dim);
      expect(tokens[0].value).toBe('dim');
      expect(tokens[1].type).toBe(TokenType.Dim);
      expect(tokens[1].value).toBe('dim');
      expect(tokens[2].type).toBe(TokenType.Dim);
      expect(tokens[2].value).toBe('dim');
    });

    it('should tokenize numbers', () => {
      const lexer = new Lexer('42 3.14 -17');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.NumberLiteral);
      expect(tokens[0].value).toBe('42');
      expect(tokens[1].type).toBe(TokenType.NumberLiteral);
      expect(tokens[1].value).toBe('3.14');
      expect(tokens[2].type).toBe(TokenType.Minus);
      expect(tokens[2].value).toBe('-');
      expect(tokens[3].type).toBe(TokenType.NumberLiteral);
      expect(tokens[3].value).toBe('17');
    });

    it('should tokenize strings', () => {
      const lexer = new Lexer('"hello" "world"');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.StringLiteral);
      expect(tokens[0].value).toBe('hello');
      expect(tokens[1].type).toBe(TokenType.StringLiteral);
      expect(tokens[1].value).toBe('world');
    });

    it('should tokenize operators', () => {
      const lexer = new Lexer('+ - * / \\ = <> < <= > >= &');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.Plus);
      expect(tokens[0].value).toBe('+');
      expect(tokens[1].type).toBe(TokenType.Minus);
      expect(tokens[1].value).toBe('-');
      expect(tokens[2].type).toBe(TokenType.Asterisk);
      expect(tokens[2].value).toBe('*');
      expect(tokens[3].type).toBe(TokenType.Slash);
      expect(tokens[3].value).toBe('/');
      expect(tokens[4].type).toBe(TokenType.Backslash);
      expect(tokens[4].value).toBe('\\');
      expect(tokens[5].type).toBe(TokenType.Eq);
      expect(tokens[5].value).toBe('=');
      expect(tokens[6].type).toBe(TokenType.Ne);
      expect(tokens[6].value).toBe('<>');
      expect(tokens[7].type).toBe(TokenType.Lt);
      expect(tokens[7].value).toBe('<');
      expect(tokens[8].type).toBe(TokenType.Le);
      expect(tokens[8].value).toBe('<=');
      expect(tokens[9].type).toBe(TokenType.Gt);
      expect(tokens[9].value).toBe('>');
      expect(tokens[10].type).toBe(TokenType.Ge);
      expect(tokens[10].value).toBe('>=');
      expect(tokens[11].type).toBe(TokenType.Ampersand);
      expect(tokens[11].value).toBe('&');
    });

    it('should tokenize punctuation', () => {
      const lexer = new Lexer('( ) , . : ;');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.LParen);
      expect(tokens[0].value).toBe('(');
      expect(tokens[1].type).toBe(TokenType.RParen);
      expect(tokens[1].value).toBe(')');
      expect(tokens[2].type).toBe(TokenType.Comma);
      expect(tokens[2].value).toBe(',');
      expect(tokens[3].type).toBe(TokenType.Dot);
      expect(tokens[3].value).toBe('.');
      expect(tokens[4].type).toBe(TokenType.Colon);
      expect(tokens[4].value).toBe(':');
    });
  });

  describe('Vbscript-specific features', () => {
    it('should tokenize date literals', () => {
      const lexer = new Lexer('#1/1/2024#');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.DateLiteral);
      expect(tokens[0].value).toBe('1/1/2024');
    });

    it('should tokenize hex numbers', () => {
      const lexer = new Lexer('&HFF &h10');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.NumberLiteral);
      expect(tokens[0].value).toBe('255');
      expect(tokens[1].type).toBe(TokenType.NumberLiteral);
      expect(tokens[1].value).toBe('16');
    });

    it('should tokenize octal numbers', () => {
      const lexer = new Lexer('&O77 &o10');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.NumberLiteral);
      expect(tokens[0].value).toBe('63');
      expect(tokens[1].type).toBe(TokenType.NumberLiteral);
      expect(tokens[1].value).toBe('8');
    });

    it('should handle line continuations', () => {
      const lexer = new Lexer('x _\r\n  + y');
      const tokens = lexer.tokenize();

      expect(tokens[0].value).toBe('x');
      expect(tokens[1].type).toBe(TokenType.Plus);
      expect(tokens[1].value).toBe('+');
      expect(tokens[2].value).toBe('y');
    });
  });

  describe('edge cases', () => {
    it('should handle empty input', () => {
      const lexer = new Lexer('');
      const tokens = lexer.tokenize();

      expect(tokens.length).toBe(1);
      expect(tokens[0].type).toBe(TokenType.EOF);
    });

    it('should handle whitespace only', () => {
      const lexer = new Lexer('   \t\n  ');
      const tokens = lexer.tokenize();

      expect(tokens.length).toBe(2);
      expect(tokens[0].type).toBe(TokenType.Newline);
      expect(tokens[1].type).toBe(TokenType.EOF);
    });

    it('should handle comments', () => {
      const lexer = new Lexer("x ' this is a comment\ny");
      const tokens = lexer.tokenize();

      expect(tokens[0].value).toBe('x');
      expect(tokens[1].type).toBe(TokenType.Newline);
      expect(tokens[2].value).toBe('y');
    });

    it('should handle Rem comments', () => {
      const lexer = new Lexer('x Rem this is a comment\ny');
      const tokens = lexer.tokenize();

      expect(tokens[0].value).toBe('x');
      expect(tokens[1].type).toBe(TokenType.Newline);
      expect(tokens[2].value).toBe('y');
    });

    it('should tokenize unknown characters', () => {
      const lexer = new Lexer('@ $ #');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.Unknown);
      expect(tokens[0].value).toBe('@');
      expect(tokens[1].type).toBe(TokenType.Unknown);
      expect(tokens[1].value).toBe('$');
      expect(tokens[2].type).toBe(TokenType.DateLiteral);
    });

    it('should handle escaped quotes in strings', () => {
      const lexer = new Lexer('"hello ""world"""');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.StringLiteral);
      expect(tokens[0].value).toBe('hello "world"');
    });

    it('should handle scientific notation', () => {
      const lexer = new Lexer('1e10 2.5E-3');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.NumberLiteral);
      expect(tokens[0].value).toBe('1e10');
      expect(tokens[1].type).toBe(TokenType.NumberLiteral);
      expect(tokens[1].value).toBe('2.5E-3');
    });

    it('should tokenize boolean literals', () => {
      const lexer = new Lexer('True False TRUE FALSE');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.BooleanLiteral);
      expect(tokens[0].value).toBe('true');
      expect(tokens[1].type).toBe(TokenType.BooleanLiteral);
      expect(tokens[1].value).toBe('false');
    });

    it('should tokenize Nothing/Null/Empty literals', () => {
      const lexer = new Lexer('Nothing Null Empty');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.NothingLiteral);
      expect(tokens[1].type).toBe(TokenType.NullLiteral);
      expect(tokens[2].type).toBe(TokenType.EmptyLiteral);
    });

    it('should track token locations', () => {
      const lexer = new Lexer('x\n  y');
      const tokens = lexer.tokenize();

      expect(tokens[0].loc.start.line).toBe(1);
      expect(tokens[0].loc.start.column).toBe(1);
      expect(tokens[1].loc.start.line).toBe(1);
      expect(tokens[2].loc.start.line).toBe(2);
      expect(tokens[2].loc.start.column).toBe(3);
    });
  });

  describe('Vbscript keywords', () => {
    it('should tokenize control flow keywords', () => {
      const lexer = new Lexer('If Then Else ElseIf End If');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.If);
      expect(tokens[1].type).toBe(TokenType.Then);
      expect(tokens[2].type).toBe(TokenType.Else);
      expect(tokens[3].type).toBe(TokenType.ElseIf);
      expect(tokens[4].type).toBe(TokenType.End);
      expect(tokens[5].type).toBe(TokenType.If);
    });

    it('should tokenize loop keywords', () => {
      const lexer = new Lexer('For To Step Next Do Loop While Until');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.For);
      expect(tokens[1].type).toBe(TokenType.To);
      expect(tokens[2].type).toBe(TokenType.Step);
      expect(tokens[3].type).toBe(TokenType.Next);
      expect(tokens[4].type).toBe(TokenType.Do);
      expect(tokens[5].type).toBe(TokenType.Loop);
      expect(tokens[6].type).toBe(TokenType.While);
      expect(tokens[7].type).toBe(TokenType.Until);
    });

    it('should tokenize procedure keywords', () => {
      const lexer = new Lexer('Sub Function End Sub End Function');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.Sub);
      expect(tokens[1].type).toBe(TokenType.Function);
      expect(tokens[2].type).toBe(TokenType.End);
      expect(tokens[3].type).toBe(TokenType.Sub);
      expect(tokens[4].type).toBe(TokenType.End);
      expect(tokens[5].type).toBe(TokenType.Function);
    });

    it('should tokenize class keywords', () => {
      const lexer = new Lexer('Class Property Get Let Set');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.Class);
      expect(tokens[1].type).toBe(TokenType.Property);
      expect(tokens[2].type).toBe(TokenType.Get);
      expect(tokens[3].type).toBe(TokenType.Let);
      expect(tokens[4].type).toBe(TokenType.Set);
    });

    it('should tokenize logical operators', () => {
      const lexer = new Lexer('And Or Not Xor Eqv Imp');
      const tokens = lexer.tokenize();

      expect(tokens[0].type).toBe(TokenType.And);
      expect(tokens[1].type).toBe(TokenType.Or);
      expect(tokens[2].type).toBe(TokenType.Not);
      expect(tokens[3].type).toBe(TokenType.Xor);
      expect(tokens[4].type).toBe(TokenType.Eqv);
      expect(tokens[5].type).toBe(TokenType.Imp);
    });
  });
});
