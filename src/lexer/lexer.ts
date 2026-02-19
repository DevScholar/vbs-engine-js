import { Token, TokenType, TokenLocation } from './token.ts';
import { KEYWORDS } from './keywords.ts';

export interface LexerOptions {
  skipWhitespace?: boolean;
  skipNewlines?: boolean;
}

export class Lexer {
  private source: string;
  private pos: number = 0;
  private line: number = 1;
  private column: number = 1;
  private options: LexerOptions;

  constructor(source: string, options: LexerOptions = {}) {
    this.source = source;
    this.options = {
      skipWhitespace: true,
      skipNewlines: false,
      ...options,
    };
  }

  private get current(): string {
    return this.source[this.pos] ?? '';
  }

  private get peek(): string {
    return this.source[this.pos + 1] ?? '';
  }

  private get isEOF(): boolean {
    return this.pos >= this.source.length;
  }

  private getLoc(): TokenLocation {
    return {
      line: this.line,
      column: this.column,
      offset: this.pos,
    };
  }

  private advance(): string {
    const char = this.current;
    this.pos++;
    if (char === '\n') {
      this.line++;
      this.column = 1;
    } else {
      this.column++;
    }
    return char;
  }

  private createToken(type: TokenType, value: string, start: TokenLocation, raw?: string): Token {
    return {
      type,
      value,
      loc: {
        start,
        end: this.getLoc(),
      },
      raw,
    };
  }

  private skipWhitespaceAndLineContinuation(): void {
    while (true) {
      while (this.current === ' ' || this.current === '\t' || this.current === '\r') {
        this.advance();
      }
      
      if (this.current === '_' && (this.peek === '\n' || this.peek === '\r')) {
        this.advance();
        if (this.peek === '\r') {
          this.advance();
        }
        if (this.peek === '\n') {
          this.advance();
        }
        continue;
      }
      
      break;
    }
  }

  private readString(quote: string): Token {
    const start = this.getLoc();
    this.advance();
    
    let value = '';
    while (!this.isEOF) {
      if (this.current === quote) {
        if (this.peek === quote) {
          value += quote;
          this.advance();
          this.advance();
        } else {
          break;
        }
      } else if (this.current === '\n') {
        break;
      } else {
        value += this.advance();
      }
    }
    
    if (this.current === quote) {
      this.advance();
    }
    
    return this.createToken(TokenType.StringLiteral, value, start, quote + value + quote);
  }

  private readNumber(): Token {
    const start = this.getLoc();
    let value = '';
    let isFloat = false;
    let isExponent = false;
    
    if (this.current === '&' && (this.peek === 'h' || this.peek === 'H')) {
      this.advance();
      this.advance();
      while (/[0-9a-fA-F]/.test(this.current)) {
        value += this.advance();
      }
      const num = parseInt(value, 16);
      return this.createToken(TokenType.NumberLiteral, String(num), start, '&' + (this.source[start.offset + 1] === 'H' ? 'H' : 'h') + value);
    }
    
    if (this.current === '&' && (this.peek === 'o' || this.peek === 'O')) {
      this.advance();
      this.advance();
      while (/[0-7]/.test(this.current)) {
        value += this.advance();
      }
      const num = parseInt(value, 8);
      return this.createToken(TokenType.NumberLiteral, String(num), start, '&' + (this.source[start.offset + 1] === 'O' ? 'O' : 'o') + value);
    }
    
    while (/[0-9]/.test(this.current)) {
      value += this.advance();
    }
    
    if (this.current === '.' && /[0-9]/.test(this.peek)) {
      isFloat = true;
      value += this.advance();
      while (/[0-9]/.test(this.current)) {
        value += this.advance();
      }
    }
    
    const currentChar = this.current;
    if (currentChar === 'e' || currentChar === 'E') {
      isExponent = true;
      value += this.advance();
      const signChar = this.current;
      if (signChar === '+' || signChar === '-') {
        value += this.advance();
      }
      while (/[0-9]/.test(this.current)) {
        value += this.advance();
      }
    }
    
    if (this.current === '#' && !isFloat && !isExponent) {
      return this.readDate(start, value);
    }
    
    return this.createToken(TokenType.NumberLiteral, value, start, value);
  }

  private readDate(start: TokenLocation, prefix?: string): Token {
    let value = prefix ?? '';
    if (this.current === '#') {
      this.advance();
    }
    
    while (!this.isEOF && this.current !== '#') {
      if (this.current === '\n') break;
      value += this.advance();
    }
    
    if (this.current === '#') {
      this.advance();
    }
    
    return this.createToken(TokenType.DateLiteral, value.trim(), start, '#' + value + '#');
  }

  private readIdentifier(): Token {
    const start = this.getLoc();
    let value = '';
    
    while (/[a-zA-Z0-9_]/.test(this.current)) {
      value += this.advance();
    }
    
    const upperValue = value.toLowerCase();
    const keywordType = KEYWORDS[upperValue];
    
    if (keywordType) {
      if (keywordType === TokenType.BooleanLiteral) {
        return this.createToken(TokenType.BooleanLiteral, upperValue, start, value);
      }
      if (keywordType === TokenType.NothingLiteral) {
        return this.createToken(TokenType.NothingLiteral, 'nothing', start, value);
      }
      if (keywordType === TokenType.NullLiteral) {
        return this.createToken(TokenType.NullLiteral, 'null', start, value);
      }
      if (keywordType === TokenType.EmptyLiteral) {
        return this.createToken(TokenType.EmptyLiteral, 'empty', start, value);
      }
      return this.createToken(keywordType, upperValue, start, value);
    }
    
    return this.createToken(TokenType.Identifier, value, start, value);
  }

  private readRemComment(): Token {
    const start = this.getLoc();
    let value = '';
    
    while (!this.isEOF && this.current !== '\n') {
      value += this.advance();
    }
    
    return this.createToken(TokenType.Rem, value, start, value);
  }

  private readSingleLineComment(): void {
    while (!this.isEOF && this.current !== '\n') {
      this.advance();
    }
  }

  nextToken(): Token {
    while (!this.isEOF) {
      if (this.options.skipWhitespace) {
        this.skipWhitespaceAndLineContinuation();
      }
      
      if (this.isEOF) {
        return this.createToken(TokenType.EOF, '', this.getLoc());
      }
      
      if (this.current === '\n') {
        const start = this.getLoc();
        this.advance();
        if (this.options.skipNewlines) {
          continue;
        }
        return this.createToken(TokenType.Newline, '\n', start, '\n');
      }
      
      if (this.current === "'" && this.options.skipWhitespace) {
        this.readSingleLineComment();
        continue;
      }
      
      if (this.current === '"' || this.current === "'") {
        return this.readString(this.current);
      }
      
      if (this.current === '#') {
        const start = this.getLoc();
        return this.readDate(start);
      }
      
      if (this.current === '&' && (this.peek === 'h' || this.peek === 'H' || this.peek === 'o' || this.peek === 'O')) {
        return this.readNumber();
      }
      
      if (/[0-9]/.test(this.current)) {
        return this.readNumber();
      }
      
      if (this.current === '.' && /[0-9]/.test(this.peek)) {
        return this.readNumber();
      }
      
      if (/[a-zA-Z_]/.test(this.current)) {
        const token = this.readIdentifier();
        if (token.type === TokenType.Rem) {
          if (this.options.skipWhitespace) {
            this.readRemComment();
            continue;
          }
          return token;
        }
        return token;
      }
      
      const start = this.getLoc();
      const char = this.current;
      
      switch (char) {
        case '+':
          this.advance();
          return this.createToken(TokenType.Plus, '+', start);
        case '-':
          this.advance();
          return this.createToken(TokenType.Minus, '-', start);
        case '*':
          this.advance();
          return this.createToken(TokenType.Asterisk, '*', start);
        case '/':
          this.advance();
          return this.createToken(TokenType.Slash, '/', start);
        case '\\':
          this.advance();
          return this.createToken(TokenType.Backslash, '\\', start);
        case '^':
          this.advance();
          return this.createToken(TokenType.Caret, '^', start);
        case '&':
          this.advance();
          return this.createToken(TokenType.Ampersand, '&', start);
        case '(':
          this.advance();
          return this.createToken(TokenType.LParen, '(', start);
        case ')':
          this.advance();
          return this.createToken(TokenType.RParen, ')', start);
        case '{':
          this.advance();
          return this.createToken(TokenType.LBrace, '{', start);
        case '}':
          this.advance();
          return this.createToken(TokenType.RBrace, '}', start);
        case '[':
          this.advance();
          return this.createToken(TokenType.LBracket, '[', start);
        case ']':
          this.advance();
          return this.createToken(TokenType.RBracket, ']', start);
        case ',':
          this.advance();
          return this.createToken(TokenType.Comma, ',', start);
        case ':':
          this.advance();
          return this.createToken(TokenType.Colon, ':', start);
        case '.':
          this.advance();
          return this.createToken(TokenType.Dot, '.', start);
        case '!':
          this.advance();
          return this.createToken(TokenType.Bang, '!', start);
        case '=':
          this.advance();
          return this.createToken(TokenType.Eq, '=', start);
        case '<':
          this.advance();
          if (this.current === '>') {
            this.advance();
            return this.createToken(TokenType.Ne, '<>', start);
          }
          if (this.current === '=') {
            this.advance();
            return this.createToken(TokenType.Le, '<=', start);
          }
          return this.createToken(TokenType.Lt, '<', start);
        case '>':
          this.advance();
          if (this.current === '=') {
            this.advance();
            return this.createToken(TokenType.Ge, '>=', start);
          }
          return this.createToken(TokenType.Gt, '>', start);
        default:
          this.advance();
          return this.createToken(TokenType.Unknown, char, start);
      }
    }
    
    return this.createToken(TokenType.EOF, '', this.getLoc());
  }

  tokenize(): Token[] {
    const tokens: Token[] = [];
    
    while (true) {
      const token = this.nextToken();
      tokens.push(token);
      
      if (token.type === TokenType.EOF) {
        break;
      }
    }
    
    return tokens;
  }
}

export function tokenize(source: string, options?: LexerOptions): Token[] {
  const lexer = new Lexer(source, options);
  return lexer.tokenize();
}
