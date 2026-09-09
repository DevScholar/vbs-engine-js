import { TokenType, type Token, type TokenLocation } from './token.ts';
import { KEYWORDS } from './keywords.ts';

// Real VBScript hex/octal integer-literal semantics (verified against Wine's
// own vbscript.dll conformance suite, dlls/vbscript/tests/lang.vbs - a
// genuine, non-obvious language quirk, not something to guess at): an
// unsuffixed hex/octal literal is interpreted using the NARROWEST width its
// raw bit pattern fits in (16-bit if <= 0xFFFF, else 32-bit), reinterpreting
// the top bit as a sign via two's complement - so `&hffff` (raw 65535, fits
// 16 bits) becomes -1, not 65535. A trailing `&` suffix forces 32-bit (Long)
// width regardless of how the raw value would otherwise fit - so `&hffff&`
// stays 65535 (only negative once the raw value exceeds 0x7FFFFFFF). Prior
// to this fix, the trailing `&` suffix wasn't consumed by the lexer at all,
// leaving it to be re-tokenized as a stray Ampersand (string-concat operator)
// token, cascading into "Unexpected token" parse errors on anything after it.
function signExtendIntegerLiteral(raw: number, forceLong: boolean): number {
  if (forceLong || raw > 0xffff) {
    return raw > 0x7fffffff ? raw - 0x100000000 : raw;
  }
  return raw > 0x7fff ? raw - 0x10000 : raw;
}

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
        if ((this.current as string) === '\n') {
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
      let suffix = '';
      if ((this.current as string) === '&') {
        suffix = this.advance();
      }
      const num = signExtendIntegerLiteral(parseInt(value, 16), suffix === '&');
      return this.createToken(
        TokenType.NumberLiteral,
        String(num),
        start,
        '&' + (this.source[start.offset + 1] === 'H' ? 'H' : 'h') + value + suffix
      );
    }

    if (this.current === '&' && (this.peek === 'o' || this.peek === 'O')) {
      this.advance();
      this.advance();
      while (/[0-7]/.test(this.current)) {
        value += this.advance();
      }
      let suffix = '';
      if ((this.current as string) === '&') {
        suffix = this.advance();
      }
      const num = signExtendIntegerLiteral(parseInt(value, 8), suffix === '&');
      return this.createToken(
        TokenType.NumberLiteral,
        String(num),
        start,
        '&' + (this.source[start.offset + 1] === 'O' ? 'O' : 'o') + value + suffix
      );
    }

    // A bare `&` directly followed by octal digits (no `o`/`O` letter) is
    // ALSO valid octal literal syntax in VBScript - `&100` = octal 100 =
    // decimal 64 (found via Wine's own vbscript.dll conformance suite,
    // dlls/vbscript/tests/lang.vbs, which explicitly comments "Bare '&'
    // followed by octal digits (no 'o'/'O') is octal too"). Previously
    // entirely unhandled - readNumber() only recognized `&h`/`&H` and
    // `&o`/`&O`, so a bare `&100` left the `&` to be tokenized as the
    // Ampersand (string-concat) operator instead, on top of `100` being
    // parsed as a separate, unrelated decimal literal - "Unexpected token:
    // Ampersand" (or worse, silently wrong values, depending on context).
    if (this.current === '&' && /[0-7]/.test(this.peek)) {
      this.advance();
      while (/[0-7]/.test(this.current)) {
        value += this.advance();
      }
      let suffix = '';
      if ((this.current as string) === '&') {
        suffix = this.advance();
      }
      const num = signExtendIntegerLiteral(parseInt(value, 8), suffix === '&');
      return this.createToken(TokenType.NumberLiteral, String(num), start, '&' + value + suffix);
    }

    while (/[0-9]/.test(this.current)) {
      value += this.advance();
    }

    // A digit was already consumed above, so a `.` here is unambiguously
    // part of this number, not a member-access operator (numbers are never
    // valid targets of `.property` in VBScript) - no digits need to follow
    // it: `10.` is valid VBScript for `10.0` (found via Wine's own
    // vbscript.dll conformance suite, dlls/vbscript/tests/lang.vbs:
    // `10. = 10`). Previously required a digit after the dot, so `10.` left
    // the `.` unconsumed, re-tokenized as a stray Dot (member-access)
    // operator and cascading into "Expected property name"/"Expected
    // RParen" parse errors depending on context.
    if (this.current === '.') {
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

      if (this.current === '"') {
        return this.readString(this.current);
      }

      if (this.current === '#') {
        const start = this.getLoc();
        return this.readDate(start);
      }

      if (
        this.current === '&' &&
        (this.peek === 'h' || this.peek === 'H' || this.peek === 'o' || this.peek === 'O' || /[0-7]/.test(this.peek))
      ) {
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
          // Real VBScript accepts `=<`/`=>` as alternate spellings of `<=`/`>=`
          // (the `=` may come before or after the relational operator, both
          // equally valid) - found via Wine's own vbscript.dll conformance
          // suite, dlls/vbscript/tests/lang.vbs: `ok(2 => 1, ...)`, a real
          // assertion in Microsoft-accuracy test code, not a typo.
          if (this.current === '<') {
            this.advance();
            return this.createToken(TokenType.Le, '=<', start);
          }
          if (this.current === '>') {
            this.advance();
            return this.createToken(TokenType.Ge, '=>', start);
          }
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
          // `><` is likewise a valid alternate spelling of `<>` (see the `=<`/`=>`
          // comment above - same "either character order" VB grammar quirk,
          // also found via Wine's lang.vbs: `ok(not (2 >< 2), ...)`).
          if (this.current === '<') {
            this.advance();
            return this.createToken(TokenType.Ne, '><', start);
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
