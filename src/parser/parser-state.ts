import type { Token, TokenType } from '../lexer/index.ts';

export class ParseError extends Error {
  public token: Token;
  public expected?: TokenType | TokenType[];

  constructor(message: string, token: Token, expected?: TokenType | TokenType[]) {
    super(message);
    this.name = 'ParseError';
    this.token = token;
    this.expected = expected;
  }
}

export class ParserState {
  private tokens: Token[];
  private pos: number = 0;

  constructor(tokens: Token[]) {
    this.tokens = tokens;
  }

  get current(): Token {
    return this.tokens[this.pos] ?? this.tokens[this.tokens.length - 1];
  }

  get previous(): Token {
    return this.tokens[this.pos - 1] ?? this.tokens[0];
  }

  get isEOF(): boolean {
    return this.current.type === 'EOF' as TokenType;
  }

  save(): number {
    return this.pos;
  }

  restore(pos: number): void {
    this.pos = pos;
  }

  advance(): Token {
    if (!this.isEOF) {
      this.pos++;
    }
    return this.previous;
  }

  peek(offset: number = 0): Token {
    return this.tokens[this.pos + offset] ?? this.tokens[this.tokens.length - 1];
  }

  check(type: TokenType): boolean {
    return this.current.type === type;
  }

  checkIdentifier(): boolean {
    return this.current.type === 'Identifier' as TokenType;
  }

  checkAny(...types: TokenType[]): boolean {
    return types.includes(this.current.type);
  }

  checkNewline(): boolean {
    return this.current.type === 'Newline' as TokenType;
  }

  match(type: TokenType): Token | null {
    if (this.check(type)) {
      return this.advance();
    }
    return null;
  }

  matchAny(...types: TokenType[]): Token | null {
    if (this.checkAny(...types)) {
      return this.advance();
    }
    return null;
  }

  expect(type: TokenType, message?: string): Token {
    if (this.check(type)) {
      return this.advance();
    }
    throw new ParseError(
      message ?? `Expected ${type}, got ${this.current.type}`,
      this.current,
      type
    );
  }

  expectAny(types: TokenType[], message?: string): Token {
    if (this.checkAny(...types)) {
      return this.advance();
    }
    throw new ParseError(
      message ?? `Expected one of ${types.join(', ')}, got ${this.current.type}`,
      this.current,
      types
    );
  }

  skipNewlines(): void {
    while (this.match('Newline' as TokenType)) {}
  }

  skipOptionalNewlines(): void {
    this.skipNewlines();
  }
}
