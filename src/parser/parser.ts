import type { Program } from '../ast/index.ts';
import type { Token } from '../lexer/index.ts';
import { Lexer } from '../lexer/index.ts';
import { ParserState } from './parser-state.ts';
import { StatementParser } from './statement-parser.ts';

export interface ParserOptions {
  source?: string;
}

export class Parser {
  private state: ParserState;
  private stmtParser: StatementParser;

  constructor(tokens: Token[]) {
    this.state = new ParserState(tokens);
    this.stmtParser = new StatementParser(this.state);
  }

  parse(): Program {
    const body = [];

    while (!this.state.isEOF) {
      this.state.skipNewlines();
      if (this.state.isEOF) break;

      const stmt = this.stmtParser.parseStatement();
      body.push(stmt);
    }

    const lastToken = this.state.current;

    return {
      type: 'Program',
      body,
      sourceType: 'script',
      loc: {
        start: { line: 1, column: 0 },
        end: lastToken.loc.end,
      },
    };
  }
}

export function parse(source: string, _options?: ParserOptions): Program {
  const lexer = new Lexer(source);
  const tokens = lexer.tokenize();
  const parser = new Parser(tokens);
  return parser.parse();
}
