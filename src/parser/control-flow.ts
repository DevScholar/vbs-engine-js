import type {
  Statement,
  BlockStatement,
  Expression,
  IfStatement,
  VbForToStatement,
  VbForEachStatement,
  VbDoLoopStatement,
  VbSelectCaseStatement,
  VbCaseClause,
} from '../ast/index.ts';
import type { Token } from '../lexer/index.ts';
import { ParserState } from './parser-state.ts';
import { ExpressionParser } from './expression-parser.ts';
import { createLocation } from './location.ts';

export class ControlFlowParser {
  constructor(
    private state: ParserState,
    private exprParser: ExpressionParser,
    private parseStatement: () => Statement
  ) {}

  parseIfStatement(): IfStatement {
    const ifToken = this.state.advance();
    const test = this.exprParser.parseExpression();
    this.state.skipOptionalNewlines();
    this.state.expect('Then' as any);

    const hasNewlineAfterThen = this.state.checkNewline();
    this.state.skipOptionalNewlines();
    
    if (hasNewlineAfterThen) {
      return this.parseMultiLineIf(ifToken, test);
    }

    const consequent = this.parseSingleLineIfBody();
    let alternate: Statement | null = null;

    if (this.state.check('Else' as any)) {
      this.state.advance();
      alternate = this.parseSingleLineIfBody();
    }

    if (this.state.check('End' as any)) {
      this.state.advance();
      this.state.expect('If' as any);
    }

    return {
      type: 'IfStatement',
      test,
      consequent,
      alternate,
      loc: createLocation(ifToken, this.state.previous),
    };
  }

  private parseMultiLineIf(ifToken: Token, test: Expression): IfStatement {
    this.state.skipNewlines();

    const consequent = this.parseIfBlock();
    let alternate: Statement | null = null;

    if (this.state.check('ElseIf' as any)) {
      alternate = this.parseElseIfStatement();
    } else if (this.state.check('Else' as any)) {
      this.state.advance();
      this.state.skipNewlines();
      alternate = this.parseIfBlock();
    }

    this.state.expect('End' as any);
    this.state.expect('If' as any);

    return {
      type: 'IfStatement',
      test,
      consequent,
      alternate,
      loc: createLocation(ifToken, this.state.previous),
    };
  }

  parseForStatement(): Statement {
    const forToken = this.state.advance();

    if (this.state.check('Each' as any)) {
      return this.parseForEachStatement(forToken);
    }

    return this.parseForToStatement(forToken);
  }

  parseDoStatement(): VbDoLoopStatement {
    const doToken = this.state.advance();
    let test: Expression | null = null;
    let testPosition: VbDoLoopStatement['testPosition'] = null;

    if (this.state.check('While' as any)) {
      this.state.advance();
      test = this.exprParser.parseExpression();
      testPosition = 'while-do';
    } else if (this.state.check('Until' as any)) {
      this.state.advance();
      test = this.exprParser.parseExpression();
      testPosition = 'until-do';
    }

    this.state.skipNewlines();
    const body = this.parseDoBody();

    if (this.state.check('While' as any)) {
      this.state.advance();
      test = this.exprParser.parseExpression();
      testPosition = 'do-while';
    } else if (this.state.check('Until' as any)) {
      this.state.advance();
      test = this.exprParser.parseExpression();
      testPosition = 'do-until';
    }

    this.state.expect('Loop' as any);

    return {
      type: 'VbDoLoopStatement',
      body,
      test,
      testPosition,
      loc: createLocation(doToken, this.state.previous),
    };
  }

  parseWhileStatement(): Statement {
    const whileToken = this.state.advance();
    const test = this.exprParser.parseExpression();
    this.state.skipNewlines();
    const body = this.parseWhileBody();

    return {
      type: 'WhileStatement',
      test,
      body,
      loc: createLocation(whileToken, this.state.previous),
    };
  }

  parseSelectStatement(): VbSelectCaseStatement {
    const selectToken = this.state.advance();
    this.state.expect('Case' as any);
    const discriminant = this.exprParser.parseExpression();
    this.state.skipNewlines();

    const cases = this.parseCaseClauses();

    this.state.expect('End' as any);
    this.state.expect('Select' as any);

    return {
      type: 'VbSelectCaseStatement',
      discriminant,
      cases,
      loc: createLocation(selectToken, this.state.previous),
    };
  }

  private parseSingleLineIfBody(): BlockStatement {
    const body: Statement[] = [];

    while (!this.state.isEOF) {
      if (this.state.checkNewline()) {
        break;
      }
      
      if (this.state.check('Else' as any) || this.state.check('ElseIf' as any)) {
        break;
      }

      if (this.state.check('End' as any)) {
        break;
      }

      const stmt = this.parseStatement();
      body.push(stmt);
    }

    return {
      type: 'BlockStatement',
      body,
    };
  }

  private parseIfBlock(): BlockStatement {
    const body: Statement[] = [];
    let nestedIfDepth = 0;

    while (!this.state.isEOF) {
      this.state.skipStatementSeparators();

      if (this.state.check('If' as any)) {
        const savedPos = this.state.save();
        this.state.advance();
        while (!this.state.isEOF && !this.state.check('Then' as any)) {
          this.state.advance();
        }
        if (this.state.check('Then' as any)) {
          this.state.advance();
          const isSingleLine = !this.state.checkNewline();
          if (isSingleLine) {
            this.state.restore(savedPos);
            const stmt = this.parseStatement();
            body.push(stmt);
            continue;
          } else {
            this.state.restore(savedPos);
            this.state.advance();
            while (!this.state.isEOF && !this.state.check('Then' as any)) {
              this.state.advance();
            }
            this.state.advance();
            nestedIfDepth++;
            continue;
          }
        }
        this.state.restore(savedPos);
      }

      if (this.state.check('End' as any)) {
        const savedPos = this.state.save();
        this.state.advance();
        if (this.state.check('If' as any)) {
          if (nestedIfDepth > 0) {
            nestedIfDepth--;
            this.state.restore(savedPos);
            const stmt = this.parseStatement();
            body.push(stmt);
            continue;
          }
          this.state.restore(savedPos);
          break;
        }
        this.state.restore(savedPos);
      }

      if (this.state.checkAny('Else' as any, 'ElseIf' as any)) {
        if (nestedIfDepth > 0) {
          const stmt = this.parseStatement();
          body.push(stmt);
          continue;
        }
        break;
      }

      if (this.state.isEOF) break;

      const stmt = this.parseStatement();
      body.push(stmt);
    }

    return {
      type: 'BlockStatement',
      body,
    };
  }

  private parseElseIfStatement(): Statement {
    this.state.advance();
    const test = this.exprParser.parseExpression();
    this.state.skipOptionalNewlines();
    this.state.expect('Then' as any);
    this.state.skipNewlines();

    const consequent = this.parseIfBlock();
    let alternate: Statement | null = null;

    if (this.state.check('ElseIf' as any)) {
      alternate = this.parseElseIfStatement();
    } else if (this.state.check('Else' as any)) {
      this.state.advance();
      this.state.skipNewlines();
      alternate = this.parseIfBlock();
    }

    return {
      type: 'IfStatement',
      test,
      consequent,
      alternate,
      loc: createLocation(this.state.previous, this.state.previous),
    };
  }

  private parseForToStatement(forToken: Token): VbForToStatement {
    const left = this.exprParser.parseIdentifier();
    this.state.expect('Eq' as any);
    const init = this.exprParser.parseExpression();
    this.state.expect('To' as any);
    const to = this.exprParser.parseExpression();

    let step: Expression | null = null;
    if (this.state.match('Step' as any)) {
      step = this.exprParser.parseExpression();
    }

    this.state.skipNewlines();
    const body = this.parseForBody();

    return {
      type: 'VbForToStatement',
      left,
      init,
      to,
      step,
      body,
      loc: createLocation(forToken, this.state.previous),
    };
  }

  private parseForEachStatement(forToken: Token): VbForEachStatement {
    this.state.expect('Each' as any);
    const left = this.exprParser.parseIdentifier();
    this.state.expect('In' as any);
    const right = this.exprParser.parseExpression();
    this.state.skipNewlines();
    const body = this.parseForBody();

    return {
      type: 'VbForEachStatement',
      left,
      right,
      body,
      loc: createLocation(forToken, this.state.previous),
    };
  }

  private parseForBody(): Statement {
    const body: Statement[] = [];

    while (!this.state.isEOF) {
      this.state.skipStatementSeparators();

      if (this.state.check('Next' as any)) {
        break;
      }

      if (this.state.isEOF) break;

      const stmt = this.parseStatement();
      body.push(stmt);
    }

    this.state.expect('Next' as any);
    if (this.state.check('Identifier' as any)) {
      this.state.advance();
    }

    return {
      type: 'BlockStatement',
      body,
    };
  }

  private parseDoBody(): Statement {
    const body: Statement[] = [];

    while (!this.state.isEOF) {
      this.state.skipStatementSeparators();

      if (this.state.checkAny('Loop' as any, 'While' as any, 'Until' as any)) {
        break;
      }

      if (this.state.isEOF) break;

      const stmt = this.parseStatement();
      body.push(stmt);
    }

    return {
      type: 'BlockStatement',
      body,
    };
  }

  private parseWhileBody(): Statement {
    const body: Statement[] = [];

    while (!this.state.isEOF) {
      this.state.skipStatementSeparators();

      if (this.state.check('Wend' as any)) {
        break;
      }

      if (this.state.isEOF) break;

      const stmt = this.parseStatement();
      body.push(stmt);
    }

    this.state.expect('Wend' as any);

    return {
      type: 'BlockStatement',
      body,
    };
  }

  private parseCaseClauses(): VbCaseClause[] {
    const cases: VbCaseClause[] = [];

    while (!this.state.isEOF) {
      this.state.skipStatementSeparators();

      if (this.state.check('End' as any)) {
        break;
      }

      if (this.state.check('Case' as any)) {
        const caseClause = this.parseCaseClause();
        cases.push(caseClause);
      } else {
        break;
      }
    }

    return cases;
  }

  private parseCaseClause(): VbCaseClause {
    const caseToken = this.state.advance();
    let test: Expression | Expression[] | null = null;
    let isElse = false;

    if (this.state.check('Else' as any)) {
      this.state.advance();
      isElse = true;
    } else {
      test = this.parseCaseExpressions();
    }

    this.state.skipNewlines();
    const consequent = this.parseCaseBody();

    return {
      type: 'VbCaseClause',
      test,
      consequent,
      isElse,
      loc: createLocation(caseToken, this.state.previous),
    };
  }

  private parseCaseExpressions(): Expression | Expression[] {
    const expressions: Expression[] = [];

    do {
      if (this.state.check('Is' as any)) {
        this.state.advance();
        const op = this.state.advance();
        const value = this.exprParser.parseExpression();
        expressions.push({
          type: 'BinaryExpression',
          operator: this.getComparisonOperator(op),
          left: { type: 'Identifier', name: '__select_expr__' },
          right: value,
        } as Expression);
      } else {
        expressions.push(this.exprParser.parseExpression());
      }
    } while (this.state.match('Comma' as any));

    return expressions.length === 1 ? expressions[0] : expressions;
  }

  private getComparisonOperator(token: Token): string {
    const opMap: Record<string, string> = {
      'Eq': '==',
      'Lt': '<',
      'Gt': '>',
      'Le': '<=',
      'Ge': '>=',
      'Ne': '!=',
    };
    return opMap[token.type] ?? '==';
  }

  private parseCaseBody(): Statement[] {
    const body: Statement[] = [];

    while (!this.state.isEOF) {
      this.state.skipStatementSeparators();

      if (this.state.check('Case' as any) || this.state.check('End' as any)) {
        break;
      }

      if (this.state.isEOF) break;

      const stmt = this.parseStatement();
      body.push(stmt);
    }

    return body;
  }
}
