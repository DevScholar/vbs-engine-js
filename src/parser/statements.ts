import type {
  Statement,
  Expression,
  VbWithStatement,
  VbExitStatement,
  VbOptionExplicitStatement,
} from '../ast/index.ts';
import { ParserState } from './parser-state.ts';
import { ExpressionParser } from './expression-parser.ts';
import { createLocation } from './location.ts';

export class StatementsParser {
  constructor(
    private state: ParserState,
    private exprParser: ExpressionParser,
    private parseStatement: () => Statement
  ) {}

  parseWithStatement(): VbWithStatement {
    const withToken = this.state.advance();
    const object = this.exprParser.parseExpression();
    this.state.skipNewlines();
    const body = this.parseWithBody();

    this.state.expect('End' as any);
    this.state.expect('With' as any);

    return {
      type: 'VbWithStatement',
      object,
      body,
      loc: createLocation(withToken, this.state.previous),
    };
  }

  parseExitStatement(): VbExitStatement {
    const exitToken = this.state.advance();

    const targetTypes: Record<string, VbExitStatement['target']> = {
      'Sub': 'Sub',
      'Function': 'Function',
      'Property': 'Property',
      'Do': 'Do',
      'For': 'For',
      'Select': 'Select',
    };

    const targetToken = this.state.current;
    const target = targetTypes[targetToken.type as string];

    if (target) {
      this.state.advance();
    }

    return {
      type: 'VbExitStatement',
      target: target ?? 'Sub',
      loc: createLocation(exitToken, this.state.previous),
    };
  }

  parseOptionStatement(): VbOptionExplicitStatement {
    const optionToken = this.state.advance();
    this.state.expect('Explicit' as any);

    return {
      type: 'VbOptionExplicitStatement',
      loc: createLocation(optionToken, this.state.previous),
    };
  }

  parseCallStatement(): Statement {
    const callToken = this.state.advance();
    const callee = this.exprParser.parseMemberExpression();

    let args: Expression[] = [];
    if (this.state.check('LParen' as any)) {
      this.state.advance();
      args = this.parseCallArguments();
      this.state.expect('RParen' as any);
    } else if (!this.state.checkAny('Newline' as any, 'Colon' as any, 'EOF' as any)) {
      args = this.parseCallArgumentsNoParens();
    }

    return {
      type: 'VbCallStatement',
      callee,
      arguments: args,
      loc: createLocation(callToken, this.state.previous),
    };
  }

  parseSetStatement(): Statement {
    const setToken = this.state.advance();
    const left = this.exprParser.parseCallExpression();
    this.state.expect('Eq' as any);
    const right = this.exprParser.parseExpression();

    const assignment = {
      type: 'AssignmentExpression',
      operator: '=',
      left: left as Expression,
      right,
      isSet: true,
      loc: createLocation(setToken, this.state.previous),
    } as Expression;

    return {
      type: 'ExpressionStatement',
      expression: assignment,
      loc: createLocation(setToken, this.state.previous),
    };
  }

  parseOnStatement(): Statement {
    const onToken = this.state.advance();
    this.state.expect('Error' as any);

    if (this.state.match('Resume' as any)) {
      this.state.expect('Next' as any);
      return {
        type: 'VbOnErrorHandlerStatement',
        action: 'resume_next',
        loc: createLocation(onToken, this.state.previous),
      };
    }

    if (this.state.match('Goto' as any)) {
      if (this.state.check('NumberLiteral' as any)) {
        const numToken = this.state.advance();
        if (String(numToken.value) === '0') {
          return {
            type: 'VbOnErrorHandlerStatement',
            action: 'goto_0',
            loc: createLocation(onToken, this.state.previous),
          };
        }
      }
      const label = this.exprParser.parseIdentifier();
      return {
        type: 'VbOnErrorHandlerStatement',
        action: 'goto_label',
        label,
        loc: createLocation(onToken, this.state.previous),
      };
    }

    throw new Error('Expected Resume Next or Goto after On Error');
  }

  parseResumeStatement(): Statement {
    const resumeToken = this.state.advance();

    if (this.state.match('Next' as any)) {
      return {
        type: 'VbResumeStatement',
        target: 'next',
        loc: createLocation(resumeToken, this.state.previous),
      };
    }

    return {
      type: 'VbResumeStatement',
      target: null,
      loc: createLocation(resumeToken, this.state.previous),
    };
  }

  parseExpressionStatement(): Statement {
    const expr = this.exprParser.parseStatementExpression();

    return {
      type: 'ExpressionStatement',
      expression: expr,
      loc: expr.loc,
    };
  }

  private parseWithBody(): Statement {
    const body: Statement[] = [];

    while (!this.state.isEOF) {
      this.state.skipNewlines();

      if (this.state.check('End' as any) && this.state.peek(1).type === 'With' as any) {
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

  private parseCallArguments(): Expression[] {
    const args: Expression[] = [];

    if (!this.state.check('RParen' as any)) {
      while (true) {
        this.state.skipOptionalNewlines();
        
        if (this.state.check('RParen' as any)) {
          break;
        }
        
        // Handle empty arguments (consecutive commas)
        if (this.state.check('Comma' as any)) {
          args.push({
            type: 'VbEmptyLiteral',
            value: undefined,
            raw: '',
            loc: this.state.current.loc,
          });
        } else {
          args.push(this.exprParser.parseExpression());
        }
        
        this.state.skipOptionalNewlines();
        
        if (this.state.check('Comma' as any)) {
          this.state.advance();
        } else {
          break;
        }
      }
    }

    return args;
  }

  private parseCallArgumentsNoParens(): Expression[] {
    const args: Expression[] = [];

    while (!this.state.checkAny('Newline' as any, 'Colon' as any, 'EOF' as any)) {
      args.push(this.exprParser.parseExpression());
      if (!this.state.match('Comma' as any)) {
        break;
      }
    }

    return args;
  }
}
