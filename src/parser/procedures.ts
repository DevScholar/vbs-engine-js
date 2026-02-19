import type {
  Statement,
  BlockStatement,
  VbSubStatement,
  VbFunctionStatement,
  VbClassStatement,
  VbClassElement,
  VbParameter,
  VbPropertyGetStatement,
  VbPropertyLetStatement,
  VbPropertySetStatement,
} from '../ast/index.ts';
import type { Token } from '../lexer/index.ts';
import { ParserState } from './parser-state.ts';
import { ExpressionParser } from './expression-parser.ts';
import { createLocation } from './location.ts';

function isVbClassElement(stmt: Statement): stmt is VbClassElement {
  const elementTypes = [
    'VbSubStatement',
    'VbFunctionStatement',
    'VbPropertyGetStatement',
    'VbPropertyLetStatement',
    'VbPropertySetStatement',
    'VbDimStatement',
    'VbConstStatement',
  ];
  return elementTypes.includes(stmt.type);
}

export class ProcedureParser {
  constructor(
    private state: ParserState,
    private exprParser: ExpressionParser,
    private parseStatement: () => Statement
  ) {}

  parseSubStatement(visibility?: string): VbSubStatement {
    const startToken = this.state.current;
    this.state.expect('Sub' as any);
    const name = this.exprParser.parseIdentifier();
    const params = this.parseParameters();
    this.state.skipNewlines();
    const body = this.parseBlock('End');
    this.state.expect('End' as any);
    this.state.expect('Sub' as any);

    return {
      type: 'VbSubStatement',
      name,
      params,
      body,
      visibility: (visibility ?? 'public') as 'public' | 'private',
      loc: createLocation(startToken, this.state.previous),
    };
  }

  parseFunctionStatement(visibility?: string): VbFunctionStatement {
    const startToken = this.state.current;
    this.state.expect('Function' as any);
    const name = this.exprParser.parseIdentifier();
    const params = this.parseParameters();
    this.state.skipNewlines();
    const body = this.parseBlock('End');
    this.state.expect('End' as any);
    this.state.expect('Function' as any);

    return {
      type: 'VbFunctionStatement',
      name,
      params,
      body,
      visibility: (visibility ?? 'public') as 'public' | 'private',
      loc: createLocation(startToken, this.state.previous),
    };
  }

  parseClassStatement(): VbClassStatement {
    const classToken = this.state.advance();
    const name = this.exprParser.parseIdentifier();
    this.state.skipNewlines();

    const body = this.parseClassBody();

    this.state.expect('End' as any);
    this.state.expect('Class' as any);

    return {
      type: 'VbClassStatement',
      name,
      body,
      loc: createLocation(classToken, this.state.previous),
    };
  }

  parsePropertyStatement(visibility?: string): Statement {
    this.state.expect('Property' as any);
    const propVisibility = visibility ?? 'public';

    if (this.state.check('Get' as any)) {
      return this.parsePropertyGet(propVisibility);
    }
    if (this.state.check('Let' as any)) {
      return this.parsePropertyLet(propVisibility);
    }
    if (this.state.check('Set' as any)) {
      return this.parsePropertySet(propVisibility);
    }

    throw new Error('Expected Get, Let, or Set after Property');
  }

  private parsePropertyGet(visibility: string): VbPropertyGetStatement {
    this.state.expect('Get' as any);
    const name = this.exprParser.parseIdentifier();
    const params = this.parseParameters();
    this.state.skipNewlines();
    const body = this.parseBlock('End');
    this.state.expect('End' as any);
    this.state.expect('Property' as any);

    return {
      type: 'VbPropertyGetStatement',
      name,
      params,
      body,
      visibility: visibility as 'public' | 'private',
      loc: createLocation(this.state.previous, this.state.previous),
    };
  }

  private parsePropertyLet(visibility: string): VbPropertyLetStatement {
    this.state.expect('Let' as any);
    const name = this.exprParser.parseIdentifier();
    const params = this.parseParameters();
    this.state.skipNewlines();
    const body = this.parseBlock('End');
    this.state.expect('End' as any);
    this.state.expect('Property' as any);

    return {
      type: 'VbPropertyLetStatement',
      name,
      params,
      body,
      visibility: visibility as 'public' | 'private',
      loc: createLocation(this.state.previous, this.state.previous),
    };
  }

  private parsePropertySet(visibility: string): VbPropertySetStatement {
    this.state.expect('Set' as any);
    const name = this.exprParser.parseIdentifier();
    const params = this.parseParameters();
    this.state.skipNewlines();
    const body = this.parseBlock('End');
    this.state.expect('End' as any);
    this.state.expect('Property' as any);

    return {
      type: 'VbPropertySetStatement',
      name,
      params,
      body,
      visibility: visibility as 'public' | 'private',
      loc: createLocation(this.state.previous, this.state.previous),
    };
  }

  parseParameters(): VbParameter[] {
    const params: VbParameter[] = [];

    if (!this.state.check('LParen' as any)) {
      return params;
    }

    this.state.advance();

    if (!this.state.check('RParen' as any)) {
      do {
        this.state.skipOptionalNewlines();
        const param = this.parseParameter();
        params.push(param);
        this.state.skipOptionalNewlines();
      } while (this.state.match('Comma' as any));
    }

    this.state.expect('RParen' as any);
    return params;
  }

  private parseParameter(): VbParameter {
    let byRef = true;
    let isOptional = false;
    let isParamArray = false;

    if (this.state.check('Optional' as any)) {
      this.state.advance();
      isOptional = true;
    }

    if (this.state.check('ParamArray' as any)) {
      this.state.advance();
      isParamArray = true;
    }

    if (this.state.check('ByRef' as any)) {
      this.state.advance();
      byRef = true;
    } else if (this.state.check('ByVal' as any)) {
      this.state.advance();
      byRef = false;
    }

    const name = this.exprParser.parseIdentifier();
    let isArray = false;
    let defaultValue: any;

    if (this.state.check('LParen' as any)) {
      this.state.advance();
      isArray = true;
      this.state.expect('RParen' as any);
    }

    if (this.state.match('Eq' as any)) {
      defaultValue = this.exprParser.parseExpression();
    }

    return {
      type: 'VbParameter',
      name,
      byRef,
      isArray,
      defaultValue,
      isOptional,
      isParamArray,
      loc: createLocation({ loc: name.loc! } as Token, this.state.previous),
    };
  }

  parseBlock(endKeyword?: string): BlockStatement {
    const body: Statement[] = [];

    while (!this.state.isEOF) {
      this.state.skipNewlines();

      if (endKeyword && this.state.check(endKeyword as any)) {
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

  private parseClassBody(): VbClassElement[] {
    const body: VbClassElement[] = [];

    while (!this.state.isEOF) {
      this.state.skipNewlines();

      if (this.state.check('End' as any)) {
        break;
      }

      if (this.state.isEOF) break;

      const stmt = this.parseStatement();
      if (isVbClassElement(stmt)) {
        body.push(stmt);
      }
    }

    return body;
  }
}
