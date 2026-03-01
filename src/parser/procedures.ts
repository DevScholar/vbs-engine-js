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
import { TokenType } from '../lexer/token.ts';
import { ParserState } from './parser-state.ts';
import { ExpressionParser } from './expression-parser.ts';
import { createLocation, createLocationFromNodeAndToken } from './location.ts';

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
    this.state.expect(TokenType.Sub);
    const name = this.exprParser.parseIdentifier();
    const params = this.parseParameters();
    this.state.skipNewlines();
    const body = this.parseBlock(TokenType.End);
    this.state.expect(TokenType.End);
    this.state.expect(TokenType.Sub);

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
    this.state.expect(TokenType.Function);
    const name = this.exprParser.parseIdentifier();
    const params = this.parseParameters();
    this.state.skipNewlines();
    const body = this.parseBlock(TokenType.End);
    this.state.expect(TokenType.End);
    this.state.expect(TokenType.Function);

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

    this.state.expect(TokenType.End);
    this.state.expect(TokenType.Class);

    return {
      type: 'VbClassStatement',
      name,
      body,
      loc: createLocation(classToken, this.state.previous),
    };
  }

  parsePropertyStatement(visibility?: string): Statement {
    this.state.expect(TokenType.Property);
    const propVisibility = visibility ?? 'public';

    if (this.state.check(TokenType.Get)) {
      return this.parsePropertyGet(propVisibility);
    }
    if (this.state.check(TokenType.Let)) {
      return this.parsePropertyLet(propVisibility);
    }
    if (this.state.check(TokenType.Set)) {
      return this.parsePropertySet(propVisibility);
    }

    throw new Error('Expected Get, Let, or Set after Property');
  }

  private parsePropertyGet(visibility: string): VbPropertyGetStatement {
    this.state.expect(TokenType.Get);
    const name = this.exprParser.parseIdentifier();
    const params = this.parseParameters();
    this.state.skipNewlines();
    const body = this.parseBlock(TokenType.End);
    this.state.expect(TokenType.End);
    this.state.expect(TokenType.Property);

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
    this.state.expect(TokenType.Let);
    const name = this.exprParser.parseIdentifier();
    const params = this.parseParameters();
    this.state.skipNewlines();
    const body = this.parseBlock(TokenType.End);
    this.state.expect(TokenType.End);
    this.state.expect(TokenType.Property);

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
    this.state.expect(TokenType.Set);
    const name = this.exprParser.parseIdentifier();
    const params = this.parseParameters();
    this.state.skipNewlines();
    const body = this.parseBlock(TokenType.End);
    this.state.expect(TokenType.End);
    this.state.expect(TokenType.Property);

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

    if (!this.state.check(TokenType.LParen)) {
      return params;
    }

    this.state.advance();

    if (!this.state.check(TokenType.RParen)) {
      do {
        this.state.skipOptionalNewlines();
        const param = this.parseParameter();
        params.push(param);
        this.state.skipOptionalNewlines();
      } while (this.state.match(TokenType.Comma));
    }

    this.state.expect(TokenType.RParen);
    return params;
  }

  private parseParameter(): VbParameter {
    let byRef = true;
    let isOptional = false;
    let isParamArray = false;

    if (this.state.check(TokenType.Optional)) {
      this.state.advance();
      isOptional = true;
    }

    if (this.state.check(TokenType.ParamArray)) {
      this.state.advance();
      isParamArray = true;
    }

    if (this.state.check(TokenType.ByRef)) {
      this.state.advance();
      byRef = true;
    } else if (this.state.check(TokenType.ByVal)) {
      this.state.advance();
      byRef = false;
    }

    const name = this.exprParser.parseIdentifier();
    let isArray = false;
    let defaultValue: unknown;

    if (this.state.check(TokenType.LParen)) {
      this.state.advance();
      isArray = true;
      this.state.expect(TokenType.RParen);
    }

    if (this.state.match(TokenType.Eq)) {
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
      loc: createLocationFromNodeAndToken(name, this.state.previous),
    };
  }

  parseBlock(endKeyword?: TokenType): BlockStatement {
    const body: Statement[] = [];

    while (!this.state.isEOF) {
      this.state.skipStatementSeparators();

      if (endKeyword && this.state.check(endKeyword)) {
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
      this.state.skipStatementSeparators();

      if (this.state.check(TokenType.End)) {
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
