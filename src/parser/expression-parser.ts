import type {
  Expression,
  Identifier,
  Literal,
  VbEmptyLiteral,
  VbNewExpression,
  MemberExpression,
  CallExpression,
  BinaryExpression,
  LogicalExpression,
  AssignmentExpression,
} from '../ast/index.ts';
import type { Token } from '../lexer/index.ts';
import { TokenType } from '../lexer/token.ts';
import { ParserState } from './parser-state.ts';
import {
  createLocation,
  createLocationFromNode,
  createLocationFromNodeAndToken,
} from './location.ts';

export class ExpressionParser {
  private state: ParserState;

  constructor(state: ParserState) {
    this.state = state;
  }

  parseExpression(): Expression {
    return this.parseAssignment();
  }

  parseStatementExpression(): Expression {
    return this.parseStatementAssignment();
  }

  parseCallExpression(): Expression {
    return this.parseCall();
  }

  parseMemberExpression(): Expression {
    let expr = this.parseIdentifierOnly();

    while (true) {
      if (this.state.check('Dot' as TokenType)) {
        this.state.advance();
        const property = this.parsePropertyName();
        expr = {
          type: 'MemberExpression',
          object: expr,
          property,
          computed: false,
          optional: false,
          loc: createLocationFromNode(expr, property),
        } as MemberExpression;
      } else if (this.state.check('Bang' as TokenType)) {
        this.state.advance();
        const property = this.parsePropertyName();
        expr = {
          type: 'MemberExpression',
          object: expr,
          property,
          computed: true,
          optional: false,
          loc: createLocationFromNode(expr, property),
        } as MemberExpression;
      } else if (this.state.check('LBracket' as TokenType)) {
        this.state.advance();
        const index = this.parseExpression();
        const rbracket = this.state.expect('RBracket' as TokenType);
        expr = {
          type: 'MemberExpression',
          object: expr,
          property: index,
          computed: true,
          optional: false,
          loc: createLocationFromNodeAndToken(expr, rbracket),
        } as MemberExpression;
      } else {
        break;
      }
    }

    return expr;
  }

  private parseIdentifierOnly(): Expression {
    if (this.state.check('Dot' as TokenType)) {
      return this.parseWithMemberExpression();
    }

    const token = this.state.expect('Identifier' as TokenType);
    return {
      type: 'Identifier',
      name: token.value,
      loc: token.loc,
    };
  }

  private parseStatementAssignment(): Expression {
    if (this.state.check('Identifier' as TokenType) || this.state.check('Dot' as TokenType)) {
      const savedState = this.state.save();
      const left = this.parseCall();
      
      if (left.type === 'CallExpression') {
        return this.continueStringConcat(left);
      }
      
      if (this.state.check('Eq' as TokenType)) {
        if (left.type === 'Identifier' || left.type === 'MemberExpression') {
          const op = this.state.advance();
          const right = this.parseStatementAssignment();
          return this.createAssignmentExpression(left, '=', right, op);
        }
      }
      
      if ((left.type === 'Identifier' || left.type === 'MemberExpression') &&
          this.isStatementCallArgumentStart()) {
        const args = this.parseStatementCallArguments();
        const callExpr: CallExpression = {
          type: 'CallExpression',
          callee: left,
          arguments: args,
          optional: false,
          loc: createLocationFromNodeAndToken(left, this.state.previous),
        } as CallExpression;
        return this.continueStringConcat(callExpr);
      }
      
      this.state.restore(savedState);
    }

    return this.parseStringConcat();
  }

  private continueStringConcat(left: Expression): Expression {
    while (this.state.check('Ampersand' as TokenType)) {
      this.state.advance();
      const right = this.parseLogicalOr();
      left = {
        type: 'BinaryExpression',
        operator: '&',
        left,
        right,
        loc: createLocationFromNode(left, right),
      } as BinaryExpression;
    }
    return left;
  }

  private isStatementCallArgumentStart(): boolean {
    return this.state.checkAny(
      'StringLiteral' as TokenType,
      'NumberLiteral' as TokenType,
      'DateLiteral' as TokenType,
      'BooleanLiteral' as TokenType,
      'NothingLiteral' as TokenType,
      'NullLiteral' as TokenType,
      'EmptyLiteral' as TokenType,
      'Identifier' as TokenType,
      'LParen' as TokenType,
      'New' as TokenType,
    );
  }

  private parseStatementCallArguments(): Expression[] {
    const args: Expression[] = [];

    while (!this.state.checkAny('Newline' as TokenType, 'Colon' as TokenType, 'EOF' as TokenType)) {
      args.push(this.parseExpression());
      if (!this.state.match('Comma' as TokenType)) {
        break;
      }
    }

    return args;
  }

  private parseAssignment(): Expression {
    const expr = this.parseStringConcat();

    if (this.state.checkAny('Eq' as TokenType)) {
      const op = this.state.advance();
      const right = this.parseAssignment();
      return this.createAssignmentExpression(expr, '=', right, op);
    }

    return expr;
  }

  private createAssignmentExpression(
    left: Expression,
    operator: string,
    right: Expression,
    token: Token
  ): AssignmentExpression {
    return {
      type: 'AssignmentExpression',
      operator: operator as AssignmentExpression['operator'],
      left: left as AssignmentExpression['left'],
      right,
      loc: createLocationFromNodeAndToken(left, token),
    };
  }

  private parseStringConcat(): Expression {
    let left = this.parseLogicalOr();

    while (this.state.check('Ampersand' as TokenType)) {
      this.state.advance();
      const right = this.parseLogicalOr();
      left = {
        type: 'BinaryExpression',
        operator: '&',
        left,
        right,
        loc: createLocationFromNode(left, right),
      } as BinaryExpression;
    }

    return left;
  }

  private parseLogicalOr(): Expression {
    let left = this.parseLogicalAnd();

    while (this.state.check('Or' as TokenType)) {
      this.state.advance();
      const right = this.parseLogicalAnd();
      left = {
        type: 'LogicalExpression',
        operator: '||',
        left,
        right,
        loc: createLocationFromNode(left, right),
      };
    }

    return left;
  }

  private parseLogicalAnd(): Expression {
    let left = this.parseLogicalNot();

    while (this.state.check('And' as TokenType)) {
      this.state.advance();
      const right = this.parseLogicalNot();
      left = {
        type: 'LogicalExpression',
        operator: '&&',
        left,
        right,
        loc: createLocation(
          { loc: left.loc! } as Token,
          { loc: right.loc! } as Token
        ),
      };
    }

    return left;
  }

  private parseLogicalNot(): Expression {
    let left = this.parseComparison();

    while (this.state.checkAny('Xor' as TokenType, 'Eqv' as TokenType, 'Imp' as TokenType)) {
      const op = this.state.advance();
      const right = this.parseComparison();
      const operator = op.value.toLowerCase() === 'xor' ? 'xor' : 
                      op.value.toLowerCase() === 'eqv' ? 'eqv' : 'imp';
      left = {
        type: 'LogicalExpression',
        operator: operator as LogicalExpression['operator'],
        left,
        right,
        loc: createLocation(
          { loc: left.loc! } as Token,
          { loc: right.loc! } as Token
        ),
      };
    }

    return left;
  }

  private parseComparison(): Expression {
    let left = this.parseIs();

    while (
      this.state.checkAny(
        'Eq' as TokenType,
        'Lt' as TokenType,
        'Gt' as TokenType,
        'Le' as TokenType,
        'Ge' as TokenType,
        'Ne' as TokenType
      )
    ) {
      const op = this.state.advance();
      const right = this.parseIs();
      const operator = this.getComparisonOperator(op);
      left = {
        type: 'BinaryExpression',
        operator,
        left,
        right,
        loc: createLocation(
          { loc: left.loc! } as Token,
          { loc: right.loc! } as Token
        ),
      };
    }

    return left;
  }

  private getComparisonOperator(token: Token): BinaryExpression['operator'] {
    switch (token.type) {
      case 'Eq':
        return '==';
      case 'Lt':
        return '<';
      case 'Gt':
        return '>';
      case 'Le':
        return '<=';
      case 'Ge':
        return '>=';
      case 'Ne':
        return '!=';
      default:
        return '==';
    }
  }

  private parseIs(): Expression {
    const left = this.parseConcatenation();

    if (this.state.check('Is' as TokenType)) {
      this.state.advance();
      const right = this.parseConcatenation();
      return {
        type: 'BinaryExpression',
        operator: 'Is',
        left,
        right,
        loc: createLocation(
          { loc: left.loc! } as Token,
          { loc: right.loc! } as Token
        ),
      } as BinaryExpression;
    }

    return left;
  }

  private parseConcatenation(): Expression {
    return this.parseAdditive();
  }

  private parseAdditive(): Expression {
    let left = this.parseMultiplicative();

    while (this.state.checkAny('Plus' as TokenType, 'Minus' as TokenType)) {
      const op = this.state.advance();
      const right = this.parseMultiplicative();
      left = {
        type: 'BinaryExpression',
        operator: op.type === 'Plus' ? '+' : '-',
        left,
        right,
        loc: createLocation(
          { loc: left.loc! } as Token,
          { loc: right.loc! } as Token
        ),
      };
    }

    return left;
  }

  private parseMultiplicative(): Expression {
    let left = this.parseIntegerDivision();

    while (
      this.state.checkAny('Asterisk' as TokenType, 'Slash' as TokenType)
    ) {
      const op = this.state.advance();
      const right = this.parseIntegerDivision();
      left = {
        type: 'BinaryExpression',
        operator: op.type === 'Asterisk' ? '*' : '/',
        left,
        right,
        loc: createLocation(
          { loc: left.loc! } as Token,
          { loc: right.loc! } as Token
        ),
      };
    }

    return left;
  }

  private parseIntegerDivision(): Expression {
    let left = this.parseMod();

    while (this.state.check('Backslash' as TokenType)) {
      this.state.advance();
      const right = this.parseMod();
      left = {
        type: 'BinaryExpression',
        operator: '\\' as BinaryExpression['operator'],
        left,
        right,
        loc: createLocation(
          { loc: left.loc! } as Token,
          { loc: right.loc! } as Token
        ),
      };
    }

    return left;
  }

  private parseMod(): Expression {
    let left = this.parsePower();

    while (this.state.check('Mod' as TokenType)) {
      this.state.advance();
      const right = this.parsePower();
      left = {
        type: 'BinaryExpression',
        operator: '%' as BinaryExpression['operator'],
        left,
        right,
        loc: createLocation(
          { loc: left.loc! } as Token,
          { loc: right.loc! } as Token
        ),
      };
    }

    return left;
  }

  private parsePower(): Expression {
    let left = this.parseUnary();

    while (this.state.check('Caret' as TokenType)) {
      this.state.advance();
      const right = this.parseUnary();
      left = {
        type: 'BinaryExpression',
        operator: '**' as BinaryExpression['operator'],
        left,
        right,
        loc: createLocation(
          { loc: left.loc! } as Token,
          { loc: right.loc! } as Token
        ),
      };
    }

    return left;
  }

  private parseUnary(): Expression {
    if (this.state.check('Not' as TokenType)) {
      const op = this.state.advance();
      const argument = this.parseUnary();
      return {
        type: 'UnaryExpression',
        operator: '!',
        prefix: true,
        argument,
        loc: createLocation(op, { loc: argument.loc! } as Token),
      };
    }

    if (this.state.check('Minus' as TokenType)) {
      const op = this.state.advance();
      const argument = this.parseUnary();
      return {
        type: 'UnaryExpression',
        operator: '-',
        prefix: true,
        argument,
        loc: createLocation(op, { loc: argument.loc! } as Token),
      };
    }

    if (this.state.check('Plus' as TokenType)) {
      const op = this.state.advance();
      const argument = this.parseUnary();
      return {
        type: 'UnaryExpression',
        operator: '+',
        prefix: true,
        argument,
        loc: createLocation(op, { loc: argument.loc! } as Token),
      };
    }

    return this.parsePostfix();
  }

  private parsePostfix(): Expression {
    return this.parseCall();
  }

  private parseCall(): Expression {
    let expr = this.parsePrimary();

    while (true) {
      if (this.state.check('Dot' as TokenType)) {
        this.state.advance();
        const property = this.parsePropertyName();
        expr = {
          type: 'MemberExpression',
          object: expr,
          property,
          computed: false,
          optional: false,
          loc: createLocation(
            { loc: expr.loc! } as Token,
            { loc: property.loc! } as Token
          ),
        } as MemberExpression;
      } else if (this.state.check('Bang' as TokenType)) {
        this.state.advance();
        const property = this.parsePropertyName();
        expr = {
          type: 'MemberExpression',
          object: expr,
          property,
          computed: true,
          optional: false,
          loc: createLocation(
            { loc: expr.loc! } as Token,
            { loc: property.loc! } as Token
          ),
        } as MemberExpression;
      } else if (this.state.check('LParen' as TokenType)) {
        this.state.advance();
        const args = this.parseArguments();
        const rparen = this.state.expect('RParen' as TokenType);
        expr = {
          type: 'CallExpression',
          callee: expr,
          arguments: args,
          optional: false,
          loc: createLocation(
            { loc: expr.loc! } as Token,
            rparen
          ),
        } as CallExpression;
      } else if (this.state.check('LBracket' as TokenType)) {
        this.state.advance();
        const index = this.parseExpression();
        const rbracket = this.state.expect('RBracket' as TokenType);
        expr = {
          type: 'MemberExpression',
          object: expr,
          property: index,
          computed: true,
          optional: false,
          loc: createLocation(
            { loc: expr.loc! } as Token,
            rbracket
          ),
        } as MemberExpression;
      } else {
        break;
      }
    }

    return expr;
  }

  private parseArguments(): Expression[] {
    const args: Expression[] = [];

    if (!this.state.check('RParen' as TokenType)) {
      while (true) {
        this.state.skipOptionalNewlines();
        
        if (this.state.check('RParen' as TokenType)) {
          break;
        }
        
        // Handle empty arguments (consecutive commas)
        if (this.state.check('Comma' as TokenType)) {
          // Empty argument - push Empty literal
          args.push({
            type: 'VbEmptyLiteral',
            value: undefined,
            raw: '',
            loc: this.state.current.loc,
          });
        } else {
          args.push(this.parseExpression());
        }
        
        this.state.skipOptionalNewlines();
        
        if (this.state.check('Comma' as TokenType)) {
          this.state.advance();
        } else {
          break;
        }
      }
    }

    return args;
  }

  parsePrimary(): Expression {
    if (this.state.check('LParen' as TokenType)) {
      return this.parseParenExpression();
    }

    if (this.state.check('StringLiteral' as TokenType)) {
      return this.parseStringLiteral();
    }

    if (this.state.check('NumberLiteral' as TokenType)) {
      return this.parseNumberLiteral();
    }

    if (this.state.check('DateLiteral' as TokenType)) {
      return this.parseDateLiteral();
    }

    if (this.state.check('BooleanLiteral' as TokenType)) {
      return this.parseBooleanLiteral();
    }

    if (this.state.check('NothingLiteral' as TokenType)) {
      return this.parseNothingLiteral();
    }

    if (this.state.check('NullLiteral' as TokenType)) {
      return this.parseNullLiteral();
    }

    if (this.state.check('EmptyLiteral' as TokenType)) {
      return this.parseEmptyLiteral();
    }

    if (this.state.check('New' as TokenType)) {
      return this.parseNewExpression();
    }

    if (this.state.check('Dot' as TokenType)) {
      return this.parseWithMemberExpression();
    }

    if (this.state.check('Identifier' as TokenType)) {
      return this.parseIdentifierOrCall();
    }

    throw new Error(`Unexpected token: ${this.state.current.type}`);
  }

  private parseWithMemberExpression(): MemberExpression {
    const dotToken = this.state.advance();
    const property = this.parsePropertyName();
    return {
      type: 'MemberExpression',
      object: { type: 'VbWithObject', loc: dotToken.loc } as Expression,
      property,
      computed: false,
      optional: false,
      loc: createLocation(dotToken, { loc: property.loc! } as Token),
    } as MemberExpression;
  }

  private parseParenExpression(): Expression {
    this.state.advance();
    this.state.skipOptionalNewlines();
    const expr = this.parseExpression();
    this.state.skipOptionalNewlines();
    this.state.expect('RParen' as TokenType);
    return expr;
  }

  private parseStringLiteral(): Literal {
    const token = this.state.advance();
    return {
      type: 'Literal',
      value: token.value,
      raw: token.raw ?? undefined,
      loc: token.loc,
    };
  }

  private parseNumberLiteral(): Literal {
    const token = this.state.advance();
    const value = token.value.includes('.') || token.value.includes('e') || token.value.includes('E')
      ? parseFloat(token.value)
      : parseInt(token.value, 10);
    return {
      type: 'Literal',
      value,
      raw: token.raw ?? token.value,
      loc: token.loc,
    };
  }

  private parseDateLiteral(): Literal {
    const token = this.state.advance();
    return {
      type: 'Literal',
      value: new Date(token.value),
      raw: token.raw ?? undefined,
      loc: token.loc,
    };
  }

  private parseBooleanLiteral(): Literal {
    const token = this.state.advance();
    return {
      type: 'Literal',
      value: token.value.toLowerCase() === 'true',
      raw: token.raw ?? undefined,
      loc: token.loc,
    };
  }

  private parseNothingLiteral(): Literal {
    const token = this.state.advance();
    return {
      type: 'Literal',
      value: Symbol.for('Nothing'),
      raw: token.raw ?? undefined,
      loc: token.loc,
    };
  }

  private parseNullLiteral(): Literal {
    const token = this.state.advance();
    return {
      type: 'Literal',
      value: null,
      raw: token.raw ?? undefined,
      loc: token.loc,
    };
  }

  private parseEmptyLiteral(): VbEmptyLiteral {
    const token = this.state.advance();
    return {
      type: 'VbEmptyLiteral',
      value: undefined,
      raw: token.raw ?? undefined,
      loc: token.loc,
    };
  }

  private parseNewExpression(): VbNewExpression {
    const newToken = this.state.advance();
    const callee = this.parseIdentifier();

    let args: Expression[] = [];
    if (this.state.check('LParen' as TokenType)) {
      this.state.advance();
      args = this.parseArguments();
      this.state.expect('RParen' as TokenType);
    }

    return {
      type: 'VbNewExpression',
      callee,
      arguments: args,
      loc: createLocation(newToken, this.state.previous),
    };
  }

  private parseIdentifierOrCall(): Expression {
    const id = this.parseIdentifier();

    if (this.state.check('LParen' as TokenType)) {
      this.state.advance();
      const args = this.parseArguments();
      const rparen = this.state.expect('RParen' as TokenType);
      return {
        type: 'CallExpression',
        callee: id,
        arguments: args,
        optional: false,
        loc: createLocation({ loc: id.loc! } as Token, rparen),
      } as CallExpression;
    }

    return id;
  }

  parseIdentifier(): Identifier {
    const token = this.state.expect('Identifier' as TokenType);
    return {
      type: 'Identifier',
      name: token.value,
      loc: token.loc,
    };
  }

  private parsePropertyName(): Identifier {
    const token = this.state.current;
    if (token.type === 'Identifier' as TokenType || 
        (token.type !== 'EOF' as TokenType && 
         token.type !== 'Newline' as TokenType && 
         token.type !== 'LParen' as TokenType &&
         token.type !== 'RParen' as TokenType &&
         token.type !== 'Comma' as TokenType &&
         token.type !== 'Colon' as TokenType)) {
      this.state.advance();
      return {
        type: 'Identifier',
        name: token.value,
        loc: token.loc,
      };
    }
    throw new Error(`Expected property name, got ${token.type}`);
  }
}
