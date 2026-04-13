import type {
  Expression,
  VbDimStatement,
  VbVariableDeclarator,
  VbReDimStatement,
  VbEraseStatement,
  VbConstStatement,
  VbConstDeclarator,
  TSTypeAnnotation,
  TSType,
  TSEnumDeclaration,
  TSEnumMember,
} from '../ast/index.ts';
import type { SourceLocation } from '../ast/base.ts';
import { TokenType } from '../lexer/token.ts';
import { ParserState } from './parser-state.ts';
import { ExpressionParser } from './expression-parser.ts';
import { createLocation, createLocationFromNodeAndToken } from './location.ts';

export class DeclarationParser {
  constructor(
    private state: ParserState,
    private exprParser: ExpressionParser
  ) {}

  parseDimStatement(): VbDimStatement {
    const dimToken = this.state.advance();
    const declarations = this.parseVariableDeclarations();

    return {
      type: 'VbDimStatement',
      declarations,
      visibility: 'dim',
      loc: createLocation(dimToken, this.state.previous),
    };
  }

  parsePublicDimStatement(): VbDimStatement {
    const visibilityToken = this.state.advance();
    this.state.expect(TokenType.Dim);
    const declarations = this.parseVariableDeclarations();

    return {
      type: 'VbDimStatement',
      declarations,
      visibility: 'public',
      loc: createLocation(visibilityToken, this.state.previous),
    };
  }

  parsePrivateDimStatement(): VbDimStatement {
    const visibilityToken = this.state.advance();
    this.state.expect(TokenType.Dim);
    const declarations = this.parseVariableDeclarations();

    return {
      type: 'VbDimStatement',
      declarations,
      visibility: 'private',
      loc: createLocation(visibilityToken, this.state.previous),
    };
  }

  parseReDimStatement(): VbReDimStatement {
    const redimToken = this.state.advance();
    const preserve = this.state.match(TokenType.Preserve) !== null;
    const declarations = this.parseVariableDeclarations();

    return {
      type: 'VbReDimStatement',
      declarations,
      preserve,
      loc: createLocation(redimToken, this.state.previous),
    };
  }

  parseEraseStatement(): VbEraseStatement {
    const eraseToken = this.state.advance();
    const arrayName = this.exprParser.parseIdentifier();

    return {
      type: 'VbEraseStatement',
      arrayName,
      loc: createLocation(eraseToken, this.state.previous),
    };
  }

  parseEnumStatement(visibility?: string): TSEnumDeclaration {
    const enumToken = this.state.advance(); // consume 'Enum'
    const id = this.exprParser.parseIdentifier();
    this.state.skipNewlines();

    const members: TSEnumMember[] = [];
    while (!this.state.isEOF) {
      this.state.skipStatementSeparators();
      if (this.state.check(TokenType.End)) break;
      if (this.state.isEOF) break;

      const memberName = this.exprParser.parseIdentifier();
      let initializer: Expression | null = null;
      if (this.state.match(TokenType.Eq)) {
        initializer = this.exprParser.parseExpression();
      }
      members.push({
        type: 'TSEnumMember',
        id: memberName,
        initializer,
        loc: createLocationFromNodeAndToken(memberName, this.state.previous),
      });
    }

    this.state.expect(TokenType.End);
    this.state.expect(TokenType.Enum);

    void visibility; // stored in parent context (Public/Private), not in TSEnumDeclaration

    return {
      type: 'TSEnumDeclaration',
      id,
      members,
      loc: createLocation(enumToken, this.state.previous),
    };
  }

  parseConstStatement(): VbConstStatement {
    const constToken = this.state.advance();
    const declarations = this.parseConstDeclarators();

    return {
      type: 'VbConstStatement',
      declarations,
      visibility: 'private',
      loc: createLocation(constToken, this.state.previous),
    };
  }

  parsePublicConstStatement(): VbConstStatement {
    const visibilityToken = this.state.advance();
    this.state.expect(TokenType.Const);
    const declarations = this.parseConstDeclarators();

    return {
      type: 'VbConstStatement',
      declarations,
      visibility: 'public',
      loc: createLocation(visibilityToken, this.state.previous),
    };
  }

  parsePrivateConstStatement(): VbConstStatement {
    const visibilityToken = this.state.advance();
    this.state.expect(TokenType.Const);
    const declarations = this.parseConstDeclarators();

    return {
      type: 'VbConstStatement',
      declarations,
      visibility: 'private',
      loc: createLocation(visibilityToken, this.state.previous),
    };
  }

  parseVariableDeclarations(): VbVariableDeclarator[] {
    const declarations: VbVariableDeclarator[] = [];

    do {
      const decl = this.parseVariableDeclarator();
      declarations.push(decl);
    } while (this.state.match(TokenType.Comma));

    return declarations;
  }

  private parseVariableDeclarator(): VbVariableDeclarator {
    const id = this.exprParser.parseIdentifier();
    let init: Expression | null = null;
    let isArray = false;
    let arrayBounds: Expression[] = [];

    if (this.state.check(TokenType.LParen)) {
      this.state.advance();
      isArray = true;
      arrayBounds = this.parseArrayBounds();
      this.state.expect(TokenType.RParen);
    }

    if (this.state.match(TokenType.Eq)) {
      init = this.exprParser.parseExpression();
    }

    let typeAnnotation: TSTypeAnnotation | undefined;
    if (this.state.check(TokenType.As)) {
      typeAnnotation = this.parseTypeAnnotation();
    }

    return {
      type: 'VbVariableDeclarator',
      id,
      init,
      isArray,
      arrayBounds: arrayBounds.length > 0 ? arrayBounds : undefined,
      typeAnnotation,
      loc: createLocationFromNodeAndToken(id, this.state.previous),
    };
  }

  private parseArrayBounds(): Expression[] {
    const bounds: Expression[] = [];

    do {
      this.state.skipOptionalNewlines();
      bounds.push(this.exprParser.parseExpression());
      this.state.skipOptionalNewlines();
    } while (this.state.match(TokenType.Comma));

    return bounds;
  }

  parseTypeAnnotation(): TSTypeAnnotation {
    this.state.expect(TokenType.As);
    const typeToken = this.state.current;
    let typeName = '';
    let isArray = false;

    if (
      this.state.checkAny(
        TokenType.Integer,
        TokenType.Long,
        TokenType.LongLong,
        TokenType.Single,
        TokenType.Double,
        TokenType.Currency,
        TokenType.String,
        TokenType.Boolean,
        TokenType.Date,
        TokenType.Object,
        TokenType.Variant,
        TokenType.Byte,
        TokenType.Identifier
      )
    ) {
      typeName = this.state.advance().value;
    }

    if (this.state.check(TokenType.LParen)) {
      this.state.advance();
      isArray = true;
      this.state.expect(TokenType.RParen);
    }

    const innerType = this.vbTypeNameToTSType(typeName, typeToken.loc);
    const tsType: TSType = isArray
      ? { type: 'TSArrayType', elementType: innerType, loc: typeToken.loc }
      : innerType;

    return {
      type: 'TSTypeAnnotation',
      typeAnnotation: tsType,
      loc: createLocation(typeToken, this.state.previous),
    };
  }

  private vbTypeNameToTSType(typeName: string, loc: SourceLocation | null | undefined): TSType {
    switch (typeName.toLowerCase()) {
      case 'string':  return { type: 'TSStringKeyword',  loc };
      case 'boolean': return { type: 'TSBooleanKeyword', loc };
      case 'object':  return { type: 'TSObjectKeyword',  loc };
      case 'variant': return { type: 'TSAnyKeyword',     loc };
      default:
        return {
          type: 'TSTypeReference',
          typeName: { type: 'Identifier', name: typeName || 'Variant', loc },
          loc,
        };
    }
  }

  private parseConstDeclarators(): VbConstDeclarator[] {
    const declarations: VbConstDeclarator[] = [];

    do {
      const decl = this.parseConstDeclarator();
      declarations.push(decl);
    } while (this.state.match(TokenType.Comma));

    return declarations;
  }

  private parseConstDeclarator(): VbConstDeclarator {
    const id = this.exprParser.parseIdentifier();
    this.state.expect(TokenType.Eq);
    const init = this.exprParser.parseExpression();

    let typeAnnotation: TSTypeAnnotation | undefined;
    if (this.state.check(TokenType.As)) {
      typeAnnotation = this.parseTypeAnnotation();
    }

    return {
      type: 'VbConstDeclarator',
      id,
      init,
      typeAnnotation,
      loc: createLocationFromNodeAndToken(id, this.state.previous),
    };
  }
}
