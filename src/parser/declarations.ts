import type {
  Expression,
  VbDimStatement,
  VbVariableDeclarator,
  VbReDimStatement,
  VbEraseStatement,
  VbConstStatement,
  VbConstDeclarator,
  VbTypeAnnotation,
} from '../ast/index.ts';
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

    let typeAnnotation: VbTypeAnnotation | undefined;
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

  private parseTypeAnnotation(): VbTypeAnnotation {
    this.state.expect(TokenType.As);
    const typeToken = this.state.current;
    let typeName = '';
    let isArray = false;

    if (this.state.checkAny(
      TokenType.Integer,
      TokenType.Long,
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
    )) {
      typeName = this.state.advance().value;
    }

    if (this.state.check(TokenType.LParen)) {
      this.state.advance();
      isArray = true;
      this.state.expect(TokenType.RParen);
    }

    return {
      type: 'VbTypeAnnotation',
      typeName,
      isArray,
      loc: createLocation(typeToken, this.state.previous),
    };
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

    let typeAnnotation: VbTypeAnnotation | undefined;
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
