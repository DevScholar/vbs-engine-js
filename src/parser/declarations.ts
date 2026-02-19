import type {
  Expression,
  VbDimStatement,
  VbVariableDeclarator,
  VbReDimStatement,
  VbConstStatement,
  VbConstDeclarator,
  VbTypeAnnotation,
} from '../ast/index.ts';
import type { Token } from '../lexer/index.ts';
import { ParserState } from './parser-state.ts';
import { ExpressionParser } from './expression-parser.ts';
import { createLocation } from './location.ts';

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
    this.state.expect('Dim' as any);
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
    this.state.expect('Dim' as any);
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
    const preserve = this.state.match('Preserve' as any) !== null;
    const declarations = this.parseVariableDeclarations();

    return {
      type: 'VbReDimStatement',
      declarations,
      preserve,
      loc: createLocation(redimToken, this.state.previous),
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
    this.state.expect('Const' as any);
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
    this.state.expect('Const' as any);
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
    } while (this.state.match('Comma' as any));

    return declarations;
  }

  private parseVariableDeclarator(): VbVariableDeclarator {
    const id = this.exprParser.parseIdentifier();
    let init: Expression | null = null;
    let isArray = false;
    let arrayBounds: Expression[] = [];

    if (this.state.check('LParen' as any)) {
      this.state.advance();
      isArray = true;
      arrayBounds = this.parseArrayBounds();
      this.state.expect('RParen' as any);
    }

    if (this.state.match('Eq' as any)) {
      init = this.exprParser.parseExpression();
    }

    let typeAnnotation: VbTypeAnnotation | undefined;
    if (this.state.check('As' as any)) {
      typeAnnotation = this.parseTypeAnnotation();
    }

    return {
      type: 'VbVariableDeclarator',
      id,
      init,
      isArray,
      arrayBounds: arrayBounds.length > 0 ? arrayBounds : undefined,
      typeAnnotation,
      loc: createLocation({ loc: id.loc! } as Token, this.state.previous),
    };
  }

  private parseArrayBounds(): Expression[] {
    const bounds: Expression[] = [];

    do {
      this.state.skipOptionalNewlines();
      bounds.push(this.exprParser.parseExpression());
      this.state.skipOptionalNewlines();
    } while (this.state.match('Comma' as any));

    return bounds;
  }

  private parseTypeAnnotation(): VbTypeAnnotation {
    this.state.expect('As' as any);
    const typeToken = this.state.current;
    let typeName = '';
    let isArray = false;

    if (this.state.checkAny(
      'Integer' as any,
      'Long' as any,
      'Single' as any,
      'Double' as any,
      'Currency' as any,
      'String' as any,
      'Boolean' as any,
      'Date' as any,
      'Object' as any,
      'Variant' as any,
      'Byte' as any,
      'Identifier' as any
    )) {
      typeName = this.state.advance().value;
    }

    if (this.state.check('LParen' as any)) {
      this.state.advance();
      isArray = true;
      this.state.expect('RParen' as any);
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
    } while (this.state.match('Comma' as any));

    return declarations;
  }

  private parseConstDeclarator(): VbConstDeclarator {
    const id = this.exprParser.parseIdentifier();
    this.state.expect('Eq' as any);
    const init = this.exprParser.parseExpression();

    let typeAnnotation: VbTypeAnnotation | undefined;
    if (this.state.check('As' as any)) {
      typeAnnotation = this.parseTypeAnnotation();
    }

    return {
      type: 'VbConstDeclarator',
      id,
      init,
      typeAnnotation,
      loc: createLocation({ loc: id.loc! } as Token, this.state.previous),
    };
  }
}
