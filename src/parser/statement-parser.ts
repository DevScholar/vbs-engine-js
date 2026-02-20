import type { Statement, VbGotoStatement, VbLabelStatement } from '../ast/index.ts';
import type { TokenType } from '../lexer/index.ts';
import { ParserState } from './parser-state.ts';
import { ExpressionParser } from './expression-parser.ts';
import { DeclarationParser } from './declarations.ts';
import { ProcedureParser } from './procedures.ts';
import { ControlFlowParser } from './control-flow.ts';
import { StatementsParser } from './statements.ts';
import { createLocation } from './location.ts';

export class StatementParser {
  private state: ParserState;
  private exprParser: ExpressionParser;
  private declarationParser: DeclarationParser;
  private procedureParser: ProcedureParser;
  private controlFlowParser: ControlFlowParser;
  private statementsParser: StatementsParser;

  constructor(state: ParserState) {
    this.state = state;
    this.exprParser = new ExpressionParser(state);
    
    this.declarationParser = new DeclarationParser(state, this.exprParser);
    this.procedureParser = new ProcedureParser(state, this.exprParser, () => this.parseStatement());
    this.controlFlowParser = new ControlFlowParser(state, this.exprParser, () => this.parseStatement());
    this.statementsParser = new StatementsParser(state, this.exprParser, () => this.parseStatement());
  }

  parseStatement(): Statement {
    this.state.skipNewlines();

    const token = this.state.current;

    switch (token.type) {
      case 'Dim' as TokenType:
        return this.declarationParser.parseDimStatement();
      case 'ReDim' as TokenType:
        return this.declarationParser.parseReDimStatement();
      case 'Erase' as TokenType:
        return this.declarationParser.parseEraseStatement();
      case 'Const' as TokenType:
        return this.declarationParser.parseConstStatement();
      case 'Public' as TokenType:
        return this.parseVisibilityStatement();
      case 'Private' as TokenType:
        return this.parseVisibilityStatement();
      case 'Sub' as TokenType:
        return this.procedureParser.parseSubStatement();
      case 'Function' as TokenType:
        return this.procedureParser.parseFunctionStatement();
      case 'Class' as TokenType:
        return this.procedureParser.parseClassStatement();
      case 'Property' as TokenType:
        return this.procedureParser.parsePropertyStatement();
      case 'If' as TokenType:
        return this.controlFlowParser.parseIfStatement();
      case 'For' as TokenType:
        return this.controlFlowParser.parseForStatement();
      case 'Do' as TokenType:
        return this.controlFlowParser.parseDoStatement();
      case 'While' as TokenType:
        return this.controlFlowParser.parseWhileStatement();
      case 'Select' as TokenType:
        return this.controlFlowParser.parseSelectStatement();
      case 'With' as TokenType:
        return this.statementsParser.parseWithStatement();
      case 'Exit' as TokenType:
        return this.statementsParser.parseExitStatement();
      case 'Option' as TokenType:
        return this.statementsParser.parseOptionStatement();
      case 'Call' as TokenType:
        return this.statementsParser.parseCallStatement();
      case 'Set' as TokenType:
        return this.statementsParser.parseSetStatement();
      case 'On' as TokenType:
        return this.statementsParser.parseOnStatement();
      case 'Resume' as TokenType:
        return this.statementsParser.parseResumeStatement();
      case 'Goto' as TokenType:
        return this.parseGotoStatement();
      case 'Colon' as TokenType:
        this.state.advance();
        return this.parseStatement();
      default:
        if (this.state.checkIdentifier() && this.state.peek(1).type === 'Colon' as TokenType) {
          return this.parseLabelStatement();
        }
        return this.statementsParser.parseExpressionStatement();
    }
  }

  private parseGotoStatement(): VbGotoStatement {
    const gotoToken = this.state.advance();
    const label = this.exprParser.parseIdentifier();
    
    return {
      type: 'VbGotoStatement',
      label,
      loc: createLocation(gotoToken, this.state.previous),
    };
  }

  private parseLabelStatement(): VbLabelStatement {
    const labelToken = this.state.advance();
    this.state.expect('Colon' as TokenType);
    
    return {
      type: 'VbLabelStatement',
      label: {
        type: 'Identifier',
        name: labelToken.value,
        loc: labelToken.loc,
      },
      loc: createLocation(labelToken, this.state.previous),
    };
  }

  private parseVisibilityStatement(): Statement {
    const visibilityToken = this.state.advance();
    const visibility = visibilityToken.type === 'Public' ? 'public' : 'private';

    if (this.state.check('Dim' as TokenType)) {
      return this.declarationParser.parsePublicDimStatement();
    }

    if (this.state.check('Const' as TokenType)) {
      return visibility === 'public' 
        ? this.declarationParser.parsePublicConstStatement()
        : this.declarationParser.parsePrivateConstStatement();
    }

    if (this.state.check('Sub' as TokenType)) {
      return this.procedureParser.parseSubStatement(visibility);
    }

    if (this.state.check('Function' as TokenType)) {
      return this.procedureParser.parseFunctionStatement(visibility);
    }

    if (this.state.check('Property' as TokenType)) {
      return this.procedureParser.parsePropertyStatement(visibility);
    }

    if (this.state.checkIdentifier()) {
      const declarations = this.declarationParser.parseVariableDeclarations();
      return {
        type: 'VbDimStatement',
        declarations,
        visibility: visibility as 'public' | 'private',
        loc: { start: visibilityToken.loc.start, end: this.state.previous.loc.end },
      };
    }

    throw new Error(`Expected Dim, Const, Sub, Function, or Property after ${visibility}`);
  }
}
