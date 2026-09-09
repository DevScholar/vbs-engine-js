import type {
  Expression,
  Identifier,
  Literal,
  VbEmptyLiteral,
  NewExpression,
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

/**
 * Converts a call-expression chain (as produced by parseCall() for `name(args)`
 * or chained `name(args)(args)`) into the equivalent nested MemberExpression
 * chain, for use as an assignment target. Returns null if the chain isn't
 * convertible: any level with more than one argument (true multi-dimensional
 * indexing isn't supported by the single-index array-write path this feeds
 * into), or a base that isn't ultimately an Identifier/MemberExpression.
 */
function callChainToMemberChain(expr: Expression): MemberExpression | null {
  if (expr.type === 'Identifier' || expr.type === 'MemberExpression') {
    return expr as MemberExpression;
  }
  if (expr.type !== 'CallExpression') {
    return null;
  }
  const call = expr as CallExpression;
  if (call.arguments.length !== 1) {
    return null;
  }
  const object =
    call.callee.type === 'Identifier' || call.callee.type === 'MemberExpression'
      ? (call.callee as Expression)
      : callChainToMemberChain(call.callee as Expression);
  if (!object) {
    return null;
  }
  return {
    type: 'MemberExpression',
    object,
    property: call.arguments[0],
    computed: true,
    optional: false,
    loc: call.loc,
  } as MemberExpression;
}

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
    // The `Call` statement (this method's only caller) can target ANY
    // expression, not just a plain identifier - including a parenthesized
    // expression whose result is then immediately member-accessed/called,
    // e.g. `Call (New testclass).publicSub()`. Found via Wine's own
    // vbscript.dll conformance suite, dlls/vbscript/tests/lang.vbs. Every
    // other base (bare identifier) still goes through parseIdentifierOnly()
    // unchanged; the postfix loop below (Dot/Bang/LBracket chaining) already
    // works uniformly regardless of what `expr` started as.
    let expr: Expression;
    if (this.state.check('LParen' as TokenType)) {
      this.state.advance();
      expr = this.parseExpression();
      this.state.expect('RParen' as TokenType);
    } else {
      expr = this.parseIdentifierOnly();
    }

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
        // `name(args)` is ambiguous in VBScript between an array-index read/write
        // and a sub/function call - parseCall() always builds a (possibly nested,
        // for chained/jagged access like `arr(0)(4)`) CallExpression for it, since
        // arrays and calls share identical syntax. That's correct for a *read*
        // (`y = arr(1)`), but for a *write* (`arr(1) = y`) this used to return
        // immediately, before ever checking for a following `=`, so indexed-array
        // assignment could never be recognized at the statement level (bug found
        // and root-caused 2026-09-03 evaluating this engine for vp-engine-wasm's
        // VP-table VBScript replacement). Fix: if `=` follows a call-chain target
        // built entirely of single-argument calls, convert the whole chain into
        // the equivalent nested MemberExpression and treat it as an indexed-
        // assignment target instead of a bare call-statement. Multi-argument
        // calls anywhere in the chain (true multi-dimensional `arr(i, j) = x`)
        // are left unconverted - assignToMember only supports a single index per
        // level today, so converting those here would silently mis-handle them
        // rather than fix them; that's a separate, narrower gap.
        const memberChain = this.state.check('Eq' as TokenType)
          ? callChainToMemberChain(left as CallExpression)
          : null;
        if (memberChain) {
          const op = this.state.advance();
          const right = this.parseStatementAssignment();
          return this.createAssignmentExpression(memberChain, '=', right, op);
        }

        // A Comma immediately following a single call-chain (not `=`) means this is a
        // bare statement-style call whose first argument merely happens to be written
        // with its own parens - a classic, well-documented VBScript idiom:
        //   PlaySound("fx_x" & b), -1, Vol(...), Pan(...), 0, pitch, 1, 0, fade
        // parseCall() has no way to know ahead of time that `PlaySound(...)` isn't a
        // complete, self-contained call - arrays and calls share identical syntax, so
        // it correctly builds a CallExpression from just the parenthesized part. Real
        // VBScript then treats everything after the comma as MORE arguments to that
        // same statement call, not a syntax error and not a second statement. Found
        // via a real production table script (Chance, Playmatic 1971) that hit
        // "Unexpected token: Comma" here - confirmed via a minimal repro
        // (`Foo("x"), 1, 2`) before this fix. Only applies when `left` is a plain call
        // (callee is Identifier/MemberExpression, not itself a further call chain) -
        // `arr(0)(1), 2` (jagged-array read result comma-continued) isn't a real
        // statement-call pattern and shouldn't be coerced into one.
        if (
          this.state.check('Comma' as TokenType) &&
          (left as CallExpression).callee.type !== 'CallExpression'
        ) {
          const args = [...(left as CallExpression).arguments];
          while (this.state.match('Comma' as TokenType)) {
            args.push(this.parseExpression());
          }
          const callExpr: CallExpression = {
            type: 'CallExpression',
            callee: (left as CallExpression).callee,
            arguments: args,
            optional: false,
            loc: left.loc,
          } as CallExpression;
          return this.continueStringConcat(callExpr);
        }

        // Not an indexed-assignment target either - `left` (e.g. `CInt(1)`)
        // is a genuine call/read whose result may still be the LEFT operand
        // of any binary operator (`CInt(1) / Empty`, `Foo() + 5`,
        // `Bar() > threshold`, etc.), not just `&` string-concat.
        // continueStringConcat() only knows about `&`, so anything else left
        // dangling after `left` either got silently discarded (Plus/Minus,
        // which can also start a new top-level statement, so the leftover
        // ` + 5` was reparsed as an unrelated, meaningless second statement)
        // or threw outright ("Unexpected token: Asterisk"/"...Gt" etc. -
        // found via Wine's own vbscript.dll conformance suite,
        // dlls/vbscript/tests/lang.vbs: `CInt(1) / Empty` failed to parse at
        // all). Fix: abandon this fast-path attempt and restore+reparse via
        // the general expression grammar, exactly like the "left.type isn't
        // Identifier/MemberExpression/CallExpression at all" fallback at the
        // bottom of this function already does - it correctly re-derives
        // `left` fresh and continues through every real operator uniformly,
        // there's no need to hand-splice continuation logic for each one.
        this.state.restore(savedState);
        return this.parseStringConcat();
      }

      if (this.state.check('Eq' as TokenType)) {
        if (left.type === 'Identifier' || left.type === 'MemberExpression') {
          const op = this.state.advance();
          const right = this.parseStatementAssignment();
          return this.createAssignmentExpression(left, '=', right, op);
        }
      }

      if (
        (left.type === 'Identifier' || left.type === 'MemberExpression') &&
        // A bare Comma right after the callee (no parens at all) is real, valid
        // VBScript too: `PlayersReel.SetValue, PlayersPlayingGame` deliberately
        // omits the first positional argument, same idiom as `Foo(, x)` inside
        // parens (parseArguments() already handles that case via VbEmptyLiteral
        // below) - just without any parens wrapping the whole call here. Found
        // via the same production table script as the CallExpression/Comma fix
        // above (Chance, Playmatic 1971) - a second, distinct instance of the
        // "statement-call arg list starts unusually" gap, not a duplicate.
        (this.isStatementCallArgumentStart() || this.state.check('Comma' as TokenType))
      ) {
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
      // `Not` - a real, common no-parens call pattern (found via Wine's own
      // vbscript.dll conformance test suite, dlls/vbscript/tests/lang.vbs:
      // `ok not false, "msg"` failed with "Unexpected token: Comma" because
      // this list didn't recognize `Not` as a valid argument start at all,
      // so the call was never recognized as a call - `ok` got parsed as a
      // bare read, then `not false, "msg"` was left dangling as an invalid
      // top-level statement). `Not` is unambiguous here (always unary in
      // VBScript, never a binary operator), unlike Minus/Plus - deliberately
      // NOT added: `x = x + i` needs `+` parsed as binary addition once `x`
      // (the RHS's own leading identifier) is reached recursively here, but
      // adding Plus/Minus made this branch instead misparse it as `x(+i)` -
      // a real regression, caught immediately by the existing test suite
      // (`For-To loop` etc. silently computed 0 instead of 15) before ever
      // reaching the wine-conformance harness. Minus/Plus genuinely cannot
      // be disambiguated from the next token alone at this decision point.
      'Not' as TokenType
    );
  }

  private parseStatementCallArguments(): Expression[] {
    const args: Expression[] = [];

    while (!this.state.checkAny('Newline' as TokenType, 'Colon' as TokenType, 'EOF' as TokenType)) {
      // Omitted positional argument (leading or consecutive Comma, no expression
      // between) - same VbEmptyLiteral representation parseArguments() already
      // uses for the parenthesized-call equivalent (`Foo(, x)`).
      if (this.state.check('Comma' as TokenType)) {
        args.push({
          type: 'VbEmptyLiteral',
          value: undefined,
          raw: '',
          loc: this.state.current.loc,
        } as Expression);
      } else {
        args.push(this.parseExpression());
      }
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
        loc: createLocation({ loc: left.loc! } as Token, { loc: right.loc! } as Token),
      };
    }

    return left;
  }

  private parseLogicalNot(): Expression {
    let left = this.parseComparison();

    while (this.state.checkAny('Xor' as TokenType, 'Eqv' as TokenType, 'Imp' as TokenType)) {
      const op = this.state.advance();
      const right = this.parseComparison();
      const operator =
        op.value.toLowerCase() === 'xor' ? 'xor' : op.value.toLowerCase() === 'eqv' ? 'eqv' : 'imp';
      left = {
        type: 'LogicalExpression',
        operator: operator as LogicalExpression['operator'],
        left,
        right,
        loc: createLocation({ loc: left.loc! } as Token, { loc: right.loc! } as Token),
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
        loc: createLocation({ loc: left.loc! } as Token, { loc: right.loc! } as Token),
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
        loc: createLocation({ loc: left.loc! } as Token, { loc: right.loc! } as Token),
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
        loc: createLocation({ loc: left.loc! } as Token, { loc: right.loc! } as Token),
      };
    }

    return left;
  }

  private parseMultiplicative(): Expression {
    let left = this.parseIntegerDivision();

    while (this.state.checkAny('Asterisk' as TokenType, 'Slash' as TokenType)) {
      const op = this.state.advance();
      const right = this.parseIntegerDivision();
      left = {
        type: 'BinaryExpression',
        operator: op.type === 'Asterisk' ? '*' : '/',
        left,
        right,
        loc: createLocation({ loc: left.loc! } as Token, { loc: right.loc! } as Token),
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
        loc: createLocation({ loc: left.loc! } as Token, { loc: right.loc! } as Token),
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
        loc: createLocation({ loc: left.loc! } as Token, { loc: right.loc! } as Token),
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
        loc: createLocation({ loc: left.loc! } as Token, { loc: right.loc! } as Token),
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
          loc: createLocation({ loc: expr.loc! } as Token, { loc: property.loc! } as Token),
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
          loc: createLocation({ loc: expr.loc! } as Token, { loc: property.loc! } as Token),
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
          loc: createLocation({ loc: expr.loc! } as Token, rparen),
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
          loc: createLocation({ loc: expr.loc! } as Token, rbracket),
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

    // `Me` (the current class instance) was never actually wired up here -
    // the interpreter's evaluateMe()/ThisExpression handling already existed
    // but nothing in the parser ever produced a ThisExpression node, and
    // `Me` wasn't even a registered lexer keyword, so it silently parsed as
    // an ordinary Identifier and failed at runtime with "Variable is
    // undefined: 'Me'" the moment a method body actually read it. Found via
    // Wine's own vbscript.dll conformance suite, dlls/vbscript/tests/
    // lang.vbs (a chained-call test: `(New testclass).publicSub()` style
    // code returning `Me` from a method).
    if (this.state.check('Me' as TokenType)) {
      const token = this.state.advance();
      return { type: 'ThisExpression', loc: token.loc } as Expression;
    }

    if (this.state.check('Dot' as TokenType)) {
      return this.parseWithMemberExpression();
    }

    if (
      this.state.checkAny(
        'Identifier' as TokenType,
        // Type-annotation keywords (Dim x As String, etc. - see parseTypeAnnotation() in
        // declarations.ts) double as ordinary identifiers everywhere else, same as any other
        // built-in name shadowing in real VBScript. See parseFlexibleIdentifier()'s comment.
        'Integer' as TokenType,
        'Long' as TokenType,
        'LongLong' as TokenType,
        'Single' as TokenType,
        'Double' as TokenType,
        'Currency' as TokenType,
        'String' as TokenType,
        'Boolean' as TokenType,
        'Date' as TokenType,
        'Object' as TokenType,
        'Variant' as TokenType,
        'Byte' as TokenType,
        // `Property` too, when NOT immediately starting a real Property
        // Get/Let/Set block (that case is already routed to
        // parsePropertyStatement() before parsePrimary() is ever reached) -
        // e.g. `Dim Property` then later `Property = true`. Found via
        // Wine's own vbscript.dll conformance suite, dlls/vbscript/tests/
        // lang.vbs.
        'Property' as TokenType,
        // `Error`/`Explicit`/`Step` too (each declared via `Dim` then
        // assigned as a plain variable, e.g. `Dim step : step = "xx"`) -
        // same rule, same suite (Wine's lang.vbs `test_identifiers` sub
        // exhaustively exercises every VBScript keyword that's also
        // required to work as an ordinary identifier).
        'Error' as TokenType,
        'Explicit' as TokenType,
        'Step' as TokenType
      )
    ) {
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
    const value =
      token.value.includes('.') || token.value.includes('e') || token.value.includes('E')
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

  private parseNewExpression(): NewExpression {
    const newToken = this.state.advance();

    // Parse the first identifier, then consume any dot-chained members
    // so that `New Forms.Form` and `New NS.Sub.Class` are supported.
    let callee: import('../ast/types.ts').Identifier | import('../ast/types.ts').MemberExpression =
      this.parseIdentifier();

    while (this.state.check('Dot' as TokenType)) {
      this.state.advance();
      const prop = this.parseIdentifier();
      callee = {
        type: 'MemberExpression',
        object: callee,
        property: prop,
        computed: false,
        optional: false,
        loc: callee.loc,
      } as import('../ast/types.ts').MemberExpression;
    }

    let args: Expression[] = [];
    if (this.state.check('LParen' as TokenType)) {
      this.state.advance();
      args = this.parseArguments();
      this.state.expect('RParen' as TokenType);
    }

    return {
      type: 'NewExpression',
      callee,
      arguments: args,
      loc: createLocation(newToken, this.state.previous),
    };
  }

  private parseIdentifierOrCall(): Expression {
    const id = this.parseFlexibleIdentifier();

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

  // Type-annotation keywords (String/Integer/Long/Boolean/Date/Object/Variant/etc. - see
  // parseTypeAnnotation() in declarations.ts, this engine's `Dim x As String` extension) are
  // valid ordinary identifiers everywhere EXCEPT immediately after `As` - real VBScript has no
  // reserved type names at all, and shadowing built-in function names (`String`, `Date`, etc.)
  // as a variable/parameter name is completely normal. Found via a real production table script
  // (Chance, Playmatic 1971): `Function GetHSChar(String, Index)` then `Mid(String, Index, 1)`
  // inside its body - both the parameter declaration AND every later reference to it hit strict
  // identifier checks that only accepted the plain Identifier token type. Exposed publicly
  // (parsePropertyName stays private, used only for member-access names within this file) -
  // procedures.ts's parseParameter() and this file's own parseIdentifierOrCall() both need it.
  parseFlexibleIdentifier(): Identifier {
    return this.parsePropertyName();
  }

  private parsePropertyName(): Identifier {
    const token = this.state.current;
    if (
      token.type === ('Identifier' as TokenType) ||
      (token.type !== ('EOF' as TokenType) &&
        token.type !== ('Newline' as TokenType) &&
        token.type !== ('LParen' as TokenType) &&
        token.type !== ('RParen' as TokenType) &&
        token.type !== ('Comma' as TokenType) &&
        token.type !== ('Colon' as TokenType))
    ) {
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
