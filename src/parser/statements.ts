import type {
  Statement,
  Expression,
  WithStatement,
  VbExitStatement,
  VbOptionExplicitStatement,
} from "../ast/index.ts";
import { ParserState } from "./parser-state.ts";
import { ExpressionParser } from "./expression-parser.ts";
import { createLocation } from "./location.ts";
import { TokenType } from "../lexer/token.ts";

export class StatementsParser {
  constructor(
    private state: ParserState,
    private exprParser: ExpressionParser,
    private parseStatement: () => Statement,
  ) {}

  parseWithStatement(): WithStatement {
    const withToken = this.state.advance();
    const object = this.exprParser.parseExpression();
    this.state.skipNewlines();
    const body = this.parseWithBody();

    this.state.expect(TokenType.End);
    this.state.expect(TokenType.With);

    return {
      type: "WithStatement",
      object,
      body,
      loc: createLocation(withToken, this.state.previous),
    };
  }

  parseExitStatement(): VbExitStatement {
    const exitToken = this.state.advance();

    const targetTypes: Record<string, VbExitStatement["target"]> = {
      Sub: "Sub",
      Function: "Function",
      Property: "Property",
      Do: "Do",
      For: "For",
      Select: "Select",
    };

    const targetToken = this.state.current;
    const target = targetTypes[targetToken.type as string];

    if (target) {
      this.state.advance();
    }

    return {
      type: "VbExitStatement",
      target: target ?? "Sub",
      loc: createLocation(exitToken, this.state.previous),
    };
  }

  parseOptionStatement(): VbOptionExplicitStatement {
    const optionToken = this.state.advance();
    this.state.expect(TokenType.Explicit);

    return {
      type: "VbOptionExplicitStatement",
      loc: createLocation(optionToken, this.state.previous),
    };
  }

  parseCallStatement(): Statement {
    const callToken = this.state.advance();
    const callee = this.exprParser.parseMemberExpression();

    let args: Expression[] = [];
    if (this.state.check(TokenType.LParen)) {
      this.state.advance();
      args = this.parseCallArguments();
      this.state.expect(TokenType.RParen);
    } else if (
      !this.state.checkAny(TokenType.Newline, TokenType.Colon, TokenType.EOF)
    ) {
      args = this.parseCallArgumentsNoParens();
    }

    const callExpr = {
      type: "CallExpression" as const,
      callee,
      arguments: args,
      optional: false,
      // The `Call` keyword is a statement call in every shape, an array included.
      statementCall: true,
      loc: createLocation(callToken, this.state.previous),
    };

    return {
      type: "ExpressionStatement",
      expression: callExpr,
      loc: createLocation(callToken, this.state.previous),
    };
  }

  parseSetStatement(): Statement {
    const setToken = this.state.advance();
    const left = this.exprParser.parseSetTarget();
    this.state.expect(TokenType.Eq);
    const right = this.exprParser.parseExpression();

    const assignment = {
      type: "AssignmentExpression",
      operator: "=",
      left: left as Expression,
      right,
      isSet: true,
      loc: createLocation(setToken, this.state.previous),
    } as Expression;

    return {
      type: "ExpressionStatement",
      expression: assignment,
      loc: createLocation(setToken, this.state.previous),
    };
  }

  parseOnStatement(): Statement {
    const onToken = this.state.advance();
    this.state.expect(TokenType.Error);

    if (this.state.match(TokenType.Resume)) {
      this.state.expect(TokenType.Next);
      return {
        type: "VbOnErrorHandlerStatement",
        action: "resume_next",
        loc: createLocation(onToken, this.state.previous),
      };
    }

    if (this.state.match(TokenType.Goto)) {
      if (this.state.check(TokenType.NumberLiteral)) {
        const numToken = this.state.advance();
        if (String(numToken.value) === "0") {
          return {
            type: "VbOnErrorHandlerStatement",
            action: "goto_0",
            loc: createLocation(onToken, this.state.previous),
          };
        }
      }
      const label = this.exprParser.parseIdentifier();
      return {
        type: "VbOnErrorHandlerStatement",
        action: "goto_label",
        label,
        loc: createLocation(onToken, this.state.previous),
      };
    }

    throw new Error("Expected Resume Next or Goto after On Error");
  }

  parseResumeStatement(): Statement {
    const resumeToken = this.state.advance();

    if (this.state.match(TokenType.Next)) {
      return {
        type: "VbResumeStatement",
        target: "next",
        loc: createLocation(resumeToken, this.state.previous),
      };
    }

    return {
      type: "VbResumeStatement",
      target: null,
      loc: createLocation(resumeToken, this.state.previous),
    };
  }

  parseExpressionStatement(): Statement {
    const expr = this.exprParser.parseStatementExpression();

    // A statement that is just a name is a call on that name with no arguments, which a fresh cscript shows
    // from every side: "NoSuchFunction" alone is error 13, a variable holding a number, a text, an Empty, a
    // Null or an array is 13 as well, `NoArgs` and `NoArgFunction` run the Sub or Function they name,
    // `holder.DoIt`, `holder.GetIt` and `holder.Prop` run the method or property, and `holder` where the class
    // has no default member is 438. A statement that is a literal is a compile error there ("Empty" alone and
    // `5` alone are both refused), and this leaves those as they are.
    if (expr.type === "Identifier" || expr.type === "MemberExpression") {
      return {
        type: "ExpressionStatement",
        expression: {
          type: "CallExpression",
          callee: expr,
          arguments: [],
          optional: false,
          // A statement and an expression read a bare name differently, and Eval is handed an expression:
          // Eval("number") is 5 and Eval("NoSuchFunction") is 0, where `number` and `NoSuchFunction` written
          // as statements are both error 13. The mark is what lets Eval read the callee as a value again.
          statementCall: true,
          bareName: true,
          loc: expr.loc,
        },
        loc: expr.loc,
      };
    }

    // A whole statement that is a call is a Sub call rather than a read, which tells an array apart: a fresh
    // cscript answers `arr(9)` written as a statement with error 13 while `x = arr(9)` indexes it, so the
    // mark is carried on the call the statement already had.
    if (expr.type === "CallExpression") {
      expr.statementCall = true;
      // Parentheses around the arguments of a statement call hand them over by value, which a fresh cscript
      // shows: `FillWithNine a` changes the caller's array and `FillWithNine(a)` does not. A call written
      // without them is marked where it is built, and the `Call` keyword goes through parseCallStatement.
      if (expr.byValueArguments === undefined) {
        expr.byValueArguments = true;
      }
    }

    return {
      type: "ExpressionStatement",
      expression: expr,
      loc: expr.loc,
    };
  }

  private parseWithBody(): Statement {
    const body: Statement[] = [];

    while (!this.state.isEOF) {
      this.state.skipStatementSeparators();

      if (
        this.state.check(TokenType.End) &&
        this.state.peek(1).type === TokenType.With
      ) {
        break;
      }

      if (this.state.isEOF) break;

      const stmt = this.parseStatement();
      body.push(stmt);
    }

    return {
      type: "BlockStatement",
      body,
    };
  }

  private parseCallArguments(): Expression[] {
    const args: Expression[] = [];

    if (!this.state.check(TokenType.RParen)) {
      while (true) {
        this.state.skipOptionalNewlines();

        if (this.state.check(TokenType.RParen)) {
          break;
        }

        // Handle empty arguments (consecutive commas)
        if (this.state.check(TokenType.Comma)) {
          args.push({
            type: "VbEmptyLiteral",
            value: undefined,
            raw: "",
            loc: this.state.current.loc,
          });
        } else {
          args.push(this.exprParser.parseExpression());
        }

        this.state.skipOptionalNewlines();

        if (this.state.check(TokenType.Comma)) {
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

    while (
      !this.state.checkAny(TokenType.Newline, TokenType.Colon, TokenType.EOF)
    ) {
      args.push(this.exprParser.parseExpression());
      if (!this.state.match(TokenType.Comma)) {
        break;
      }
    }

    return args;
  }
}
