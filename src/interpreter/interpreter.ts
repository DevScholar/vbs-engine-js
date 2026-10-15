import type { Expression, Program, Statement, VbLabelStatement } from '../ast/index.ts';
import type { VbValue } from '../runtime/index.ts';
import { VbContext, VbEmpty, createVbError } from '../runtime/index.ts';
import { StatementExecutor, GotoSignal, ControlFlowSignal } from './statement-executor.ts';
import { ExpressionEvaluator } from './expression-evaluator.ts';
import { registerBuiltins } from '../builtins/index.ts';
import { parse } from '../parser/index.ts';

/**
 * Takes the whole-statement marks off an expression, which Eval needs since it is handed an expression and
 * not a statement list.
 */
function clearStatementCallMarks(expression: Expression): void {
  if (expression.type === 'CallExpression') {
    expression.statementCall = false;
    expression.bareName = false;
  }
}

interface LabelInfo {
  index: number;
  statement: VbLabelStatement;
}

export class Interpreter {
  private context: VbContext;
  private executor: StatementExecutor;
  private maxExecutionTime: number = -1;
  private startTime: number = 0;

  constructor() {
    this.context = new VbContext();
    this.executor = new StatementExecutor(this.context);
    registerBuiltins(this.context);
  }

  setMaxExecutionTime(ms: number): void {
    this.maxExecutionTime = ms;
  }

  checkTimeout(): void {
    if (this.maxExecutionTime > 0) {
      const elapsed = Date.now() - this.startTime;
      if (elapsed > this.maxExecutionTime) {
        throw new Error(`Script execution timed out after ${this.maxExecutionTime}ms`);
      }
    }
  }

  private collectLabels(statements: Statement[]): Map<string, LabelInfo> {
    const labels = new Map<string, LabelInfo>();
    for (let i = 0; i < statements.length; i++) {
      const stmt = statements[i]!;
      if (stmt.type === 'VbLabelStatement') {
        const labelStmt = stmt as VbLabelStatement;
        labels.set(labelStmt.label.name.toLowerCase(), { index: i, statement: labelStmt });
      }
    }
    return labels;
  }

  // Real VBScript hoists every Sub/Function (and Class) declaration in a scope
  // before running any statement in that scope, so a call can textually precede
  // its own declaration - e.g. many VPX table scripts open with a bare call to a
  // loader Sub defined further down the file. This interpreter otherwise
  // registers a Sub/Function into functionRegistry only when execution reaches
  // it sequentially, so without this pre-pass such a forward call throws
  // "Variable is undefined" (or, without Option Explicit, silently no-ops
  // instead of calling the Sub). executor.execute() on a Sub/Function statement
  // only registers it - it doesn't run the body - so calling it again here has
  // no side effect beyond the one it already has each time it's reached in the
  // main loop below.
  private hoistDeclarations(statements: Statement[]): void {
    for (const stmt of statements) {
      this.hoistDeclarationsIn(stmt);
    }
  }

  // Real VBScript hoists a Sub/Function declaration no matter how deeply it's
  // textually nested inside a conditionally-dead branch - declarations are a
  // compile-time construct, resolved before ANY code runs, completely
  // independent of whether the branch containing them would ever actually
  // execute. `If False Then / Sub Foo ... End Sub / End If` still makes Foo
  // callable, even though that If branch never runs. Found via Wine's own
  // vbscript.dll conformance suite, dlls/vbscript/tests/lang.vbs. So this
  // walks into every kind of nested statement container a Sub/Function
  // declaration could textually appear inside - real VBScript doesn't allow
  // nested Sub/Function declarations INSIDE another Sub/Function/Property
  // body, so there's no need to recurse into those.
  private hoistDeclarationsIn(stmt: Statement): void {
    switch (stmt.type) {
      case 'VbSubStatement':
      case 'VbFunctionStatement':
      case 'VbClassStatement':
        // Class declarations are hoisted too, same as Sub/Function - `New
        // EmptyClass` can textually precede `Class EmptyClass ... End
        // Class`. executeClassStatement() just (re-)registers the class
        // into classRegistry, so calling it again when execution naturally
        // reaches it is harmless, same as Sub/Function.
        this.executor.execute(stmt);
        return;
      case 'BlockStatement':
        this.hoistDeclarations(stmt.body);
        return;
      case 'IfStatement':
        this.hoistDeclarationsIn(stmt.consequent);
        if (stmt.alternate) this.hoistDeclarationsIn(stmt.alternate);
        return;
      case 'VbDoLoopStatement':
      case 'WhileStatement':
      case 'VbForToStatement':
      case 'ForOfStatement':
      case 'WithStatement':
        this.hoistDeclarationsIn(stmt.body);
        return;
      case 'VbSelectCaseStatement':
        for (const caseClause of stmt.cases) {
          this.hoistDeclarations(caseClause.consequent);
        }
        return;
      default:
        return;
    }
  }

  run(program: Program): VbValue {
    this.startTime = Date.now();
    let result: VbValue = VbEmpty;
    const statements = program.body;
    const labels = this.collectLabels(statements);
    this.hoistDeclarations(statements);
    let i = 0;
    const maxIterations = statements.length * 10000;
    let iterations = 0;

    while (i < statements.length) {
      if (iterations++ > maxIterations) {
        throw new Error('Possible infinite loop detected (too many goto jumps)');
      }

      this.checkTimeout();
      const stmt = statements[i]!;

      try {
        result = this.executor.execute(stmt);
        i++;
      } catch (error) {
        if (error instanceof GotoSignal) {
          const labelInfo = labels.get(error.labelName);
          if (labelInfo) {
            i = labelInfo.index + 1;
            continue;
          }
          throw new Error(`Label not found: ${error.labelName}`, { cause: error });
        }
        if (error instanceof ControlFlowSignal) {
          throw error;
        }
        throw error;
      }
    }

    return result;
  }

  getVariable(name: string): VbValue {
    return this.context.getVariable(name);
  }

  setVariable(name: string, value: VbValue): void {
    this.context.setVariable(name, value);
  }

  registerFunction(name: string, func: (...args: VbValue[]) => VbValue): void {
    this.context.functionRegistry.register(name, func);
  }

  evaluate(code: string): VbValue {
    // Real VBScript's Eval() raises a catchable error 1002 ("Syntax error")
    // for malformed source, not a silent no-op - this previously caught and
    // swallowed EVERY error (parse failures and real runtime errors alike)
    // into a console.error + Empty return, so `On Error Resume Next` around
    // an Eval() call could never see anything went wrong at all (Err.Number
    // stayed 0). Only the parse phase gets remapped to 1002 here; a
    // syntactically valid Eval'd expression that fails at runtime (e.g.
    // `Eval("1/0")`) should keep propagating its own real error (11), not
    // 1002 - found via Wine's own vbscript.dll conformance suite,
    // dlls/vbscript/tests/lang.vbs (`Eval("&O8")` / `Eval("&19")`).
    let ast;
    try {
      ast = parse(code);
    } catch {
      throw createVbError(1002, 'Syntax error');
    }
    const evaluator = new ExpressionEvaluator(this.context);
    // Eval reads an expression, and a bare name in an expression is the value it holds or the procedure it
    // names rather than the call a bare name statement is. A fresh cscript answers Eval("number") with 5 and
    // Eval("NoSuchFunction") with 0, where `number` and `NoSuchFunction` written as statements are both error
    // 13, and Eval("Len") is the 450 the builtin raises when it is called with no arguments.
    //
    // The marks a statement list puts on a whole-statement call are cleared for the one expression Eval is
    // handed, since those marks change how an array and a host object are read: `WScript.Arguments(0)` alone
    // is 438 and `arr(9)` alone is 13, while Eval of either is the argument and the element.
    const body = ast.body;
    if (body.length === 1) {
      const only = body[0]!;
      if (only.type === 'ExpressionStatement') {
        const expression = only.expression;
        const wasBareName = expression.type === 'CallExpression' && expression.bareName === true;
        clearStatementCallMarks(expression);
        if (wasBareName) {
          return evaluator.evaluate((expression as { callee: Expression }).callee);
        }
        return evaluator.evaluate(expression);
      }
    }
    return evaluator.evaluateProgram(ast);
  }

  getContext(): VbContext {
    return this.context;
  }

  executeStatements(statements: Statement[]): VbValue {
    const labels = this.collectLabels(statements);
    this.hoistDeclarations(statements);
    let result: VbValue = VbEmpty;
    let i = 0;
    const maxIterations = statements.length * 10000;
    let iterations = 0;

    while (i < statements.length) {
      if (iterations++ > maxIterations) {
        throw new Error('Possible infinite loop detected (too many goto jumps)');
      }

      this.checkTimeout();
      const stmt = statements[i]!;

      try {
        result = this.executor.execute(stmt);
        i++;
      } catch (error) {
        if (error instanceof GotoSignal) {
          const labelInfo = labels.get(error.labelName);
          if (labelInfo) {
            i = labelInfo.index + 1;
            continue;
          }
          throw new Error(`Label not found: ${error.labelName}`, { cause: error });
        }
        if (error instanceof ControlFlowSignal) {
          throw error;
        }
        throw error;
      }
    }

    return result;
  }

  executeInCurrentScope(code: string): VbValue {
    const ast = parse(code);
    return this.executeStatements(ast.body);
  }

  executeInGlobalScope(code: string): VbValue {
    const savedScope = this.context.currentScope;
    this.context.currentScope = this.context.globalScope;
    try {
      const ast = parse(code);
      return this.executeStatements(ast.body);
    } finally {
      this.context.currentScope = savedScope;
    }
  }
}

export function interpret(program: Program): VbValue {
  const interpreter = new Interpreter();
  return interpreter.run(program);
}
