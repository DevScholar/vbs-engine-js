import type { Program, Statement, VbLabelStatement } from '../ast/index.ts';
import type { VbValue } from '../runtime/index.ts';
import { VbContext, VbEmpty } from '../runtime/index.ts';
import { StatementExecutor, GotoSignal, ControlFlowSignal } from './statement-executor.ts';
import { ExpressionEvaluator } from './expression-evaluator.ts';
import { registerBuiltins } from '../builtins/index.ts';
import { parse } from '../parser/index.ts';

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

  run(program: Program): VbValue {
    this.startTime = Date.now();
    let result: VbValue = VbEmpty;
    const statements = program.body;
    const labels = this.collectLabels(statements);
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
          throw new Error(`Label not found: ${error.labelName}`);
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
    try {
      const ast = parse(code);
      const evaluator = new ExpressionEvaluator(this.context);
      const result = evaluator.evaluateProgram(ast);
      return result;
    } catch (e) {
      console.error('Eval error:', e);
      return { type: 'Empty', value: undefined };
    }
  }

  getContext(): VbContext {
    return this.context;
  }

  executeStatements(statements: Statement[]): VbValue {
    const labels = this.collectLabels(statements);
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
          throw new Error(`Label not found: ${error.labelName}`);
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
