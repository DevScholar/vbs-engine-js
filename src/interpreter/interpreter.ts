import type { Program, Expression } from '../ast/index.ts';
import type { VbValue } from '../runtime/index.ts';
import { VbContext, VbEmpty } from '../runtime/index.ts';
import { StatementExecutor } from './statement-executor.ts';
import { ExpressionEvaluator } from './expression-evaluator.ts';
import { registerBuiltins } from '../builtins/index.ts';
import { Parser } from '../parser/index.ts';

export class Interpreter {
  private context: VbContext;
  private executor: StatementExecutor;
  private parser: Parser;
  private maxExecutionTime: number = -1;
  private startTime: number = 0;

  constructor() {
    this.context = new VbContext();
    this.executor = new StatementExecutor(this.context);
    this.parser = new Parser();
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

  run(program: Program): VbValue {
    this.startTime = Date.now();
    let result: VbValue = VbEmpty;

    for (const stmt of program.body) {
      this.checkTimeout();
      result = this.executor.execute(stmt);
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
      const ast = this.parser.parse(code);
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
}

export function interpret(program: Program): VbValue {
  const interpreter = new Interpreter();
  return interpreter.run(program);
}
