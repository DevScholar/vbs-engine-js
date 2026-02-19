import { Lexer } from '../lexer/index.ts';
import { Parser } from '../parser/index.ts';
import { Interpreter } from '../interpreter/index.ts';
import type { VbValue } from '../runtime/index.ts';

function vbToJsAuto(value: VbValue): unknown {
  switch (value.type) {
    case 'Empty':
      return undefined;
    case 'Null':
      return null;
    case 'Boolean':
    case 'Long':
    case 'Double':
    case 'Integer':
    case 'String':
    case 'Date':
      return value.value;
    case 'Array': {
      const arr = value.value as { toArray(): VbValue[] };
      return arr.toArray().map(vbToJsAuto);
    }
    case 'Object':
      return value.value;
    default:
      return value.value;
  }
}

export class VbsEngine {
  private interpreter: Interpreter;
  private maxExecutionTime: number = -1;

  constructor() {
    this.interpreter = new Interpreter();
    this.interpreter.getContext().evaluate = (code: string) => this.interpreter.evaluate(code);
  }

  setMaxExecutionTime(ms: number): void {
    this.maxExecutionTime = ms;
    this.interpreter.setMaxExecutionTime(ms);
    this.interpreter.getContext().checkTimeout = () => this.interpreter.checkTimeout();
  }

  run(source: string): VbValue {
    const lexer = new Lexer(source);
    const tokens = lexer.tokenize();
    const parser = new Parser(tokens);
    const program = parser.parse();
    return this.interpreter.run(program);
  }

  getVariable(name: string): VbValue {
    return this.interpreter.getVariable(name);
  }

  getVariableAsJs(name: string): unknown {
    return vbToJsAuto(this.interpreter.getVariable(name));
  }

  setVariable(name: string, value: VbValue): void {
    this.interpreter.setVariable(name, value);
  }

  registerFunction(name: string, func: (...args: VbValue[]) => VbValue): void {
    this.interpreter.registerFunction(name, func);
  }

  getContext() {
    return this.interpreter.getContext();
  }
}

export function runVbscript(source: string): VbValue {
  const engine = new VbsEngine();
  return engine.run(source);
}
