import type { Statement, BlockStatement, ExpressionStatement, IfStatement, VbDimStatement, VbReDimStatement, VbConstStatement, VbForToStatement, VbForEachStatement, VbDoLoopStatement, VbSelectCaseStatement, VbWithStatement, VbExitStatement, VbOptionExplicitStatement, VbSubStatement, VbFunctionStatement, VbClassStatement, VbPropertyGetStatement, VbPropertyLetStatement, VbPropertySetStatement, VbOnErrorHandlerStatement, VbResumeStatement, VbCallStatement, Expression } from '../ast/index.ts';
import type { VbValue } from '../runtime/index.ts';
import { VbContext, VbEmpty, VbNull, VbNothing, createVbValue, toBoolean, toNumber, VbError, VbErrorCodes, createVbError, VbArray, createVbArray, VbObjectInstance, VbClass, VbProperty } from '../runtime/index.ts';
import { ExpressionEvaluator } from './expression-evaluator.ts';

export class ControlFlowSignal extends Error {
  constructor(public type: 'exit' | 'return' | 'continue' | 'break', public value?: VbValue) {
    super(type);
  }
}

export class StatementExecutor {
  private exprEvaluator: ExpressionEvaluator;

  constructor(private context: VbContext) {
    this.exprEvaluator = new ExpressionEvaluator(context);
  }

  execute(node: Statement): VbValue {
    try {
      switch (node.type) {
        case 'ExpressionStatement':
          return this.executeExpressionStatement(node);
        case 'BlockStatement':
          return this.executeBlockStatement(node);
        case 'IfStatement':
          return this.executeIfStatement(node);
        case 'VbDimStatement':
          return this.executeDimStatement(node);
        case 'VbReDimStatement':
          return this.executeReDimStatement(node);
        case 'VbConstStatement':
          return this.executeConstStatement(node);
        case 'VbForToStatement':
          return this.executeForToStatement(node);
        case 'VbForEachStatement':
          return this.executeForEachStatement(node);
        case 'VbDoLoopStatement':
          return this.executeDoLoopStatement(node);
        case 'VbSelectCaseStatement':
          return this.executeSelectCaseStatement(node);
        case 'VbWithStatement':
          return this.executeWithStatement(node);
        case 'VbExitStatement':
          return this.executeExitStatement(node);
        case 'VbOptionExplicitStatement':
          return this.executeOptionExplicitStatement(node);
        case 'VbSubStatement':
          return this.executeSubStatement(node);
        case 'VbFunctionStatement':
          return this.executeFunctionStatement(node);
        case 'VbClassStatement':
          return this.executeClassStatement(node);
        case 'VbPropertyGetStatement':
        case 'VbPropertyLetStatement':
        case 'VbPropertySetStatement':
          return this.executePropertyStatement(node);
        case 'VbOnErrorHandlerStatement':
          return this.executeOnErrorHandlerStatement(node);
        case 'VbCallStatement':
          return this.executeCallStatement(node);
        case 'ReturnStatement':
          throw new ControlFlowSignal('return');
        default:
          throw new Error(`Unknown statement type: ${(node as Statement).type}`);
      }
    } catch (error) {
      if (error instanceof ControlFlowSignal) {
        throw error;
      }
      if (error instanceof VbError) {
        this.context.setError(error);
        if (this.context.onErrorResumeNext) {
          return VbEmpty;
        }
        throw error;
      }
      const vbError = VbError.fromError(error as Error);
      this.context.setError(vbError);
      if (this.context.onErrorResumeNext) {
        return VbEmpty;
      }
      throw vbError;
    }
  }

  private executeExpressionStatement(node: ExpressionStatement): VbValue {
    return this.exprEvaluator.evaluate(node.expression);
  }

  private executeCallStatement(node: VbCallStatement): VbValue {
    return this.exprEvaluator.evaluateCall(node.callee, node.arguments);
  }

  private executeBlockStatement(node: BlockStatement): VbValue {
    let result: VbValue = VbEmpty;
    for (const stmt of node.body) {
      result = this.execute(stmt);
    }
    return result;
  }

  private executeIfStatement(node: IfStatement): VbValue {
    const test = this.exprEvaluator.evaluate(node.test);
    
    if (toBoolean(test)) {
      return this.execute(node.consequent);
    } else if (node.alternate) {
      return this.execute(node.alternate);
    }
    
    return VbEmpty;
  }

  private executeDimStatement(node: VbDimStatement): VbValue {
    for (const decl of node.declarations) {
      let value = VbEmpty;
      
      if (decl.isArray && decl.arrayBounds) {
        const bounds = decl.arrayBounds.map(b => {
          const boundValue = this.exprEvaluator.evaluate(b);
          return toNumber(boundValue);
        });
        const arr = createVbArray(bounds);
        value = { type: 'Array', value: arr };
      } else if (decl.init) {
        value = this.exprEvaluator.evaluate(decl.init);
      }
      
      this.context.declareVariable(decl.id.name, value);
    }
    
    return VbEmpty;
  }

  private executeReDimStatement(node: VbReDimStatement): VbValue {
    for (const decl of node.declarations) {
      if (!decl.isArray || !decl.arrayBounds) continue;
      
      const bounds = decl.arrayBounds.map(b => {
        const boundValue = this.exprEvaluator.evaluate(b);
        return toNumber(boundValue);
      });
      
      const existingVar = this.context.currentScope.get(decl.id.name);
      
      if (existingVar && existingVar.value.type === 'Array') {
        const arr = existingVar.value.value as VbArray;
        arr.redim(bounds, node.preserve);
      } else {
        const arr = createVbArray(bounds);
        this.context.setVariable(decl.id.name, { type: 'Array', value: arr });
      }
    }
    
    return VbEmpty;
  }

  private executeConstStatement(node: VbConstStatement): VbValue {
    for (const decl of node.declarations) {
      const value = this.exprEvaluator.evaluate(decl.init);
      this.context.currentScope.declare(decl.id.name, value, { isConst: true });
    }
    
    return VbEmpty;
  }

  private executeForToStatement(node: VbForToStatement): VbValue {
    const init = this.exprEvaluator.evaluate(node.init);
    const to = this.exprEvaluator.evaluate(node.to);
    const step = node.step ? this.exprEvaluator.evaluate(node.step) : createVbValue(1);
    
    const startValue = toNumber(init);
    const endValue = toNumber(to);
    const stepValue = toNumber(step);
    
    this.context.declareVariable(node.left.name, createVbValue(startValue));
    
    const shouldContinue = stepValue > 0 
      ? () => toNumber(this.context.getVariable(node.left.name)) <= endValue
      : () => toNumber(this.context.getVariable(node.left.name)) >= endValue;
    
    while (shouldContinue()) {
      if (this.context.checkTimeout) this.context.checkTimeout();
      
      try {
        this.execute(node.body);
      } catch (signal) {
        if (signal instanceof ControlFlowSignal) {
          if (signal.type === 'exit' && this.context.getExitFlag() === 'for') {
            this.context.clearExitFlag();
            break;
          }
          if (signal.type === 'return') {
            throw signal;
          }
        }
        throw signal;
      }
      
      const currentValue = toNumber(this.context.getVariable(node.left.name));
      this.context.setVariable(node.left.name, createVbValue(currentValue + stepValue));
    }
    
    return VbEmpty;
  }

  private executeForEachStatement(node: VbForEachStatement): VbValue {
    const collection = this.exprEvaluator.evaluate(node.right);
    
    if (collection.type !== 'Array' && collection.type !== 'Object') {
      throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch: expected array or collection', 'Vbscript');
    }
    
    let items: VbValue[];
    if (collection.type === 'Array') {
      const arr = collection.value as VbArray;
      items = arr.toArray();
    } else {
      const obj = collection.value as VbObjectInstance;
      const count = obj.getProperty('Count');
      const countNum = toNumber(count);
      items = [];
      for (let i = 0; i < countNum; i++) {
        items.push(obj.getProperty(String(i)));
      }
    }
    
    this.context.declareVariable(node.left.name, VbEmpty);
    
    for (const item of items) {
      this.context.setVariable(node.left.name, item);
      
      try {
        this.execute(node.body);
      } catch (signal) {
        if (signal instanceof ControlFlowSignal) {
          if (signal.type === 'exit' && this.context.getExitFlag() === 'for') {
            this.context.clearExitFlag();
            break;
          }
          if (signal.type === 'return') {
            throw signal;
          }
        }
        throw signal;
      }
    }
    
    return VbEmpty;
  }

  private executeDoLoopStatement(node: VbDoLoopStatement): VbValue {
    const checkCondition = (): boolean => {
      if (!node.test) return true;
      const testValue = this.exprEvaluator.evaluate(node.test!);
      return toBoolean(testValue);
    };
    
    const isWhile = node.testPosition === 'while-do' || node.testPosition === 'do-while';
    const isPreTest = node.testPosition === 'while-do' || node.testPosition === 'until-do';
    
    while (true) {
      if (this.context.checkTimeout) this.context.checkTimeout();
      
      if (isPreTest) {
        const cond = isWhile ? checkCondition() : !checkCondition();
        if (!cond) break;
      }
      
      try {
        this.execute(node.body);
      } catch (signal) {
        if (signal instanceof ControlFlowSignal) {
          if (signal.type === 'exit' && this.context.getExitFlag() === 'do') {
            this.context.clearExitFlag();
            break;
          }
          if (signal.type === 'return') {
            throw signal;
          }
        }
        throw signal;
      }
      
      if (!isPreTest) {
        const cond = isWhile ? checkCondition() : !checkCondition();
        if (!cond) break;
      }
    }
    
    return VbEmpty;
  }

  private executeSelectCaseStatement(node: VbSelectCaseStatement): VbValue {
    const discriminant = this.exprEvaluator.evaluate(node.discriminant);
    
    for (const caseClause of node.cases) {
      if (caseClause.isElse) {
        for (const stmt of caseClause.consequent) {
          this.execute(stmt);
        }
        return VbEmpty;
      }
      
      const matches = this.matchCase(caseClause.test, discriminant);
      
      if (matches) {
        for (const stmt of caseClause.consequent) {
          this.execute(stmt);
        }
        return VbEmpty;
      }
    }
    
    return VbEmpty;
  }

  private matchCase(test: Expression | Expression[] | null, discriminant: VbValue): boolean {
    if (!test) return false;
    
    const tests = Array.isArray(test) ? test : [test];
    
    for (const t of tests) {
      if (t.type === 'BinaryExpression' && 'left' in t && (t.left as Expression).type === 'Identifier' && ((t.left as Expression) as { name: string }).name === '__select_expr__') {
        const right = this.exprEvaluator.evaluate(t.right as Expression);
        const operator = t.operator;
        
        switch (operator) {
          case '==':
            if (toBoolean(this.equals(discriminant, right))) return true;
            break;
          case '<':
            if (toNumber(discriminant) < toNumber(right)) return true;
            break;
          case '>':
            if (toNumber(discriminant) > toNumber(right)) return true;
            break;
          case '<=':
            if (toNumber(discriminant) <= toNumber(right)) return true;
            break;
          case '>=':
            if (toNumber(discriminant) >= toNumber(right)) return true;
            break;
          case '!=':
            if (!toBoolean(this.equals(discriminant, right))) return true;
            break;
        }
      } else {
        const testValue = this.exprEvaluator.evaluate(t);
        if (toBoolean(this.equals(discriminant, testValue))) {
          return true;
        }
      }
    }
    
    return false;
  }

  private equals(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Empty' && right.type === 'Empty') {
      return { type: 'Boolean', value: true };
    }
    if (left.type === 'Null' || right.type === 'Null') {
      return { type: 'Boolean', value: false };
    }
    if (left.type === 'String' || right.type === 'String') {
      const leftStr = left.type === 'String' ? left.value as string : String(toNumber(left));
      const rightStr = right.type === 'String' ? right.value as string : String(toNumber(right));
      return { type: 'Boolean', value: leftStr.toLowerCase() === rightStr.toLowerCase() };
    }
    return { type: 'Boolean', value: toNumber(left) === toNumber(right) };
  }

  private executeWithStatement(node: VbWithStatement): VbValue {
    const object = this.exprEvaluator.evaluate(node.object);
    this.context.pushWith(object);
    
    try {
      this.execute(node.body);
    } finally {
      this.context.popWith();
    }
    
    return VbEmpty;
  }

  private executeExitStatement(node: VbExitStatement): VbValue {
    this.context.setExitFlag(node.target.toLowerCase() as 'sub' | 'function' | 'property' | 'do' | 'for' | 'select');
    throw new ControlFlowSignal('exit');
  }

  private executeOptionExplicitStatement(node: VbOptionExplicitStatement): VbValue {
    this.context.optionExplicit = true;
    return VbEmpty;
  }

  private executeSubStatement(node: VbSubStatement): VbValue {
    const subName = node.name.name;
    const self = this;
    
    const subFunc = function(args: VbValue[]): VbValue {
      self.context.pushScope();
      self.context.pushCall(subName);
      
      try {
        for (let i = 0; i < node.params.length; i++) {
          const param = node.params[i];
          const argValue = args[i] ?? (param.defaultValue ? self.exprEvaluator.evaluate(param.defaultValue) : VbEmpty);
          self.context.declareVariable(param.name.name, argValue);
        }
        
        self.executeBlockStatement(node.body);
        
        for (let i = 0; i < node.params.length && i < args.length; i++) {
          const param = node.params[i];
          if (param.byRef) {
            args[i] = self.context.getVariable(param.name.name);
          }
        }
      } catch (signal) {
        if (signal instanceof ControlFlowSignal && signal.type === 'return') {
          for (let i = 0; i < node.params.length && i < args.length; i++) {
            const param = node.params[i];
            if (param.byRef) {
              args[i] = self.context.getVariable(param.name.name);
            }
          }
        } else {
          throw signal;
        }
      } finally {
        self.context.popCall();
        self.context.popScope();
      }
      
      return VbEmpty;
    };
    
    const params = node.params.map(p => ({
      name: p.name.name,
      byRef: p.byRef,
      isArray: p.isArray
    }));
    
    this.context.functionRegistry.register(subName, subFunc, { isSub: true, params, isUserDefined: true });
    return VbEmpty;
  }

  private executeFunctionStatement(node: VbFunctionStatement): VbValue {
    const funcName = node.name.name;
    const self = this;
    
    const func = function(args: VbValue[]): VbValue {
      self.context.pushScope();
      self.context.pushCall(funcName);
      
      self.context.declareVariable(funcName, VbEmpty);
      
      let result: VbValue = VbEmpty;
      
      try {
        for (let i = 0; i < node.params.length; i++) {
          const param = node.params[i];
          const argValue = args[i] ?? (param.defaultValue ? self.exprEvaluator.evaluate(param.defaultValue) : VbEmpty);
          self.context.declareVariable(param.name.name, argValue);
        }
        
        self.executeBlockStatement(node.body);
        
        for (let i = 0; i < node.params.length && i < args.length; i++) {
          const param = node.params[i];
          if (param.byRef) {
            args[i] = self.context.getVariable(param.name.name);
          }
        }
        
        result = self.context.getVariable(funcName);
      } catch (signal) {
        if (signal instanceof ControlFlowSignal && signal.type === 'return') {
          for (let i = 0; i < node.params.length && i < args.length; i++) {
            const param = node.params[i];
            if (param.byRef) {
              args[i] = self.context.getVariable(param.name.name);
            }
          }
          result = self.context.getVariable(funcName);
        } else {
          throw signal;
        }
      } finally {
        self.context.popCall();
        self.context.popScope();
      }
      
      return result;
    };
    
    const params = node.params.map(p => ({
      name: p.name.name,
      byRef: p.byRef,
      isArray: p.isArray
    }));
    
    this.context.functionRegistry.register(funcName, func, { params, isUserDefined: true });
    return VbEmpty;
  }

  private executeClassStatement(node: VbClassStatement): VbValue {
    const className = node.name.name;
    const cls = new VbClass(className);
    const self = this;
    
    for (const member of node.body) {
      if (member.type === 'VbDimStatement') {
        for (const decl of member.declarations) {
          cls.properties.set(decl.id.name.toLowerCase(), {
            name: decl.id.name,
          });
        }
      } else if (member.type === 'VbSubStatement') {
        const memberNode = member;
        cls.methods.set(member.name.name.toLowerCase(), {
          name: member.name.name,
          func: function(this: VbObjectInstance, ...args: VbValue[]): VbValue {
            const prevInstance = self.context.currentInstance;
            self.context.currentInstance = this;
            self.context.pushScope();
            try {
              for (let i = 0; i < memberNode.params.length; i++) {
                const param = memberNode.params[i];
                const argValue = args[i] ?? (param.defaultValue ? self.exprEvaluator.evaluate(param.defaultValue) : VbEmpty);
                self.context.declareVariable(param.name.name, argValue);
              }
              self.executeBlockStatement(memberNode.body);
            } finally {
              self.context.popScope();
              self.context.currentInstance = prevInstance;
            }
            return VbEmpty;
          },
          isSub: true,
        });
      } else if (member.type === 'VbFunctionStatement') {
        const memberNode = member;
        cls.methods.set(member.name.name.toLowerCase(), {
          name: member.name.name,
          func: function(this: VbObjectInstance, ...args: VbValue[]): VbValue {
            const prevInstance = self.context.currentInstance;
            self.context.currentInstance = this;
            self.context.pushScope();
            self.context.declareVariable(memberNode.name.name, VbEmpty);
            for (let i = 0; i < memberNode.params.length; i++) {
              const param = memberNode.params[i];
              const argValue = args[i] ?? (param.defaultValue ? self.exprEvaluator.evaluate(param.defaultValue) : VbEmpty);
              self.context.declareVariable(param.name.name, argValue);
            }
            let result: VbValue = VbEmpty;
            try {
              self.executeBlockStatement(memberNode.body);
              result = self.context.getVariable(memberNode.name.name);
            } catch (signal) {
              if (signal instanceof ControlFlowSignal && signal.type === 'return') {
                result = self.context.getVariable(memberNode.name.name);
              } else {
                throw signal;
              }
            } finally {
              self.context.popScope();
              self.context.currentInstance = prevInstance;
            }
            return result;
          },
          isSub: false,
        });
      } else if (member.type === 'VbPropertyGetStatement') {
        const memberNode = member;
        const propName = member.name.name.toLowerCase();
        const existing = cls.properties.get(propName) || { name: member.name.name };
        existing.get = function(this: VbObjectInstance): VbValue {
          const prevInstance = self.context.currentInstance;
          const prevInPropertyGet = self.context.inPropertyGet;
          const prevPropertyGetName = self.context.propertyGetName;
          self.context.currentInstance = this;
          self.context.inPropertyGet = true;
          self.context.propertyGetName = memberNode.name.name.toLowerCase();
          self.context.pushScope();
          self.context.declareVariable(memberNode.name.name, VbEmpty);
          try {
            self.executeBlockStatement(memberNode.body);
            return self.context.currentScope.get(memberNode.name.name)?.value ?? VbEmpty;
          } finally {
            self.context.popScope();
            self.context.currentInstance = prevInstance;
            self.context.inPropertyGet = prevInPropertyGet;
            self.context.propertyGetName = prevPropertyGetName;
          }
        };
        cls.properties.set(propName, existing);
      } else if (member.type === 'VbPropertyLetStatement') {
        const memberNode = member;
        const propName = member.name.name.toLowerCase();
        const existing = cls.properties.get(propName) || { name: member.name.name };
        existing.let = function(this: VbObjectInstance, value: VbValue): void {
          const prevInstance = self.context.currentInstance;
          self.context.currentInstance = this;
          self.context.pushScope();
          self.context.declareVariable(memberNode.name.name, VbEmpty);
          const param = memberNode.params[0];
          if (param) {
            self.context.declareVariable(param.name.name, value);
          }
          try {
            self.executeBlockStatement(memberNode.body);
          } finally {
            self.context.popScope();
            self.context.currentInstance = prevInstance;
          }
        };
        cls.properties.set(propName, existing);
      } else if (member.type === 'VbPropertySetStatement') {
        const memberNode = member;
        const propName = member.name.name.toLowerCase();
        const existing = cls.properties.get(propName) || { name: member.name.name };
        existing.set = function(this: VbObjectInstance, value: VbValue): void {
          const prevInstance = self.context.currentInstance;
          self.context.currentInstance = this;
          self.context.pushScope();
          self.context.declareVariable(memberNode.name.name, VbEmpty);
          const param = memberNode.params[0];
          if (param) {
            self.context.declareVariable(param.name.name, value);
          }
          try {
            self.executeBlockStatement(memberNode.body);
          } finally {
            self.context.popScope();
            self.context.currentInstance = prevInstance;
          }
        };
        cls.properties.set(propName, existing);
      }
    }
    
    this.context.classRegistry.register(cls);
    return VbEmpty;
  }

  private executePropertyStatement(node: Statement): VbValue {
    return VbEmpty;
  }

  private executeOnErrorHandlerStatement(node: VbOnErrorHandlerStatement): VbValue {
    if (node.action === 'resume_next') {
      this.context.onErrorResumeNext = true;
    } else if (node.action === 'goto_0') {
      this.context.onErrorResumeNext = false;
      this.context.clearError();
    }
    return VbEmpty;
  }
}
