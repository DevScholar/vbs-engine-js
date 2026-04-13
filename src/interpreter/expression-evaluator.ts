import type {
  Expression,
  Program,
  Identifier,
  Literal,
  MemberExpression,
  BinaryExpression,
  UnaryExpression,
  LogicalExpression,
  AssignmentExpression,
  ConditionalExpression,
  ThisExpression,
  VbNewExpression,
  VbWithObjectExpression,
} from '../ast/index.ts';
import type {
  VbValue,
  VbObjectValue,
  VbObjectValueData,
  VbArrayValue,
  VbBooleanValue,
  VbLongValue,
} from '../runtime/index.ts';
import { VbContext, VbObjectInstance } from '../runtime/index.ts';
import {
  createVbValue,
  VbEmpty,
  VbNull,
  toBoolean,
  toNumber,
  toString,
  createVbError,
  VbErrorCodes,
} from '../runtime/index.ts';

interface VbMethodObject {
  type: 'method';
  object: VbObjectValueData & {
    getMethod: (name: string) => { func: (...args: VbValue[]) => VbValue };
  };
  method: string;
}

interface VbJsFunctionObject {
  type: 'jsfunction';
  func: (...args: unknown[]) => unknown;
  thisArg: unknown;
}

function isVbMethodObject(obj: VbObjectValueData): obj is VbObjectValueData & VbMethodObject {
  return obj.type === 'method' && 'object' in obj && 'method' in obj;
}

function isVbJsFunctionObject(
  obj: VbObjectValueData
): obj is VbObjectValueData & VbJsFunctionObject {
  return obj.type === 'jsfunction' && 'func' in obj;
}

export class ExpressionEvaluator {
  constructor(private context: VbContext) {}

  evaluateProgram(program: Program): VbValue {
    let result: VbValue = VbEmpty;
    for (const stmt of program.body) {
      if (stmt.type === 'ExpressionStatement') {
        result = this.evaluate((stmt as { expression: Expression }).expression);
      } else {
        result = this.evaluate(stmt as unknown as Expression);
      }
    }
    return result;
  }

  evaluate(node: Expression): VbValue {
    switch (node.type) {
      case 'Identifier':
        return this.evaluateIdentifier(node);
      case 'Literal':
        return this.evaluateLiteral(node);
      case 'VbEmptyLiteral':
        return VbEmpty;
      case 'VbNewExpression':
        return this.evaluateNew(node);
      case 'ThisExpression':
        return this.evaluateMe(node);
      case 'VbWithObject':
        return this.evaluateWithObject(node);
      case 'MemberExpression':
        return this.evaluateMember(node);
      case 'CallExpression':
        return this.evaluateCallInternal(node.callee as Expression, node.arguments);
      case 'BinaryExpression':
        return this.evaluateBinary(node);
      case 'UnaryExpression':
        return this.evaluateUnary(node);
      case 'LogicalExpression':
        return this.evaluateLogical(node);
      case 'AssignmentExpression':
        return this.evaluateAssignment(node);
      case 'ConditionalExpression':
        return this.evaluateConditional(node);
      default:
        return VbEmpty;
    }
  }

  evaluateCall(callee: Expression, args: Expression[]): VbValue {
    return this.evaluateCallInternal(callee, args);
  }

  private evaluateIdentifier(node: Identifier): VbValue {
    const name = node.name;

    if (this.context.functionRegistry.has(name)) {
      return this.context.functionRegistry.call(name, []);
    }

    return this.context.getVariable(name);
  }

  private evaluateLiteral(node: Literal): VbValue {
    return createVbValue(node.value);
  }

  private evaluateNew(node: VbNewExpression): VbObjectValue {
    // Simple identifier: try VBScript class registry first, then globalThis fallback.
    if (node.callee.type === 'Identifier') {
      const className = node.callee.name;
      const ctorArgs = node.arguments.map(arg => this.evaluate(arg));
      try {
        const instance = this.context.classRegistry.createInstance(className, ctorArgs);
        return { type: 'Object', value: instance };
      } catch {
        // Not a VBScript class — fall through to globalThis lookup below.
      }
      const ctor = (globalThis as Record<string, unknown>)[className];
      if (typeof ctor === 'function') {
        const args = ctorArgs.map(v => this.vbToJs(v));
        return { type: 'Object', value: Reflect.construct(ctor as unknown as new (...a: unknown[]) => unknown, args) };
      }
      throw createVbError(VbErrorCodes.InvalidProcedureCall, `Unknown class: ${className}`, 'Vbscript');
    }

    // Dotted path (e.g. New Forms.Form, New NS.Sub.Class): evaluate the
    // member expression to obtain the JS constructor, then call it with new.
    const ctorValue = this.evaluate(node.callee);
    let ctor: ((...args: unknown[]) => unknown) | undefined;

    if (ctorValue.type === 'Object') {
      const obj = ctorValue.value as Record<string, unknown> | null;
      if (obj && obj.type === 'jsfunction') {
        ctor = (obj as { func: (...args: unknown[]) => unknown }).func;
      } else if (typeof obj === 'function') {
        ctor = obj as (...args: unknown[]) => unknown;
      } else if (typeof obj === 'object' && obj !== null) {
        // Raw JS Proxy (e.g. node-ps1-dotnet namespace constructor)
        ctor = obj as unknown as (...args: unknown[]) => unknown;
      }
    }

    if (typeof ctor !== 'function' && !(ctor && typeof ctor === 'object')) {
      throw createVbError(VbErrorCodes.InvalidProcedureCall, 'Not a constructor', 'Vbscript');
    }

    const args = node.arguments.map(arg => this.vbToJs(this.evaluate(arg)));
    return {
      type: 'Object',
      value: Reflect.construct(ctor as unknown as new (...a: unknown[]) => unknown, args),
    };
  }

  private evaluateMe(_node: ThisExpression): VbObjectValue {
    void _node; // Intentionally unused
    if (this.context.currentInstance) {
      return { type: 'Object', value: this.context.currentInstance };
    }
    throw new Error('Me keyword not supported outside of class context');
  }

  private evaluateWithObject(_node: VbWithObjectExpression): VbValue {
    void _node; // Intentionally unused
    const withObject = this.context.getCurrentWith();
    if (!withObject) {
      throw new Error('With object not available - must be inside a With statement');
    }
    return withObject;
  }

  private evaluateMember(node: MemberExpression): VbValue {
    let object = this.evaluate(node.object as Expression);

    if (node.object.type === 'VbWithObject') {
      object = this.evaluateWithObject(node.object as VbWithObjectExpression);
    }

    let propertyName: string;
    if (node.computed) {
      const propValue = this.evaluate(node.property as Expression);
      propertyName = toString(propValue);
    } else {
      propertyName = (node.property as Identifier).name;
    }

    if (object.type === 'Array') {
      return this.getArrayElement(object, node.property as Expression);
    }

    if (object.type === 'Object') {
      return this.getObjectProperty(object, propertyName);
    }

    throw createVbError(VbErrorCodes.ObjectRequired, 'Object required', 'Vbscript');
  }

  private getArrayElement(array: VbArrayValue, indexExpr: Expression): VbValue {
    const arr = array.value as unknown as { get: (indices: number[]) => VbValue };
    const index = toNumber(this.evaluate(indexExpr));
    return arr.get([Math.floor(index)]);
  }

  private getObjectProperty(objValue: VbObjectValue, propertyName: string): VbValue {
    const obj = objValue.value as VbObjectValueData | null;

    if (obj === null) {
      throw createVbError(VbErrorCodes.ObjectRequired, 'Object required', 'Vbscript');
    }

    // If this is a JS function wrapper produced by a previous property access,
    // allow further member access to flow through to the underlying JS function/
    // Proxy object. This is required for chains like form.Controls.Add(...), where
    // `Controls` may itself be represented by a callable function-proxy.
    if (isVbJsFunctionObject(obj) && !['type', 'func', 'thisArg'].includes(propertyName)) {
      const jsValue = ((obj.func as unknown) as Record<string, unknown>)[propertyName];
      if (jsValue === undefined) {
        return { type: 'Empty', value: undefined };
      }
      if (typeof jsValue === 'function') {
        return {
          type: 'Object',
          value: {
            type: 'jsfunction',
            func: jsValue as (...args: unknown[]) => unknown,
            thisArg: obj.func,
          },
        };
      }
      return this.jsToVb(jsValue);
    }

    if (
      (Object.prototype.hasOwnProperty.call(obj, 'getProperty') ||
        Object.prototype.hasOwnProperty.call(obj, 'classInfo')) &&
      typeof obj.getProperty === 'function'
    ) {
      if (obj.hasMethod?.(propertyName)) {
        return { type: 'Object', value: { type: 'method', object: obj, method: propertyName } };
      }
      return obj.getProperty(propertyName);
    }

    const jsValue = obj[propertyName];
    if (jsValue === undefined) {
      return { type: 'Empty', value: undefined };
    }
    if (typeof jsValue === 'function') {
      return {
        type: 'Object',
        value: {
          type: 'jsfunction',
          func: jsValue as (...args: unknown[]) => unknown,
          thisArg: obj,
        },
      };
    }
    return this.jsToVb(jsValue);
  }

  private jsToVb(value: unknown): VbValue {
    if (value === undefined) return { type: 'Empty', value: undefined };
    if (value === null) return { type: 'Null', value: null };
    if (typeof value === 'boolean') return { type: 'Boolean', value };
    if (typeof value === 'bigint') return { type: 'LongLong', value };
    if (typeof value === 'number') {
      if (Number.isInteger(value) && value >= -2147483648 && value <= 2147483647) {
        return { type: 'Long', value };
      }
      return { type: 'Double', value };
    }
    if (typeof value === 'string') return { type: 'String', value };
    if (value instanceof Date) return { type: 'Date', value };
    if (Array.isArray(value)) {
      return { type: 'Array', value };
    }
    if (typeof value === 'object') {
      return { type: 'Object', value: value as VbObjectValueData };
    }
    return { type: 'String', value: String(value) };
  }

  private evaluateCallInternal(calleeExpr: Expression, callArgs: Expression[]): VbValue {
    if (calleeExpr.type === 'Identifier') {
      const name = calleeExpr.name;

      if (this.context.functionRegistry.has(name)) {
        const funcInfo = this.context.functionRegistry.get(name)!;
        const hasByRefParams = funcInfo.params.some(p => p.byRef);

        if (hasByRefParams) {
          const argRefs = callArgs.map(arg => {
            if (arg.type === 'Identifier') {
              const varName = arg.name;
              const variable = this.context.currentScope.get(varName);
              return {
                value: variable ? variable.value : this.context.getVariable(varName),
                variableName: varName,
                setValue: (val: VbValue) => {
                  if (this.context.currentScope.has(varName)) {
                    this.context.currentScope.set(varName, val);
                  } else {
                    this.context.setVariable(varName, val);
                  }
                },
              };
            }
            return { value: this.evaluate(arg) };
          });
          return this.context.functionRegistry.callWithRefs(name, argRefs);
        }

        const args = callArgs.map(arg => this.evaluate(arg));
        return this.context.functionRegistry.call(name, args);
      }

      const variable = this.context.currentScope.get(name);
      if (variable) {
        if (variable.value.type === 'Array') {
          const indices = callArgs.map(arg => Math.floor(toNumber(this.evaluate(arg))));
          const arr = variable.value.value as unknown as { get: (indices: number[]) => VbValue };
          return arr.get(indices);
        }
        if (variable.value.type === 'Object' && variable.value.value !== null) {
          const val = variable.value.value as VbObjectValueData;
          if (val.hasMethod?.('default') && val.getMethod) {
            const method = val.getMethod('default');
            const args = callArgs.map(arg => this.evaluate(arg));
            return method.func(...args);
          }
        }
      }

      const callee = this.context.getVariable(name);
      if (callee.type === 'Object' && callee.value !== null) {
        return this.callObjectMethod(callee, callArgs);
      }
      throw createVbError(VbErrorCodes.InvalidProcedureCall, 'Invalid procedure call', 'Vbscript');
    }

    if (calleeExpr.type === 'MemberExpression') {
      const memberResult = this.evaluateMember(calleeExpr);

      if (memberResult.type === 'Object' && memberResult.value !== null) {
        return this.callObjectMethod(memberResult, callArgs);
      }
      throw createVbError(VbErrorCodes.InvalidProcedureCall, 'Invalid procedure call', 'Vbscript');
    }

    const callee = this.evaluate(calleeExpr);
    if (callee.type === 'Object' && callee.value !== null) {
      return this.callObjectMethod(callee, callArgs);
    }
    throw createVbError(VbErrorCodes.InvalidProcedureCall, 'Invalid procedure call', 'Vbscript');
  }

  private callObjectMethod(objValue: VbObjectValue, callArgs: Expression[]): VbValue {
    const obj = objValue.value as VbObjectValueData | null;
    if (obj === null) {
      throw createVbError(VbErrorCodes.ObjectRequired, 'Object required', 'Vbscript');
    }

    if (isVbMethodObject(obj)) {
      const args = callArgs.map(arg => this.evaluate(arg));
      return obj.object.getMethod(obj.method).func.call(obj.object, ...args);
    }

    if (isVbJsFunctionObject(obj)) {
      const args = callArgs.map(arg => {
        const vbVal = this.evaluate(arg);
        return this.vbToJs(vbVal);
      });
      const result = obj.func.call(obj.thisArg, ...args);
      return this.jsToVb(result);
    }

    if (typeof obj.call === 'function') {
      const args = callArgs.map(arg => this.evaluate(arg));
      return obj.call(...args);
    }

    throw createVbError(VbErrorCodes.InvalidProcedureCall, 'Invalid procedure call', 'Vbscript');
  }

  private vbToJs(value: VbValue): unknown {
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
        return value.value;
      case 'LongLong':
        return value.value; // return as bigint
      case 'Date':
        return value.value instanceof Date ? value.value : new Date(value.value);
      case 'Array':
        return value.value;
      case 'Object': {
        const obj = value.value as VbObjectValueData | null;
        if (obj && typeof obj === 'object' && (obj as Record<string, unknown>).type === 'vbref') {
          const callFn = (obj as Record<string, unknown>).call as ((...args: VbValue[]) => VbValue) | undefined;
          if (typeof callFn === 'function') {
            return (...jsArgs: unknown[]): unknown => {
              const vbArgs = jsArgs.map((a) => this.jsToVb(a));
              return this.vbToJs(callFn(...vbArgs));
            };
          }
        }
        return value.value;
      }
      default:
        return value.value;
    }
  }

  private evaluateBinary(node: BinaryExpression): VbValue {
    const left = this.evaluate(node.left);
    const right = this.evaluate(node.right);

    switch (node.operator) {
      case '+':
        return this.add(left, right);
      case '-':
        return this.subtract(left, right);
      case '*':
        return this.multiply(left, right);
      case '/':
        return this.divide(left, right);
      case '\\':
        if (left.type === 'LongLong' || right.type === 'LongLong') {
          return this.longLongArithmetic(left, right, '\\');
        }
        return this.integerDivide(left, right);
      case '%':
      case 'Mod':
        return this.modulo(left, right);
      case '**':
      case '^':
        return this.power(left, right);
      case '==':
      case '=':
        return this.equals(left, right);
      case '!=':
      case '<>':
        return this.notEquals(left, right);
      case '<':
        return this.lessThan(left, right);
      case '<=':
        return this.lessThanOrEqual(left, right);
      case '>':
        return this.greaterThan(left, right);
      case '>=':
        return this.greaterThanOrEqual(left, right);
      case '&':
        return { type: 'String', value: toString(left) + toString(right) };
      case 'Is':
        return {
          type: 'Boolean',
          value: left.type === 'Object' && right.type === 'Object' && left.value === right.value,
        };
      default:
        return VbEmpty;
    }
  }

  private add(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    if (left.type === 'String' || right.type === 'String') {
      return { type: 'String', value: toString(left) + toString(right) };
    }
    if (left.type === 'LongLong' || right.type === 'LongLong') {
      return this.longLongArithmetic(left, right, '+');
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    return createVbValue(leftNum + rightNum);
  }

  private subtract(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    if (left.type === 'LongLong' || right.type === 'LongLong') {
      return this.longLongArithmetic(left, right, '-');
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    return createVbValue(leftNum - rightNum);
  }

  private multiply(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    if (left.type === 'LongLong' || right.type === 'LongLong') {
      return this.longLongArithmetic(left, right, '*');
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    return createVbValue(leftNum * rightNum);
  }

  private divide(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    if (rightNum === 0) {
      throw createVbError(VbErrorCodes.DivisionByZero, 'Division by zero', 'Vbscript');
    }
    return { type: 'Double', value: leftNum / rightNum };
  }

  private integerDivide(left: VbValue, right: VbValue): VbLongValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull as unknown as VbLongValue;
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    if (rightNum === 0) {
      throw createVbError(VbErrorCodes.DivisionByZero, 'Division by zero', 'Vbscript');
    }
    return { type: 'Long', value: Math.floor(leftNum / rightNum) };
  }

  private modulo(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    if (rightNum === 0) {
      throw createVbError(VbErrorCodes.DivisionByZero, 'Division by zero', 'Vbscript');
    }
    return createVbValue(leftNum % rightNum);
  }

  private power(left: VbValue, right: VbValue): VbValue {
    if (left.type === 'Null' || right.type === 'Null') {
      return VbNull;
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    return { type: 'Double', value: Math.pow(leftNum, rightNum) };
  }

  private equals(left: VbValue, right: VbValue): VbBooleanValue {
    if (left.type === 'Empty' && right.type === 'Empty') {
      return { type: 'Boolean', value: true };
    }
    if (left.type === 'Null' || right.type === 'Null') {
      return { type: 'Boolean', value: false };
    }
    if (left.type === 'Object' && right.type === 'Object') {
      return { type: 'Boolean', value: left.value === right.value };
    }
    if (left.type === 'String' || right.type === 'String') {
      return {
        type: 'Boolean',
        value: toString(left).toLowerCase() === toString(right).toLowerCase(),
      };
    }
    return { type: 'Boolean', value: toNumber(left) === toNumber(right) };
  }

  private notEquals(left: VbValue, right: VbValue): VbBooleanValue {
    const eq = this.equals(left, right);
    return { type: 'Boolean', value: !eq.value };
  }

  private lessThan(left: VbValue, right: VbValue): VbBooleanValue {
    if (left.type === 'String' || right.type === 'String') {
      return { type: 'Boolean', value: toString(left) < toString(right) };
    }
    return { type: 'Boolean', value: toNumber(left) < toNumber(right) };
  }

  private lessThanOrEqual(left: VbValue, right: VbValue): VbBooleanValue {
    if (left.type === 'String' || right.type === 'String') {
      return { type: 'Boolean', value: toString(left) <= toString(right) };
    }
    return { type: 'Boolean', value: toNumber(left) <= toNumber(right) };
  }

  private greaterThan(left: VbValue, right: VbValue): VbBooleanValue {
    if (left.type === 'String' || right.type === 'String') {
      return { type: 'Boolean', value: toString(left) > toString(right) };
    }
    return { type: 'Boolean', value: toNumber(left) > toNumber(right) };
  }

  private greaterThanOrEqual(left: VbValue, right: VbValue): VbBooleanValue {
    if (left.type === 'String' || right.type === 'String') {
      return { type: 'Boolean', value: toString(left) >= toString(right) };
    }
    return { type: 'Boolean', value: toNumber(left) >= toNumber(right) };
  }

  private evaluateUnary(node: UnaryExpression): VbValue {
    const argument = this.evaluate(node.argument);

    switch (node.operator) {
      case '-':
        return createVbValue(-toNumber(argument));
      case '+':
        return createVbValue(toNumber(argument));
      case '!':
      case 'Not':
        return { type: 'Boolean', value: !toBoolean(argument) };
      default:
        return VbEmpty;
    }
  }

  private evaluateLogical(node: LogicalExpression): VbValue {
    switch (node.operator) {
      case '&&':
      case 'And': {
        const left = this.evaluate(node.left);
        if (!toBoolean(left)) return left;
        return this.evaluate(node.right);
      }
      case '||':
      case 'Or': {
        const left = this.evaluate(node.left);
        if (toBoolean(left)) return left;
        return this.evaluate(node.right);
      }
      case 'Xor': {
        const left = this.evaluate(node.left);
        const right = this.evaluate(node.right);
        return { type: 'Boolean', value: toBoolean(left) !== toBoolean(right) };
      }
      case 'Eqv': {
        const left = this.evaluate(node.left);
        const right = this.evaluate(node.right);
        return { type: 'Boolean', value: toBoolean(left) === toBoolean(right) };
      }
      case 'Imp': {
        const left = this.evaluate(node.left);
        const right = this.evaluate(node.right);
        return { type: 'Boolean', value: !toBoolean(left) || toBoolean(right) };
      }
      default:
        return VbEmpty;
    }
  }

  private evaluateAssignment(node: AssignmentExpression): VbValue {
    const value = this.evaluate(node.right);
    const isSet = node.isSet ?? false;

    if (node.left.type === 'Identifier') {
      if (isSet) {
        const oldValue = this.context.getVariable(node.left.name);
        this.callTerminateIfNeeded(oldValue, value);
      }
      this.context.setVariable(node.left.name, value);
      return value;
    }

    if (node.left.type === 'MemberExpression') {
      this.assignToMember(node.left, value, isSet);
      return value;
    }

    throw new Error(`Invalid assignment target: ${node.left.type}`);
  }

  private callTerminateIfNeeded(oldValue: VbValue, newValue: VbValue): void {
    if (oldValue.type === 'Object' && oldValue.value instanceof VbObjectInstance) {
      if (newValue.type !== 'Object' || newValue.value !== oldValue.value) {
        const instance = oldValue.value as VbObjectInstance;
        const terminateProp = instance.classInfo.properties.get('class_terminate');
        if (terminateProp && terminateProp.get) {
          terminateProp.get.call(instance);
        }
      }
    }
  }

  private assignToMember(node: MemberExpression, value: VbValue, isSet: boolean): void {
    let object = this.evaluate(node.object as Expression);

    if (node.object.type === 'VbWithObject') {
      object = this.evaluateWithObject(node.object as VbWithObjectExpression);
    }

    let propertyName: string;
    if (node.computed) {
      const propValue = this.evaluate(node.property as Expression);
      propertyName = toString(propValue);
    } else {
      propertyName = (node.property as Identifier).name;
    }

    if (object.type === 'Array') {
      const arr = object.value as unknown as { set: (indices: number[], v: VbValue) => void };
      const index = toNumber(this.evaluate(node.property as Expression));
      arr.set([Math.floor(index)], value);
    } else if (object.type === 'Object') {
      const obj = object.value as VbObjectValueData | null;
      if (obj === null || (typeof obj !== 'object' && typeof obj !== 'function')) {
        throw createVbError(VbErrorCodes.ObjectRequired, 'Object required', 'Vbscript');
      }

      if (
        (Object.prototype.hasOwnProperty.call(obj, 'setProperty') ||
          Object.prototype.hasOwnProperty.call(obj, 'classInfo')) &&
        typeof obj.setProperty === 'function'
      ) {
        obj.setProperty(propertyName, value, isSet);
      } else {
        const jsValue = this.vbToJs(value);
        obj[propertyName] = jsValue;
      }
    } else {
      throw createVbError(VbErrorCodes.ObjectRequired, 'Object required', 'Vbscript');
    }
  }

  private evaluateConditional(node: ConditionalExpression): VbValue {
    const test = this.evaluate(node.test);
    if (toBoolean(test)) {
      return this.evaluate(node.consequent);
    }
    return this.evaluate(node.alternate);
  }

  private toLongLongBigInt(v: VbValue): bigint {
    if (v.type === 'LongLong') return v.value;
    if (v.type === 'Integer' || v.type === 'Long' || v.type === 'Byte') return BigInt(v.value);
    if (v.type === 'Boolean') return v.value ? BigInt(-1) : BigInt(0);
    if (v.type === 'Empty') return BigInt(0);
    // Float types demote to Double for mixed arithmetic
    return BigInt(Math.trunc(toNumber(v)));
  }

  private longLongArithmetic(left: VbValue, right: VbValue, op: '+' | '-' | '*' | '\\'): VbValue {
    // If either side is floating-point, promote to Double
    const floatTypes = ['Single', 'Double', 'Currency'];
    if (floatTypes.includes(left.type) || floatTypes.includes(right.type)) {
      const l = toNumber(left);
      const r = toNumber(right);
      if (op === '+') return { type: 'Double', value: l + r };
      if (op === '-') return { type: 'Double', value: l - r };
      if (op === '*') return { type: 'Double', value: l * r };
      if (r === 0) throw createVbError(VbErrorCodes.DivisionByZero, 'Division by zero', 'Vbscript');
      return { type: 'Long', value: Math.floor(l / r) };
    }
    const l = this.toLongLongBigInt(left);
    const r = this.toLongLongBigInt(right);
    if (op === '+') return { type: 'LongLong', value: l + r };
    if (op === '-') return { type: 'LongLong', value: l - r };
    if (op === '*') return { type: 'LongLong', value: l * r };
    if (r === BigInt(0)) throw createVbError(VbErrorCodes.DivisionByZero, 'Division by zero', 'Vbscript');
    return { type: 'LongLong', value: l / r };
  }
}
