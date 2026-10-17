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
  NewExpression,
  VbWithObjectExpression,
} from "../ast/index.ts";
import type {
  VbValue,
  VbObjectValue,
  VbObjectValueData,
  VbArrayValue,
} from "../runtime/index.ts";
import { isVbClassInstance } from "../runtime/index.ts";
import { VbContext, VbObjectInstance } from "../runtime/index.ts";
import {
  createVbValue,
  VbEmpty,
  VbNull,
  VbNothing,
  toBoolean,
  toNumber,
  toString,
  createVbError,
  VbErrorCodes,
  serialToDate,
  dateToSerial,
  currencyToNumber,
  parseVbNumericString,
  roundHalfToEven,
  toSubscript,
  toCurrencyScaled,
  toSinglePrecision,
  isPlainNumber,
  checkCurrencyRange,
  VbArray,
  checkLongLongRange,
  CURRENCY_SCALE,
  retainValue,
  releaseValue,
  releaseMadeValue,
  isMadeValue,
} from "../runtime/index.ts";

interface VbMethodObject {
  type: "method";
  object: VbObjectValueData & {
    getMethod: (name: string) => { func: (...args: VbValue[]) => VbValue };
  };
  method: string;
}

interface VbJsFunctionObject {
  type: "jsfunction";
  func: (...args: unknown[]) => unknown;
  thisArg: unknown;
}

function isVbMethodObject(
  obj: VbObjectValueData,
): obj is VbObjectValueData & VbMethodObject {
  return obj.type === "method" && "object" in obj && "method" in obj;
}

function isVbJsFunctionObject(
  obj: VbObjectValueData,
): obj is VbObjectValueData & VbJsFunctionObject {
  return obj.type === "jsfunction" && "func" in obj;
}

/**
 * "Invalid procedure call" on its own fits every call in a ninety-line procedure. Both throw
 * sites that raise it already hold the name they failed on, so it goes into the message.
 */
function invalidProcedureCall(name: string) {
  return createVbError(
    VbErrorCodes.InvalidProcedureCall,
    name ? `Invalid procedure call '${name}'` : "Invalid procedure call",
  );
}

/** `o.M` for the common shape; the empty string when neither half is a plain name. */
function memberExpressionName(node: MemberExpression): string {
  const objectName =
    node.object.type === "Identifier"
      ? `${(node.object as Identifier).name}.`
      : "";
  const propertyName =
    node.property.type === "Identifier"
      ? (node.property as Identifier).name
      : "";
  return objectName + propertyName;
}

// Real VBScript's And/Or/Xor/Eqv/Imp/Not all ultimately call the real OLE Automation
// VarAnd/VarOr/VarXor/VarEqv/VarImp/VarNot - genuine BITWISE operations, not JS-style boolean logic - with
// two VBScript-specific pre-coercions those native functions do not do themselves: Empty becomes Long(0), and
// a String becomes a number if it parses as one. A String that is not a number is a type mismatch, which a
// fresh cscript shows for all five: "abc" And Empty is error 13.
//
// The type the answer keeps is the wider of the two operands, and it is narrower than the engine used to
// answer. Measured across the whole operator matrix under 1033: Byte And Byte is a Byte, Byte And Integer an
// Integer, Boolean And Byte an Integer, and everything else, a Long, a floating point value, a Date, a
// numeric String or an Empty included, answers a Long. The narrower types wrap: Byte Eqv Byte is the Byte
// 255 where the whole number is -1.
//
// A floating point operand is rounded to the nearest whole number first, a tie going to the even neighbour,
// so CByte(200) Or CSng(1.5) is 202 and not the 201 a truncation gives.

/** The type a bitwise operand keeps in the answer, from the narrowest to the widest. */
type BitwiseKind = "Boolean" | "Byte" | "Integer" | "Long";

/** A bitwise operand: the whole number the operation uses, and the type its own answer keeps. */
interface BitwiseOperand {
  kind: BitwiseKind;
  value: number;
}

/**
 * The whole number a bitwise operator reads an operand as, or a type mismatch for a text that is no number.
 *
 * `emptyKind` is the type an Empty counts as, which is not the same for every operator family: it is a Long
 * for the five bitwise operators (Empty And Boolean is a Long) and an Integer for integer division and the
 * remainder (Empty \ Boolean is an Integer).
 */
function readBitwiseOperand(
  value: VbValue,
  emptyKind: BitwiseKind = "Long",
): BitwiseOperand {
  switch (value.type) {
    case "Boolean":
      return { kind: "Boolean", value: value.value ? -1 : 0 };
    case "Byte":
      return { kind: "Byte", value: value.value };
    case "Integer":
      return { kind: "Integer", value: value.value };
    case "Long":
      return { kind: "Long", value: value.value };
    case "LongLong":
      return { kind: "Long", value: Number(BigInt.asIntN(32, value.value)) };
    case "Empty":
      return { kind: emptyKind, value: 0 };
    case "Single":
    case "Double":
      return { kind: "Long", value: roundHalfToEven(value.value) };
    case "Currency":
      return {
        kind: "Long",
        value: roundHalfToEven(currencyToNumber(value.value)),
      };
    case "Date":
      return {
        kind: "Long",
        value: roundHalfToEven(dateToSerial(value.value)),
      };
    case "String": {
      const parsed = parseVbNumericString(value.value.trim());
      if (parsed === null)
        throw createVbError(VbErrorCodes.TypeMismatch, "Type mismatch");
      return { kind: "Long", value: roundHalfToEven(parsed) };
    }
    default:
      // An Object operand has no number here, since this engine does not yet run its default property. The
      // older lenient reading is kept rather than throwing, which only affects that unfinished feature.
      return { kind: "Long", value: toBoolean(value) ? -1 : 0 };
  }
}

/** Whether a value is a number proper, which is what a text beside it is read as one for. */
function isNumberValue(value: VbValue): boolean {
  return (
    value.type === "Integer" ||
    value.type === "Long" ||
    value.type === "LongLong" ||
    value.type === "Double" ||
    value.type === "Single" ||
    value.type === "Byte" ||
    value.type === "Currency"
  );
}

/** The order of two numbers, which is all a comparison needs of them. */
function compareNumbers(left: number, right: number): number {
  if (left === right) return 0;
  return left < right ? -1 : 1;
}

/** The type a bitwise answer keeps: the wider of the two operands. */
function bitwiseKind(left: BitwiseKind, right: BitwiseKind): BitwiseKind {
  if (left === right) return left;
  const narrow = (kind: BitwiseKind): boolean =>
    kind === "Boolean" || kind === "Byte" || kind === "Integer";
  return narrow(left) && narrow(right) ? "Integer" : "Long";
}

/** The operators that compare two values rather than working on them. */
const COMPARISON_OPERATORS = new Set([
  "==",
  "=",
  "!=",
  "<>",
  "<",
  "<=",
  ">",
  ">=",
]);

/** Whether an expression is a number written out in the source, which reads differently beside a text. */
function isNumberLiteralNode(node: Expression): boolean {
  const candidate = node as { type?: string; value?: unknown };
  return candidate.type === "Literal" && typeof candidate.value === "number";
}

/**
 * The type integer division and the remainder keep: the wider of the two again, but a Boolean counts as an
 * Integer here. Measured: Boolean \ Boolean is the Integer 1, where Boolean And Boolean is the Boolean True.
 */
function integerResultKind(left: BitwiseKind, right: BitwiseKind): BitwiseKind {
  const asInteger = (kind: BitwiseKind): BitwiseKind =>
    kind === "Boolean" ? "Integer" : kind;
  const first = asInteger(left);
  const second = asInteger(right);
  if (first === second) return first;
  return first === "Long" || second === "Long" ? "Long" : "Integer";
}

/** A whole number as the type it is kept in, wrapped the way the narrower types wrap. */
function bitwiseValue(kind: BitwiseKind, value: number): VbValue {
  switch (kind) {
    case "Boolean":
      return { type: "Boolean", value: value !== 0 };
    case "Byte":
      return { type: "Byte", value: ((value % 256) + 256) % 256 };
    case "Integer":
      return {
        type: "Integer",
        value: ((((value + 32768) % 65536) + 65536) % 65536) - 32768,
      };
    default:
      return { type: "Long", value: value | 0 };
  }
}

/**
 * One operand beside a Null, which each of the five operators reads its own way.
 *
 * Measured across the matrix: And and Or keep the side that already decides the answer (False And Null is 0,
 * True Or Null is the truthy side itself, its own type and value), Xor and Eqv are Null, and Imp answers the
 * truthy side when the Null is on the left but the complement of it when the Null is on the right, unless
 * that side is a Boolean, where it is Null.
 */
function bitwiseWithNull(
  operator: "And" | "Or" | "Xor" | "Eqv" | "Imp",
  operand: BitwiseOperand,
  nullIsLeft: boolean,
): VbValue {
  const truthy = operand.value !== 0;
  switch (operator) {
    case "And":
      return truthy ? VbNull : bitwiseValue(operand.kind, 0);
    case "Or":
      return truthy ? bitwiseValue(operand.kind, operand.value) : VbNull;
    case "Imp":
      if (nullIsLeft)
        return truthy ? bitwiseValue(operand.kind, operand.value) : VbNull;
      if (operand.kind === "Boolean") return VbNull;
      return bitwiseValue(operand.kind, ~operand.value);
    default:
      return VbNull;
  }
}

/** An array is copied where it is assigned, so a copy is what goes into the variable. */
function cloneIfArray(value: VbValue): VbValue {
  if (value.type !== "Array") return value;
  const array = value.value as VbArray;
  return { type: "Array", value: array.clone() };
}

export class ExpressionEvaluator {
  /**
   * Whether the member read that just finished was a Property Get, which a bare statement reads and drops.
   *
   * Set by getObjectProperty and read straight after it by the member call path, with nothing in between that
   * reads a member of its own beyond the arguments a call with no arguments does not have.
   */
  private lastMemberWasGetter = false;

  constructor(private context: VbContext) {}

  evaluateProgram(program: Program): VbValue {
    let result: VbValue = VbEmpty;
    for (const stmt of program.body) {
      if (stmt.type === "ExpressionStatement") {
        result = this.evaluate((stmt as { expression: Expression }).expression);
      } else {
        result = this.evaluate(stmt as unknown as Expression);
      }
    }
    return result;
  }

  evaluate(node: Expression): VbValue {
    switch (node.type) {
      case "Identifier":
        return this.evaluateIdentifier(node);
      case "Literal":
        return this.evaluateLiteral(node);
      case "VbEmptyLiteral":
        return VbEmpty;
      case "NewExpression":
        return this.evaluateNew(node);
      case "ThisExpression":
        return this.evaluateMe(node);
      case "VbWithObject":
        return this.evaluateWithObject(node);
      case "MemberExpression":
        return this.evaluateMemberAsValue(node);
      case "CallExpression":
        return this.evaluateCallInternal(
          node.callee as Expression,
          node.arguments,
          node.statementCall,
          node.byValueArguments ?? false,
          node.bareName ?? false,
        );
      case "BinaryExpression":
        return this.evaluateBinary(node);
      case "UnaryExpression":
        return this.evaluateUnary(node);
      case "LogicalExpression":
        return this.evaluateLogical(node);
      case "AssignmentExpression":
        return this.evaluateAssignment(node);
      case "ConditionalExpression":
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

    // A local variable/parameter (or class instance property) must shadow a
    // same-named built-in - real VBScript allows `Function trackEval(val,
    // label)` to use `val` as an ordinary parameter without it colliding
    // with the built-in Val() function. This used to check the function
    // registry FIRST unconditionally, so reading a parameter/variable that
    // happened to share a built-in's name (found via Wine's own
    // vbscript.dll conformance suite, dlls/vbscript/tests/lang.vbs: several
    // helper functions use `val` as a parameter name) silently called the
    // built-in with zero arguments instead - Val() with no argument then
    // crashed converting `undefined` to a string.
    if (this.context.hasDeclaredVariable(name)) {
      return this.context.getVariable(name);
    }

    if (this.context.functionRegistry.has(name)) {
      // A Sub has no value to give, so naming one where a value is expected is a type mismatch, measured on a
      // fresh cscript: `result = NoArgs` is 13 for a Sub of any arity, while `result = NoArgFunction` is 0 and
      // `result = OneArgFunction` is the 450 that Function raises when it is called with no arguments.
      const info = this.context.functionRegistry.get(name);
      if (info?.isSub === true) {
        throw createVbError(
          VbErrorCodes.TypeMismatch,
          `Type mismatch: '${name}'`,
        );
      }
      return this.context.functionRegistry.call(name, []);
    }

    return this.context.getVariable(name);
  }

  private evaluateLiteral(node: Literal): VbValue {
    return createVbValue(node.value);
  }

  private evaluateNew(node: NewExpression): VbObjectValue {
    // Simple identifier: try VBScript class registry first, then globalThis fallback.
    if (node.callee.type === "Identifier") {
      const className = node.callee.name;
      const ctorArgs = node.arguments.map((arg) => this.evaluate(arg));
      if (this.context.classRegistry.has(className)) {
        // The initializer's own failure is the script's failure, so it is not caught here: only a name that
        // is no registered class falls through to the host lookup below.
        return {
          type: "Object",
          value: this.context.classRegistry.createInstance(className, ctorArgs),
        };
      }
      const ctor = (globalThis as Record<string, unknown>)[className];
      if (typeof ctor === "function") {
        const args = ctorArgs.map((v) => this.vbToJs(v));
        return {
          type: "Object",
          value: Reflect.construct(
            ctor as unknown as new (...a: unknown[]) => unknown,
            args,
          ),
        };
      }
      // A class that is nowhere is error 506, measured: `result = New NoSuchClass` answers 506.
      throw createVbError(
        VbErrorCodes.ClassNotDefined,
        `Class not defined: ${className}`,
      );
    }

    // Dotted path (e.g. New Forms.Form, New NS.Sub.Class): evaluate the
    // member expression to obtain the JS constructor, then call it with new.
    const ctorValue = this.evaluate(node.callee);
    let ctor: ((...args: unknown[]) => unknown) | undefined;

    if (ctorValue.type === "Object") {
      const obj = ctorValue.value as Record<string, unknown> | null;
      if (obj && obj.type === "jsfunction") {
        ctor = (obj as { func: (...args: unknown[]) => unknown }).func;
      } else if (typeof obj === "function") {
        ctor = obj as (...args: unknown[]) => unknown;
      } else if (typeof obj === "object" && obj !== null) {
        // Raw JS Proxy (e.g. node-ps1-dotnet namespace constructor)
        ctor = obj as unknown as (...args: unknown[]) => unknown;
      }
    }

    if (typeof ctor !== "function" && !(ctor && typeof ctor === "object")) {
      throw createVbError(
        VbErrorCodes.InvalidProcedureCall,
        "Not a constructor",
      );
    }

    const args = node.arguments.map((arg) => this.vbToJs(this.evaluate(arg)));
    return {
      type: "Object",
      value: Reflect.construct(
        ctor as unknown as new (...a: unknown[]) => unknown,
        args,
      ),
    };
  }

  private evaluateMe(_node: ThisExpression): VbObjectValue {
    void _node; // Intentionally unused
    if (this.context.currentInstance) {
      return { type: "Object", value: this.context.currentInstance };
    }
    throw new Error("Me keyword not supported outside of class context");
  }

  private evaluateWithObject(_node: VbWithObjectExpression): VbValue {
    void _node; // Intentionally unused
    const withObject = this.context.getCurrentWith();
    if (!withObject) {
      throw new Error(
        "With object not available - must be inside a With statement",
      );
    }
    return withObject;
  }

  /**
   * A member read written WITHOUT parentheses, reached where a value is wanted. In VBScript such
   * a read of a zero-argument Sub or Function IS the call: `r = o.F` yields F's result, and a bare
   * `o.M` on its own line runs M.
   *
   * Only this dispatch is affected. `o.F(args)` reaches evaluateCallInternal(), which calls
   * evaluateMember() directly and still sees the unresolved method reference it needs, so real
   * arguments continue to reach the real call.
   */
  private evaluateMemberAsValue(node: MemberExpression): VbValue {
    const value = this.evaluateMember(node);

    if (value.type === "Object" && value.value !== null) {
      const obj = value.value as VbObjectValueData;
      if (isVbMethodObject(obj) || isVbJsFunctionObject(obj)) {
        // Goes through the ordinary call path with no arguments, so a class method and a host
        // function exposed via addObject() are invoked exactly as they are when written with
        // parentheses -- including the jsToVb conversion a host function's result needs.
        return this.callObjectMethod(value as VbObjectValue, []);
      }
    }

    return value;
  }

  private evaluateMember(node: MemberExpression): VbValue {
    let object = this.evaluate(node.object as Expression);

    if (node.object.type === "VbWithObject") {
      object = this.evaluateWithObject(node.object as VbWithObjectExpression);
    }

    // An element of an array is read before the member name is, for the same reason the write side does it:
    // the name of a computed member is an expression, and `a(Null)` is the subscript's error 13.
    if (object.type === "Array") {
      return this.getArrayElement(object, node.property as Expression);
    }

    let propertyName: string;
    if (node.computed) {
      const propValue = this.evaluate(node.property as Expression);
      propertyName = toString(propValue);
    } else {
      propertyName = (node.property as Identifier).name;
    }

    if (object.type === "Object") {
      return this.getObjectProperty(object, propertyName);
    }

    throw createVbError(VbErrorCodes.ObjectRequired, "Object required");
  }

  private getArrayElement(array: VbArrayValue, indexExpr: Expression): VbValue {
    const arr = array.value as unknown as {
      get: (indices: number[]) => VbValue;
    };
    const index = toSubscript(this.evaluate(indexExpr));
    return arr.get([index]);
  }

  private getObjectProperty(
    objValue: VbObjectValue,
    propertyName: string,
  ): VbValue {
    const obj = objValue.value as VbObjectValueData | null;

    if (obj === null) {
      throw createVbError(VbErrorCodes.ObjectRequired, "Object required");
    }

    // If this is a JS function wrapper produced by a previous property access,
    // allow further member access to flow through to the underlying JS function/
    // Proxy object. This is required for chains like form.Controls.Add(...), where
    // `Controls` may itself be represented by a callable function-proxy.
    if (
      isVbJsFunctionObject(obj) &&
      !["type", "func", "thisArg"].includes(propertyName)
    ) {
      const jsValue = (obj.func as unknown as Record<string, unknown>)[
        propertyName
      ];
      if (jsValue === undefined) {
        return { type: "Empty", value: undefined };
      }
      if (typeof jsValue === "function") {
        return {
          type: "Object",
          value: {
            type: "jsfunction",
            func: jsValue as (...args: unknown[]) => unknown,
            thisArg: obj.func,
          },
        };
      }
      return this.jsToVb(jsValue);
    }

    if (
      (Object.prototype.hasOwnProperty.call(obj, "getProperty") ||
        Object.prototype.hasOwnProperty.call(obj, "classInfo")) &&
      typeof obj.getProperty === "function"
    ) {
      // A class instance answers 438 for a member it does not have, and for a Private member reached from
      // outside the class, which a fresh cscript shows for first.NoSuchMember and for first.hiddenCount
      // alike. A host object keeps the older, lenient reading.
      this.lastMemberWasGetter = false;
      if (isVbClassInstance(obj)) {
        const instance = obj as unknown as VbObjectInstance;
        if (
          !instance.isMemberVisible(propertyName, this.context.currentInstance)
        ) {
          throw createVbError(
            VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
            `Object doesn't support this property or method: '${propertyName}'`,
          );
        }
        if (
          !instance.hasMethod(propertyName) &&
          !instance.hasMember(propertyName)
        ) {
          throw createVbError(
            VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
            `Object doesn't support this property or method: '${propertyName}'`,
          );
        }
        this.lastMemberWasGetter = instance.hasGetter(propertyName);
      }
      if (obj.hasMethod?.(propertyName)) {
        return {
          type: "Object",
          value: { type: "method", object: obj, method: propertyName },
        };
      }
      return obj.getProperty(propertyName);
    }

    const jsValue = obj[propertyName];
    if (jsValue === undefined) {
      return { type: "Empty", value: undefined };
    }
    if (typeof jsValue === "function") {
      return {
        type: "Object",
        value: {
          type: "jsfunction",
          func: jsValue as (...args: unknown[]) => unknown,
          thisArg: obj,
        },
      };
    }
    return this.jsToVb(jsValue);
  }

  private jsToVb(value: unknown): VbValue {
    if (value === undefined) return { type: "Empty", value: undefined };
    if (value === null) return { type: "Null", value: null };
    if (value === Symbol.for("Nothing")) return VbNothing;
    if (typeof value === "boolean") return { type: "Boolean", value };
    if (typeof value === "bigint") return { type: "LongLong", value };
    if (typeof value === "number") {
      if (
        Number.isInteger(value) &&
        value >= -2147483648 &&
        value <= 2147483647
      ) {
        return { type: "Long", value };
      }
      return { type: "Double", value };
    }
    if (typeof value === "string") return { type: "String", value };
    if (value instanceof Date) return { type: "Date", value };
    if (Array.isArray(value)) {
      return { type: "Array", value };
    }
    if (typeof value === "object") {
      return { type: "Object", value: value as VbObjectValueData };
    }
    return { type: "String", value: String(value) };
  }

  private evaluateCallInternal(
    calleeExpr: Expression,
    callArgs: Expression[],
    statementCall = false,
    byValueArguments = false,
    bareName = false,
  ): VbValue {
    if (calleeExpr.type === "Identifier") {
      const name = calleeExpr.name;

      if (this.context.functionRegistry.has(name)) {
        const funcInfo = this.context.functionRegistry.get(name)!;
        const hasByRefParams = funcInfo.params.some((p) => p.byRef);

        if (hasByRefParams) {
          const argRefs = callArgs.map((arg) => {
            // An argument a statement call wrote in parentheses goes over by value, so the procedure may change
            // its own copy only. Measured: `FillWithNine(a)` leaves the caller's array alone where
            // `FillWithNine a` and `Call FillWithNine(a)` both change it.
            if (arg.type === "Identifier" && !byValueArguments) {
              const varName = arg.name;
              const variable = this.context.currentScope.get(varName);
              return {
                value: variable
                  ? variable.value
                  : this.context.getVariable(varName),
                variableName: varName,
                isArrayVariable: variable?.isArray === true,
                setValue: (val: VbValue) => {
                  if (this.context.currentScope.has(varName)) {
                    this.context.currentScope.set(varName, val);
                  } else {
                    this.context.setVariable(varName, val);
                  }
                },
              };
            }
            return { value: cloneIfArray(this.evaluate(arg)) };
          });
          return this.context.functionRegistry.callWithRefs(name, argRefs);
        }

        const args = callArgs.map((arg) => this.evaluate(arg));
        return this.context.functionRegistry.call(name, args);
      }

      const variable = this.context.currentScope.get(name);
      if (variable) {
        // A whole statement naming an array is a Sub call on it rather than an index read, which a fresh
        // cscript shows: `arr(9)` written as a statement is error 13 while `x = arr(9)` is error 9 there.
        if (variable.value.type === "Array" && statementCall) {
          throw createVbError(
            VbErrorCodes.TypeMismatch,
            `Type mismatch: '${name}'`,
          );
        }
        // An object named as a whole statement is a Sub call on it, which a fresh cscript answers with 438:
        // `WScript.Arguments` and `WScript.StdOut` alone are both 438, and so is `WScript.Arguments(0)`. A
        // Nothing is the other way round and stays a type mismatch, which `heldObject` is after
        // `Set heldObject = Nothing`.
        if (
          variable.value.type === "Object" &&
          variable.value.value !== null &&
          statementCall
        ) {
          throw createVbError(
            VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
            `Object doesn't support this property or method: '${name}'`,
          );
        }
        if (variable.value.type === "Array") {
          const indices = callArgs.map((arg) =>
            toSubscript(this.evaluate(arg)),
          );
          const arr = variable.value.value as unknown as {
            get: (indices: number[]) => VbValue;
          };
          return arr.get(indices);
        }
        if (variable.value.type === "Object" && variable.value.value !== null) {
          const val = variable.value.value as VbObjectValueData;
          if (val.hasMethod?.("default") && val.getMethod) {
            const method = val.getMethod("default");
            const args = callArgs.map((arg) => this.evaluate(arg));
            return method.func(...args);
          }
        }
      }

      const callee = this.context.getVariable(name);
      if (callee.type === "Object" && callee.value !== null) {
        // The same reading as above for a name that resolves outside the script's own scope, which is where a
        // host object such as WScript lives.
        if (statementCall) {
          throw createVbError(
            VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
            `Object doesn't support this property or method: '${name}'`,
          );
        }
        return this.callObjectMethod(callee, callArgs);
      }
      // A call on a name nothing holds indexes the Empty it stands for, which is a type mismatch, and the same
      // answer comes from every position: measured without Option Explicit, "NoSuchFunction(1)",
      // "NoSuchFunction 1", "Call NoSuchFunction(1)", "x = NoSuchFunction(1)" and "x = NoSuchFunction(1) + 1"
      // are all error 13, while a bare "x = NoSuchFunction" is the Empty itself. With Option Explicit the name
      // is the error instead, "Variable is undefined", which the scope lookup above already answers for every
      // one of those shapes.
      //
      // A bare name that is one of the running class's own members is not that: it is a call of that member,
      // whose answer a statement throws away. Measured on a fresh cscript: inside `Public Sub CallInitialize()`
      // whose body is the bare name `Class_Initialize`, the Initialize event runs again, where the same text
      // at script level is error 438.
      if (bareName && this.context.currentInstance?.hasMember(name) === true) {
        return VbEmpty;
      }
      throw createVbError(
        VbErrorCodes.TypeMismatch,
        `Type mismatch: '${name}'`,
      );
    }

    if (calleeExpr.type === "MemberExpression") {
      const memberResult = this.evaluateMember(calleeExpr);

      // Chained array read, e.g. `obj.Items(0)` or `arr(0)(4)` where the inner
      // `arr(0)` was itself parsed as a MemberExpression (see the Identifier
      // branch above, which already handles the non-chained `arr(4)` case the
      // same way). Missing here until 2026-09-03: this branch only ever tried
      // callObjectMethod and threw for anything else, so a nested array read
      // could never succeed - previously unreachable in practice because
      // nested-array WRITES were separately broken (see parseStatementAssignment
      // above), so `arr(0)` never actually held a populated array to expose this.
      if (memberResult.type === "Array") {
        // The same reading a bare array name gets: as a statement this is a Sub call, so `holder.Items(0)`
        // alone is error 13 where `x = holder.Items(0)` indexes.
        if (statementCall) {
          throw createVbError(
            VbErrorCodes.TypeMismatch,
            `Type mismatch: '${memberExpressionName(calleeExpr)}'`,
          );
        }
        const indices = callArgs.map((arg) => toSubscript(this.evaluate(arg)));
        const arr = memberResult.value as unknown as {
          get: (indices: number[]) => VbValue;
        };
        return arr.get(indices);
      }
      if (memberResult.type === "Object" && memberResult.value !== null) {
        // A member that is no method but holds an object is a Sub call when the whole statement names it,
        // which a fresh cscript answers with 438 for `WScript.Arguments` and for `WScript.Arguments(0)` alike,
        // where the same text inside an expression calls the object's default member. A member that holds a
        // host function is a procedure and is called, which is how `window.setTimeout "code", 100` reaches the
        // host: reading the member answers the same thing either way, so only the call differs.
        const memberValue = memberResult.value as VbObjectValueData;
        if (
          statementCall &&
          !isVbMethodObject(memberValue) &&
          !isVbJsFunctionObject(memberValue)
        ) {
          throw createVbError(
            VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
            `Object doesn't support this property or method: '${memberExpressionName(calleeExpr)}'`,
          );
        }
        return this.callObjectMethod(memberResult, callArgs);
      }
      // A bare statement that names a Property Get reads it and discards the value, which a fresh cscript
      // shows: `heldObject.Prop` is 0 where `heldObject.Value`, a public field, is 5 and
      // `heldObject.Value(1)` is 438.
      if (callArgs.length === 0 && this.lastMemberWasGetter) {
        return memberResult;
      }
      // A call on a member that holds no procedure is error 5 written without arguments and error 438 written
      // with them, both measured on a fresh cscript: `heldObject.Value`, a public field, is 5 while
      // `heldObject.Value(1)` is 438. A member holding an object answers 438 either way.
      if (callArgs.length > 0) {
        throw createVbError(
          VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
          `Object doesn't support this property or method: '${memberExpressionName(calleeExpr)}'`,
        );
      }
      throw invalidProcedureCall(memberExpressionName(calleeExpr));
    }

    const callee = this.evaluate(calleeExpr);
    // Same chained-array-read gap as above, for callees more deeply nested than
    // one level (e.g. a CallExpression callee from `arr(0)(4)(1)`).
    if (callee.type === "Array") {
      const indices = callArgs.map((arg) => toSubscript(this.evaluate(arg)));
      const arr = callee.value as unknown as {
        get: (indices: number[]) => VbValue;
      };
      return arr.get(indices);
    }
    if (callee.type === "Object" && callee.value !== null) {
      return this.callObjectMethod(callee, callArgs);
    }
    // A value that is simply not callable is a type mismatch and not an invalid procedure call, which a
    // fresh cscript shows for every shape: with v = 1, "v 1", "v 1, 2", "v" and "v(1)" are all error 13, and
    // so is "abc"(1) and Empty(1).
    throw createVbError(VbErrorCodes.TypeMismatch, "Type mismatch");
  }

  private callObjectMethod(
    objValue: VbObjectValue,
    callArgs: Expression[],
  ): VbValue {
    const obj = objValue.value as VbObjectValueData | null;
    if (obj === null) {
      throw createVbError(VbErrorCodes.ObjectRequired, "Object required");
    }

    if (isVbMethodObject(obj)) {
      const args = callArgs.map((arg) => this.evaluate(arg));
      return obj.object.getMethod(obj.method).func.call(obj.object, ...args);
    }

    if (isVbJsFunctionObject(obj)) {
      const args = callArgs.map((arg) => {
        const vbVal = this.evaluate(arg);
        return this.vbToJs(vbVal);
      });
      const result = obj.func.call(obj.thisArg, ...args);
      return this.jsToVb(result);
    }

    if (typeof obj.call === "function") {
      const args = callArgs.map((arg) => this.evaluate(arg));
      return obj.call(...args);
    }

    // Default-property invocation: `obj(args)` where obj exposes a `default` (or
    // `item`) method. This is the mechanism behind VBScript's default-member call
    // — `WScript.Arguments(0)`, `Collection(1)`, `dict("key")`, etc. The bare-
    // Identifier form (`variable(0)`) is already handled in evaluateCallInternal's
    // Identifier branch; this covers member expressions like `WScript.Arguments(0)`
    // whose callee evaluates to a VbObjectValueData carrying getMethod/hasMethod.
    if (
      typeof obj.hasMethod === "function" &&
      typeof obj.getMethod === "function"
    ) {
      const defaultName = obj.hasMethod("default")
        ? "default"
        : obj.hasMethod("item")
          ? "item"
          : null;
      if (defaultName) {
        const method = obj.getMethod(defaultName);
        const args = callArgs.map((arg) => this.evaluate(arg));
        return method.func.call(obj, ...args);
      }
    }

    // An object that is no procedure and has no default member is error 438, measured on a fresh cscript:
    // `heldObject` alone, where the class has no default member, is 438, and so is `WScript` and
    // `heldObject(1)`. Calling a plain value with an argument is the other way round and answered above.
    throw createVbError(
      VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
      "Object doesn't support this property or method",
    );
  }

  private vbToJs(value: VbValue): unknown {
    switch (value.type) {
      case "Empty":
        return undefined;
      case "Null":
        return null;
      case "Boolean":
      case "Long":
      case "Double":
      case "Integer":
      case "String":
        return value.value;
      case "LongLong":
        return value.value; // return as bigint
      case "Date":
        return value.value instanceof Date
          ? value.value
          : new Date(value.value);
      case "Array":
        return value.value;
      case "Object": {
        const obj = value.value as VbObjectValueData | null;
        if (
          obj &&
          typeof obj === "object" &&
          (obj as Record<string, unknown>).type === "vbref"
        ) {
          const callFn = (obj as Record<string, unknown>).call as
            ((...args: VbValue[]) => VbValue) | undefined;
          if (typeof callFn === "function") {
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

    if (COMPARISON_OPERATORS.has(node.operator)) {
      this.rejectNumberLiteralBesideText(node.left, node.right, left, right);
    }

    switch (node.operator) {
      case "+":
        return this.add(left, right);
      case "-":
        return this.subtract(left, right);
      case "*":
        return this.multiply(left, right);
      case "/":
        return this.divide(left, right);
      case "\\":
        if (left.type === "LongLong" || right.type === "LongLong") {
          return this.longLongArithmetic(left, right, "\\");
        }
        return this.integerDivide(left, right);
      case "%":
      case "Mod":
        return this.modulo(left, right);
      case "**":
      case "^":
        return this.power(left, right);
      case "==":
      case "=":
        return this.equals(left, right);
      case "!=":
      case "<>":
        return this.notEquals(left, right);
      case "<":
        return this.lessThan(left, right);
      case "<=":
        return this.lessThanOrEqual(left, right);
      case ">":
        return this.greaterThan(left, right);
      case ">=":
        return this.greaterThanOrEqual(left, right);
      case "&": {
        // A Null beside a value contributes nothing: True & Null is "True" under a fresh cscript, and
        // Null & "abc" is "abc". Two of them are Null, which CStr then refuses with error 94.
        if (left.type === "Null" && right.type === "Null") return VbNull;
        if (left.type === "Null")
          return { type: "String", value: toString(right) };
        if (right.type === "Null")
          return { type: "String", value: toString(left) };
        return { type: "String", value: toString(left) + toString(right) };
      }
      case "Is":
        // VBScript compares two object references and nothing else. Anything that is not an object on
        // either side is "Object required", measured for every pair of twelve types under 1033: even
        // "Empty Is Empty" and "2 Is 2" are error 424 rather than False.
        if (left.type !== "Object" || right.type !== "Object") {
          throw createVbError(
            VbErrorCodes.ObjectRequired,
            "Object required: Is",
          );
        }
        return { type: "Boolean", value: left.value === right.value };
      default:
        return VbEmpty;
    }
  }

  private add(left: VbValue, right: VbValue): VbValue {
    if (left.type === "Null" || right.type === "Null") {
      return VbNull;
    }
    if (left.type === "String" && right.type === "String") {
      // Two strings join. A string beside a number is added instead, and the string has to be a number: a
      // fresh cscript answers "1" + 1 with the Double 2 and "a" + 1 with a type mismatch.
      return { type: "String", value: toString(left) + toString(right) };
    }
    if (
      (left.type === "String" && right.type === "Empty") ||
      (left.type === "Empty" && right.type === "String")
    ) {
      // A string beside an Empty joins too, and the Empty contributes nothing: "abc" + Empty is "abc".
      return { type: "String", value: toString(left) + toString(right) };
    }
    if (left.type === "LongLong" || right.type === "LongLong") {
      return this.longLongArithmetic(left, right, "+");
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    // A date survives addition in VBScript, from either side and with a plain number, a Currency included,
    // as the other operand: TypeName(#2026-01-05# + 1) is Date, and so is TypeName(CCur(1.2345) + #2026-01-05#).
    if (left.type === "Date" || right.type === "Date") {
      return { type: "Date", value: serialToDate(leftNum + rightNum) };
    }
    const currency = this.currencyArithmetic(left, right, "+");
    if (currency) return currency;
    return this.numberResult("+", left, right, leftNum + rightNum);
  }

  /**
   * The Currency result of +, - or *, or null when the operator does not keep one.
   *
   * Measured against a fresh cscript: a Currency keeps its type through + and - beside any number, and
   * through * beside a Currency, an Integer, a Long or a Byte. Multiplication by a Single or a Double gives
   * a Double, and so does every division. The arithmetic is done on the scaled integer, so a sum of amounts
   * is exact where a double would round.
   */
  private currencyArithmetic(
    left: VbValue,
    right: VbValue,
    operator: "+" | "-" | "*",
  ): VbValue | null {
    const leftIsCurrency = left.type === "Currency";
    const rightIsCurrency = right.type === "Currency";
    if (!leftIsCurrency && !rightIsCurrency) return null;
    const floatingOther =
      left.type === "Single" ||
      left.type === "Double" ||
      left.type === "String" ||
      right.type === "Single" ||
      right.type === "Double" ||
      right.type === "String";
    if (operator === "*" && floatingOther) return null;
    const leftScaled =
      left.type === "Currency" ? left.value : toCurrencyScaled(toNumber(left));
    const rightScaled =
      right.type === "Currency"
        ? right.value
        : toCurrencyScaled(toNumber(right));
    if (operator === "+")
      return {
        type: "Currency",
        value: checkCurrencyRange(leftScaled + rightScaled),
      };
    if (operator === "-")
      return {
        type: "Currency",
        value: checkCurrencyRange(leftScaled - rightScaled),
      };
    // A product of two scaled amounts is scaled twice, so it comes back down with the half added first,
    // which rounds away from zero the way CCur does.
    const product = leftScaled * rightScaled;
    const half = CURRENCY_SCALE / 2n;
    const rounded =
      product >= 0n
        ? (product + half) / CURRENCY_SCALE
        : (product - half) / CURRENCY_SCALE;
    return { type: "Currency", value: checkCurrencyRange(rounded) };
  }

  /**
   * The type a +, -, * or / result takes, from the two operand types.
   *
   * Measured against a fresh cscript. A Single keeps its type beside a Single, an Integer or a Byte and
   * gives a Double beside a Long, because a Long carries more bits than a Single's mantissa holds. Division
   * gives a Single only in that same company and a Double otherwise. Two Bytes stay a Byte while the result
   * fits one, and everything else is sized by the value itself.
   */
  private numberResult(
    operator: "+" | "-" | "*" | "/",
    left: VbValue,
    right: VbValue,
    value: number,
  ): VbValue {
    const types = [left.type, right.type];
    // A string beside a number always gives a Double, which a fresh cscript shows for "1" + 1 and for
    // "1" + True (the Double 0, since True counts as -1).
    if (types.includes("String")) return { type: "Double", value };
    const hasDouble = types.includes("Double");
    const hasSingle = types.includes("Single");
    const hasLong = types.includes("Long");
    if (operator === "/") {
      // Division answers a Single while both operands are narrow, an Empty counting as a narrow one:
      // CByte(200) / CInt(-3) is a Double while Empty / CByte(200) and CSng(1.5) / True are Singles.
      const narrow = (type: string): boolean =>
        type === "Empty" ||
        type === "Boolean" ||
        type === "Byte" ||
        type === "Integer" ||
        type === "Single";
      const single =
        types.every(narrow) &&
        (types.includes("Empty") || types.includes("Single"));
      return single
        ? { type: "Single", value: toSinglePrecision(value) }
        : { type: "Double", value };
    }
    if (operator === "*") {
      // A product with a Date or a Currency in it is a Double, measured across the matrix: every product
      // with a Date is one, and so is CSng(1.5) * CCur(1.2345).
      if (types.includes("Date") || types.includes("Currency"))
        return { type: "Double", value };
    }
    // A Long keeps its width beside anything whole, which the value alone does not show: CLng(5) - CLng(5)
    // is the Long 0 and not the Integer a value sized answer would give, and so is CLng(70000) * Empty. A
    // product that leaves the Long range is a Double again, which CLng(70000) * CLng(70000) is.
    if (
      hasLong &&
      !hasSingle &&
      !hasDouble &&
      value >= -2147483648 &&
      value <= 2147483647
    ) {
      return { type: "Long", value };
    }
    if (hasDouble) return { type: "Double", value };
    if (hasSingle) {
      if (hasLong) return { type: "Double", value };
      return { type: "Single", value: toSinglePrecision(value) };
    }
    // A Byte stays a Byte while the answer fits one, and an Empty counts as a Byte there: CByte(200) + Empty
    // is the Byte 200 and Empty - CByte(200) is the Integer -200, which does not fit.
    const byteOrEmpty = (type: string): boolean =>
      type === "Byte" || type === "Empty";
    if (
      types.includes("Byte") &&
      types.every(byteOrEmpty) &&
      Number.isInteger(value) &&
      value >= 0 &&
      value <= 255
    ) {
      return { type: "Byte", value };
    }
    return createVbValue(value);
  }

  private subtract(left: VbValue, right: VbValue): VbValue {
    if (left.type === "Null" || right.type === "Null") {
      return VbNull;
    }
    if (left.type === "LongLong" || right.type === "LongLong") {
      return this.longLongArithmetic(left, right, "-");
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    // One date and one number subtracts to a date, a Currency included, but two dates subtract to a plain
    // number, which is how a script gets "how many days apart" out of VBScript.
    const dates =
      (left.type === "Date" ? 1 : 0) + (right.type === "Date" ? 1 : 0);
    if (dates === 1) {
      return { type: "Date", value: serialToDate(leftNum - rightNum) };
    }
    if (dates === 2) {
      return { type: "Double", value: leftNum - rightNum };
    }
    const currency = this.currencyArithmetic(left, right, "-");
    if (currency) return currency;
    return this.numberResult("-", left, right, leftNum - rightNum);
  }

  private multiply(left: VbValue, right: VbValue): VbValue {
    if (left.type === "Null" || right.type === "Null") {
      return VbNull;
    }
    if (left.type === "LongLong" || right.type === "LongLong") {
      return this.longLongArithmetic(left, right, "*");
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    // A product with a Date is a plain number of days, measured: TypeName(#2026-01-05# * 2) is Double.
    if (left.type === "Date" || right.type === "Date") {
      return { type: "Double", value: leftNum * rightNum };
    }
    const currency = this.currencyArithmetic(left, right, "*");
    if (currency) return currency;
    return this.numberResult("*", left, right, leftNum * rightNum);
  }

  private divide(left: VbValue, right: VbValue): VbValue {
    if (left.type === "Null" || right.type === "Null") {
      return VbNull;
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    if (rightNum === 0) {
      // Zero divided by zero is an overflow, and only a dividend that is not zero divided by zero is a
      // division by zero, measured for Empty, Byte, Integer, Double and Boolean operands alike.
      if (leftNum === 0) {
        throw createVbError(VbErrorCodes.Overflow, "Overflow");
      }
      throw createVbError(VbErrorCodes.DivisionByZero, "Division by zero");
    }
    return this.numberResult("/", left, right, leftNum / rightNum);
  }

  /**
   * Integer division, which rounds both operands to a whole number and then cuts toward zero.
   *
   * Measured across the operator matrix under 1033: CSng(1.5) is rounded to 2 first, so CByte(200) \ CSng(1.5)
   * is 100 and not the 133 a plain division gives, and the cut is toward zero rather than down, so
   * CByte(200) \ CInt(-3) is -66 and not -67. A Null answers Null, a text that is no number is a type
   * mismatch, and a zero divisor is error 11.
   */
  private integerDivide(left: VbValue, right: VbValue): VbValue {
    if (left.type === "Null" || right.type === "Null") return VbNull;
    const leftOperand = readBitwiseOperand(left, "Integer");
    const rightOperand = readBitwiseOperand(right, "Integer");
    if (rightOperand.value === 0) {
      throw createVbError(VbErrorCodes.DivisionByZero, "Division by zero");
    }
    const value = Math.trunc(leftOperand.value / rightOperand.value);
    return bitwiseValue(
      integerResultKind(leftOperand.kind, rightOperand.kind),
      value,
    );
  }

  /** The remainder, with the same rounding, the same types and the same refusals as integer division. */
  private modulo(left: VbValue, right: VbValue): VbValue {
    if (left.type === "Null" || right.type === "Null") return VbNull;
    const leftOperand = readBitwiseOperand(left, "Integer");
    const rightOperand = readBitwiseOperand(right, "Integer");
    if (rightOperand.value === 0) {
      throw createVbError(VbErrorCodes.DivisionByZero, "Division by zero");
    }
    const value = leftOperand.value % rightOperand.value;
    return bitwiseValue(
      integerResultKind(leftOperand.kind, rightOperand.kind),
      value,
    );
  }

  private power(left: VbValue, right: VbValue): VbValue {
    if (left.type === "Null" || right.type === "Null") {
      return VbNull;
    }
    const leftNum = toNumber(left);
    const rightNum = toNumber(right);
    const value = Math.pow(leftNum, rightNum);
    // A power that leaves the Double range is an error and not an infinity, measured across the operator
    // matrix: -3 ^ 1.5 and 0 ^ -1 are error 5, since no such value exists, while -3 ^ 46027 and 200 ^ 200
    // are error 6, since the value exists but is out of range.
    if (Number.isNaN(value)) {
      throw createVbError(
        VbErrorCodes.InvalidProcedureCall,
        "Invalid procedure call or argument",
      );
    }
    if (!Number.isFinite(value)) {
      if (leftNum === 0 && rightNum < 0) {
        throw createVbError(
          VbErrorCodes.InvalidProcedureCall,
          "Invalid procedure call or argument",
        );
      }
      throw createVbError(VbErrorCodes.Overflow, "Overflow");
    }
    return { type: "Double", value };
  }

  /**
   * Whether a comparison raises a type mismatch because a number is written out beside a text.
   *
   * Measured on a fresh cscript, and it turns on how the number was written rather than on its type:
   * `"abc" = 1` and `"abc" = (1)` are error 13, and so are `"abc" < 1`, `1 = "abc"` and `Select Case "abc"`
   * with `Case 1`, while a number that came from anywhere else is compared as text instead and simply
   * answers False: `"abc" = 1 * 1`, `"abc" = 1 + 0`, `"abc" = CInt(1)`, `"abc" = Len("a")` and
   * `"abc" = heldNumber` (with heldNumber = 1) are all False. A text that is a number is compared as a
   * number either way, which is why `"1" = 1` is True.
   */
  checkNumberLiteralBesideText(
    leftNode: Expression,
    rightNode: Expression,
    left: VbValue,
    right: VbValue,
  ): void {
    this.rejectNumberLiteralBesideText(leftNode, rightNode, left, right);
  }

  private rejectNumberLiteralBesideText(
    leftNode: Expression,
    rightNode: Expression,
    left: VbValue,
    right: VbValue,
  ): void {
    const pairs: ReadonlyArray<readonly [Expression, VbValue]> = [
      [leftNode, right],
      [rightNode, left],
    ];
    for (const [node, other] of pairs) {
      if (!isNumberLiteralNode(node) || other.type !== "String") continue;
      if (parseVbNumericString(other.value.trim()) === null) {
        throw createVbError(
          VbErrorCodes.TypeMismatch,
          `Type mismatch: ${other.value}`,
        );
      }
    }
  }

  /**
   * The order of two values as every comparison operator reads them, or null when the answer is Null.
   *
   * Select Case shares this reading rather than comparing its own way, which is what a fresh cscript shows:
   * `Select Case "b"` matches `Case "b"` and not `Case "B"`, and `Select Case "abc"` with `Case 1` is error
   * 13, exactly as the = operator answers for the same pair.
   */
  compareValues(left: VbValue, right: VbValue): number | null {
    return this.compareOperands(left, right);
  }

  /**
   * The two sides of a comparison as scaled Currency amounts, or null when no Currency is involved.
   *
   * Measured against a fresh cscript: a comparison with a Currency on one side brings the other side to the
   * Currency's own scale first. CCur(0.12345) = 0.12345 is True because the Double becomes 0.1235, and
   * CCur(1) = 1.00001 is True because 1.00001 becomes 1.0000. The comparison itself is then done on the
   * scaled integers, so two amounts that differ in the fourth decimal are still told apart:
   * CCur(123456789012.3456) = CCur(123456789012.3457) is False, where comparing two doubles would say True.
   */
  private currencyComparison(
    left: VbValue,
    right: VbValue,
  ): readonly [bigint, bigint] | null {
    if (left.type === "Currency" && right.type === "Currency")
      return [left.value, right.value];
    const currencySide =
      left.type === "Currency"
        ? left
        : right.type === "Currency"
          ? right
          : null;
    if (currencySide === null) return null;
    const otherSide = currencySide === left ? right : left;
    if (!isPlainNumber(otherSide)) return null;
    const scaled = toCurrencyScaled(toNumber(otherSide));
    return left.type === "Currency"
      ? [currencySide.value, scaled]
      : [scaled, currencySide.value];
  }

  // Every comparison operator reads its two sides the same way, which a fresh cscript shows over the pairs
  // below. The rules, each measured:
  //
  //   * An Array on either side is a type mismatch, and it wins over Null propagation.
  //   * A Null on either side answers Null, whatever the other side is: `"x" = Null`, `1 = Null`,
  //     `"x" < Null` and `Null <> Null` are all Null.
  //   * An Object on either side is error 91 when it is Nothing and error 438 otherwise, since the
  //     comparison asks it for a value it does not have. Two objects are not compared at all, which is
  //     what the Is operator is for: `someObject = someObject` is 438.
  //   * A text beside a number is read as a number, so "1" = 1, "1.5" = 1.5, " 1 " = 1 and "2" < 10 are
  //     all True, while a text that is no number is a type mismatch: "abc" = 1 and "abc" < 1 are error 13.
  //   * A text beside anything else is compared as text, character by character and case sensitively:
  //     "b" = "B" is False and "B" < "b" is True. This is why "True" = True is True and "1" = True is
  //     False, why "" = Empty is True, and why "" = "x" and "x" < Empty are False.
  //   * Anything else is compared as a number, with a Currency on one side bringing the other side to the
  //     Currency's own scale first (see currencyComparison).
  private compareOperands(left: VbValue, right: VbValue): number | null {
    if (left.type === "Array" || right.type === "Array") {
      throw createVbError(VbErrorCodes.TypeMismatch, "Type mismatch");
    }
    if (left.type === "Null" || right.type === "Null") return null;
    if (left.type === "Object" || right.type === "Object") {
      const isNothing = [left, right].some(
        (side) => side.type === "Object" && side.value === null,
      );
      if (isNothing) {
        throw createVbError(
          VbErrorCodes.ObjectVariableNotSet,
          "Object variable not set",
        );
      }
      throw createVbError(
        VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
        "Object doesn't support this property or method",
      );
    }

    if (left.type === "String" || right.type === "String") {
      const textSide = left.type === "String" ? left : right;
      const otherSide = textSide === left ? right : left;
      // A text that is a number is compared as a number, so "1" = 1, "1.5" = 1.5 and " 1 " = 1 are all
      // True. A text beside a number that is no number at all is compared as text, which answers False
      // unless the number was written out as a literal, in which case the caller has already raised error
      // 13: see rejectNumberLiteralBesideText.
      const parsed = isNumberValue(otherSide)
        ? parseVbNumericString(toString(textSide).trim())
        : null;
      if (parsed !== null) {
        const order = compareNumbers(parsed, toNumber(otherSide));
        return textSide === left ? order : -order;
      }
      const text = toString(textSide);
      const otherText = toString(otherSide);
      const order = text < otherText ? -1 : text > otherText ? 1 : 0;
      return textSide === left ? order : -order;
    }

    const currencyPair = this.currencyComparison(left, right);
    if (currencyPair !== null) {
      if (currencyPair[0] === currencyPair[1]) return 0;
      return currencyPair[0] < currencyPair[1] ? -1 : 1;
    }
    return compareNumbers(toNumber(left), toNumber(right));
  }

  private equals(left: VbValue, right: VbValue): VbValue {
    const order = this.compareOperands(left, right);
    if (order === null) return VbNull;
    return { type: "Boolean", value: order === 0 };
  }

  private notEquals(left: VbValue, right: VbValue): VbValue {
    const order = this.compareOperands(left, right);
    if (order === null) return VbNull;
    return { type: "Boolean", value: order !== 0 };
  }

  private lessThan(left: VbValue, right: VbValue): VbValue {
    const order = this.compareOperands(left, right);
    if (order === null) return VbNull;
    return { type: "Boolean", value: order < 0 };
  }

  private lessThanOrEqual(left: VbValue, right: VbValue): VbValue {
    const order = this.compareOperands(left, right);
    if (order === null) return VbNull;
    return { type: "Boolean", value: order <= 0 };
  }

  private greaterThan(left: VbValue, right: VbValue): VbValue {
    const order = this.compareOperands(left, right);
    if (order === null) return VbNull;
    return { type: "Boolean", value: order > 0 };
  }

  private greaterThanOrEqual(left: VbValue, right: VbValue): VbValue {
    const order = this.compareOperands(left, right);
    if (order === null) return VbNull;
    return { type: "Boolean", value: order >= 0 };
  }

  private evaluateUnary(node: UnaryExpression): VbValue {
    const argument = this.evaluate(node.argument);

    switch (node.operator) {
      case "-":
        // Null propagates through unary minus too (`-Null` is Null, not a crash), which a fresh cscript shows.
        if (argument.type === "Null") return VbNull;
        // The operand's own type survives where VBScript keeps it: a Single, a Currency and a Date keep
        // theirs, and a numeric String becomes a Double ("-"2"" is the Double -2, not an Integer).
        if (argument.type === "Single")
          return { type: "Single", value: -argument.value };
        if (argument.type === "Double")
          return { type: "Double", value: -argument.value };
        if (argument.type === "Currency")
          return { type: "Currency", value: -argument.value };
        if (argument.type === "LongLong")
          return { type: "LongLong", value: -argument.value };
        if (argument.type === "Date") {
          return {
            type: "Date",
            value: serialToDate(-dateToSerial(argument.value)),
          };
        }
        if (argument.type === "String")
          return { type: "Double", value: -toNumber(argument) };
        return createVbValue(-toNumber(argument));
      case "+":
        if (argument.type === "Null") return VbNull;
        return createVbValue(toNumber(argument));
      case "!":
      case "Not": {
        // Null propagates through Not too (`Not Null` is Null), and a text that is no number is a type
        // mismatch, as it is for the five binary bitwise operators.
        if (argument.type === "Null") return VbNull;
        // Real VBScript's Not is VarNot, a bitwise complement that keeps the operand's own type: Not True is
        // the Boolean False, Not CByte(200) is the Byte 55, Not CInt(-3) is the Integer 2, and Not Empty is
        // the Long -1. Confirmed across the operator matrix under 1033.
        const operand = readBitwiseOperand(argument);
        return bitwiseValue(operand.kind, ~operand.value);
      }
      default:
        return VbEmpty;
    }
  }

  // Real VBScript's And/Or/Xor/Eqv/Imp are three-valued (True/False/Null),
  // not JS's two-valued short-circuit logic - `Null` propagates through all
  // five, EXCEPT And/Or's documented "absorbing value" special case (`False
  // And Null = False`, `True Or Null = True`, since the result is already
  // determined regardless of what the Null side would have been). This also
  // means And/Or can no longer short-circuit: `Null Or True` must actually
  // see the True on the right before it can resolve to True rather than
  // Null. Found via Wine's own vbscript.dll conformance suite, dlls/
  // vbscript/tests/lang.vbs - `Null Or True` previously crashed outright
  // (toBoolean(Null) throws) since the old short-circuit form evaluated
  // toBoolean(left) before ever looking at the right operand.
  private evaluateLogical(node: LogicalExpression): VbValue {
    const left = this.evaluate(node.left);
    const right = this.evaluate(node.right);
    switch (node.operator) {
      case "&&":
      case "And":
        return this.applyBitwise("And", left, right);
      case "||":
      case "Or":
        return this.applyBitwise("Or", left, right);
      case "Xor":
        return this.applyBitwise("Xor", left, right);
      case "Eqv":
        return this.applyBitwise("Eqv", left, right);
      case "Imp":
        return this.applyBitwise("Imp", left, right);
      default:
        return VbEmpty;
    }
  }

  /**
   * One of the five bitwise operators, which are three-valued where a Null is involved.
   *
   * Both operands are read before anything is decided, so a text that is no number is a type mismatch even
   * beside a Null: a fresh cscript answers `"abc" And Null` with error 13 and not 94.
   */
  private applyBitwise(
    operator: "And" | "Or" | "Xor" | "Eqv" | "Imp",
    left: VbValue,
    right: VbValue,
  ): VbValue {
    const leftIsNull = left.type === "Null";
    const rightIsNull = right.type === "Null";
    if (leftIsNull && rightIsNull) return VbNull;
    if (leftIsNull || rightIsNull) {
      const operand = readBitwiseOperand(leftIsNull ? right : left);
      return bitwiseWithNull(operator, operand, leftIsNull);
    }
    const leftOperand = readBitwiseOperand(left);
    const rightOperand = readBitwiseOperand(right);
    const first = leftOperand.value;
    const second = rightOperand.value;
    const value =
      operator === "And"
        ? first & second
        : operator === "Or"
          ? first | second
          : operator === "Xor"
            ? first ^ second
            : operator === "Eqv"
              ? ~(first ^ second)
              : ~first | second;
    return bitwiseValue(
      bitwiseKind(leftOperand.kind, rightOperand.kind),
      value,
    );
  }

  private evaluateAssignment(node: AssignmentExpression): VbValue {
    const value = this.evaluate(node.right);
    const isSet = node.isSet ?? false;
    // A value an expression made carries a reference with it, which the place it is stored in takes; a value
    // read out of another place carries none, so one is taken for the new place here.
    const made = isMadeValue(value);
    if (!made) {
      retainValue(value);
    }

    try {
      return this.storeAssignmentValue(node, value, isSet);
    } finally {
      // An instance whose last reference goes runs Class_Terminate right here, which is why `Set o = New Thing`
      // runs the previous instance's Terminate between the new instance's initialize and this point.
      if (made) {
        releaseMadeValue(value);
      } else {
        releaseValue(value);
      }
    }
  }

  private storeAssignmentValue(
    node: AssignmentExpression,
    value: VbValue,
    isSet: boolean,
  ): VbValue {
    if (node.left.type === "Identifier") {
      if (isSet) {
        // Set takes an object and nothing else, measured on a fresh cscript: Set o = 5, Set o = "abc",
        // Set o = Empty, Set o = Null and even Set o = Array(1, 2) are all error 424, while
        // Set o = Nothing is the one way to empty a variable that holds an object.
        if (value.type !== "Object") {
          throw createVbError(VbErrorCodes.ObjectRequired, "Object required");
        }
      } else if (value.type === "Object") {
        // An object needs Set, and assigning one without it is error 438 for an object of a class, measured
        // on a fresh cscript for every shape it can be written in: `x = New Holder`, `x = Maker()`,
        // `x = held` and `x = items(0)` are all 438, while `x = Nothing` is 91, the object variable that is
        // not set. Passing an object as an argument and reading one are both fine.
        if (value.value === null) {
          throw createVbError(
            VbErrorCodes.ObjectVariableNotSet,
            "Object variable not set",
          );
        }
        throw createVbError(
          VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
          "Object doesn't support this property or method",
        );
      }
      // A variable Dim wrote with parentheses is an array for the rest of the run, so a whole value cannot go
      // into it: a fresh cscript answers `Dim a(1)` with `a = 5` and with `a = Array(1, 2)` alike with error 13,
      // `Dim a()` with `a = 5` with error 13, and `Dim a()` with `a = Array(1, 2)` with error 9, which is the
      // subscript the array it has no size for cannot take. Only a Dim marks a name this way, so a plain
      // `Dim x` takes an array whatever ReDim did to it first. An object without Set is the reading that comes
      // first: measured, `slots = Array(1, 2)` on an array variable is 438, the error that write already is.
      const declared = this.context.currentScope.get(node.left.name);
      if (declared?.isArray && !isSet) {
        if (value.type === "Object") {
          throw createVbError(
            VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
            "Object doesn't support this property or method",
          );
        }
        const held = declared.value;
        const heldArray =
          held.type === "Array" ? (held.value as VbArray) : null;
        if (
          value.type === "Array" &&
          heldArray !== null &&
          !heldArray.isSized()
        ) {
          throw createVbError(
            VbErrorCodes.SubscriptOutOfRange,
            "Subscript out of range",
          );
        }
        throw createVbError(VbErrorCodes.TypeMismatch, "Type mismatch");
      }
      // An array is copied where it is assigned, which a fresh cscript shows: after b = a, changing b(0) or
      // erasing b leaves a as it was. The copy holds what the original held, so the elements are kept too.
      const stored = cloneIfArray(value);
      this.releaseStoredValue(node.left.name);
      this.context.setVariable(node.left.name, stored);
      retainValue(stored);
      return value;
    }

    if (node.left.type === "MemberExpression") {
      this.assignToMember(node.left, value, isSet);
      return value;
    }

    throw new Error(`Invalid assignment target: ${node.left.type}`);
  }

  /**
   * Lets go of what a name held before a new value goes into it, which runs that value's Terminate event when
   * it was the last reference.
   *
   * Measured: `Sub ByRefReplace(item)` written with `Set item = New Probe` runs the first instance's Terminate
   * event at that statement, since that name stopped holding it. What it was handed was not a reference of its
   * own, which is why leaving the procedure does not run it a second time.
   */
  private releaseStoredValue(name: string): void {
    // A class property of the running instance is not one of the scope's own variables, and the scope may hold
    // a global of the same name that the property read takes priority over.
    if (this.context.currentInstance?.hasProperty(name)) return;
    const variable = this.context.currentScope.get(name);
    if (!variable) return;
    releaseValue(variable.value);
  }

  private assignToMember(
    node: MemberExpression,
    value: VbValue,
    isSet: boolean,
  ): void {
    let object = this.evaluate(node.object as Expression);

    if (node.object.type === "VbWithObject") {
      object = this.evaluateWithObject(node.object as VbWithObjectExpression);
    }

    // An array element is read before the member name is, since the name of a computed one is written as a
    // plain expression there: `a(Null)` is error 13 from the subscript and not the 94 that reading the Null as
    // a name would raise.
    if (object.type === "Array") {
      const arr = object.value as unknown as {
        set: (indices: number[], v: VbValue) => void;
      };
      const indices = (node.vbIndices ?? [node.property as Expression]).map(
        (subscript) => toSubscript(this.evaluate(subscript)),
      );
      // An element that is written over lets go of what it held, which the array does itself: measured,
      // `Set slots(0) = New Probe` twice runs term before the second init, and `Set slots(0) = Nothing` too.
      arr.set(indices, value);
      retainValue(value);
      return;
    }

    let propertyName: string;
    if (node.computed) {
      const propValue = this.evaluate(node.property as Expression);
      propertyName = toString(propValue);
    } else {
      propertyName = (node.property as Identifier).name;
    }

    if (object.type === "Object") {
      const obj = object.value as VbObjectValueData | null;
      if (
        obj === null ||
        (typeof obj !== "object" && typeof obj !== "function")
      ) {
        throw createVbError(VbErrorCodes.ObjectRequired, "Object required");
      }

      if (
        (Object.prototype.hasOwnProperty.call(obj, "setProperty") ||
          Object.prototype.hasOwnProperty.call(obj, "classInfo")) &&
        typeof obj.setProperty === "function"
      ) {
        // Writing a member a class does not have is 438, and so is writing a Private one from outside:
        // measured, `first.hiddenCount = 5` from script level is 438 and leaves the class's own value alone,
        // where `first.NoSuchMember = 1` is 438 too.
        if (isVbClassInstance(obj)) {
          const instance = obj as unknown as VbObjectInstance;
          if (
            !instance.isMemberVisible(
              propertyName,
              this.context.currentInstance,
            )
          ) {
            throw createVbError(
              VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
              `Object doesn't support this property or method: '${propertyName}'`,
            );
          }
          if (
            !instance.hasMethod(propertyName) &&
            !instance.hasMember(propertyName)
          ) {
            throw createVbError(
              VbErrorCodes.ObjectDoesntSupportPropertyOrMethod,
              `Object doesn't support this property or method: '${propertyName}'`,
            );
          }
          // A plain field that is written over lets go of what it held, while a Property Set is code of its
          // own: measured, `Set probeOne.Linked = New Probe` runs the new instance's init and its Terminate
          // when the instance holding it goes, and `Set obj.Field = x` twice runs the first value's Terminate.
          if (
            instance.hasStoredField(propertyName) &&
            !instance.hasGetter(propertyName)
          ) {
            releaseValue(instance.getStoredField(propertyName));
            obj.setProperty(propertyName, value, isSet);
            retainValue(value);
            return;
          }
        }
        obj.setProperty(propertyName, value, isSet);
      } else {
        const jsValue = this.vbToJs(value);
        obj[propertyName] = jsValue;
      }
    } else {
      throw createVbError(VbErrorCodes.ObjectRequired, "Object required");
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
    if (v.type === "LongLong") return v.value;
    if (v.type === "Integer" || v.type === "Long" || v.type === "Byte")
      return BigInt(v.value);
    if (v.type === "Boolean") return v.value ? BigInt(-1) : BigInt(0);
    if (v.type === "Empty") return BigInt(0);
    // Float types demote to Double for mixed arithmetic
    return BigInt(Math.trunc(toNumber(v)));
  }

  private longLongArithmetic(
    left: VbValue,
    right: VbValue,
    op: "+" | "-" | "*" | "\\",
  ): VbValue {
    // If either side is floating-point, promote to Double
    const floatTypes = ["Single", "Double", "Currency"];
    if (floatTypes.includes(left.type) || floatTypes.includes(right.type)) {
      const l = toNumber(left);
      const r = toNumber(right);
      if (op === "+") return { type: "Double", value: l + r };
      if (op === "-") return { type: "Double", value: l - r };
      if (op === "*") return { type: "Double", value: l * r };
      if (r === 0)
        throw createVbError(VbErrorCodes.DivisionByZero, "Division by zero");
      return { type: "Long", value: Math.floor(l / r) };
    }
    const l = this.toLongLongBigInt(left);
    const r = this.toLongLongBigInt(right);
    if (op === "+")
      return { type: "LongLong", value: checkLongLongRange(l + r) };
    if (op === "-")
      return { type: "LongLong", value: checkLongLongRange(l - r) };
    if (op === "*")
      return { type: "LongLong", value: checkLongLongRange(l * r) };
    if (r === BigInt(0))
      throw createVbError(VbErrorCodes.DivisionByZero, "Division by zero");
    return { type: "LongLong", value: l / r };
  }
}
