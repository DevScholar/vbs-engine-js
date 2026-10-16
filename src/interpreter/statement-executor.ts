import type {
  Statement,
  BlockStatement,
  ExpressionStatement,
  IfStatement,
  VbDimStatement,
  VbReDimStatement,
  VbEraseStatement,
  VbConstStatement,
  VbForToStatement,
  ForOfStatement,
  VbDoLoopStatement,
  WhileStatement,
  VbSelectCaseStatement,
  WithStatement,
  VbExitStatement,
  VbOptionExplicitStatement,
  VbSubStatement,
  VbFunctionStatement,
  VbClassStatement,
  VbOnErrorHandlerStatement,
  TSEnumDeclaration,
  VbTypeStatement,
  Expression,
  VbGotoStatement,
} from '../ast/index.ts';
import type { VbValue } from '../runtime/index.ts';
import {
  VbContext,
  VbEmpty,
  createVbValue,
  toBoolean,
  toNumber,
  toSubscript,
  toWholeNumber,
  VbError,
  VbErrorCodes,
  createVbError,
  createVbErrorFromHost,
  VbArray,
  createVbArray,
  VbObjectInstance,
  VbClass,
  retainValue,
  releaseValue,
  markMadeValue,
  type VbProperty,
  type VbParameterInfo,
  type VbArgumentRef,
  declaredArity,
  checkArgumentRange,
} from '../runtime/index.ts';
import { ExpressionEvaluator } from './expression-evaluator.ts';

// Real VBScript's control-flow conditions (If/While/Do/Select Case) treat
// Null as simply False - unlike toBoolean() itself, which correctly throws
// "Invalid use of Null" for Null in a general boolean-coercion context
// (comparisons, And/Or, CBool). `If Null Then`/`While Null` silently don't
// execute their body rather than crashing the whole script. Found via
// Wine's own vbscript.dll conformance suite, dlls/vbscript/tests/lang.vbs:
// `if null then call ok(false, ...)` / `while null : call ok(false, ...) :
// wend` both assert the body never runs, not that an error is raised.
function toConditionBoolean(value: VbValue): boolean {
  if (value.type === 'Null') return false;
  return toBoolean(value);
}

/**
 * An array handed to a ByVal parameter is a copy, so what the procedure does to it stays inside.
 *
 * Measured on a fresh cscript: a Sub with `ByVal values` that writes values(0) leaves the caller's array as it
 * was, where the same Sub written with the default ByRef changes it.
 */
function copyArrayArgument(value: VbValue): VbValue {
  if (value.type !== 'Array') return value;
  return { type: 'Array', value: (value.value as VbArray).clone() };
}

/** The parameter list of a class member as the argument counts are read from it. */
function classParameterInfo(params: import('../ast/index.ts').VbParameter[]): VbParameterInfo[] {
  return params.map(param => ({
    name: param.name.name,
    byRef: param.byRef,
    isArray: param.isArray,
    isOptional: param.isOptional,
    isParamArray: param.isParamArray,
  }));
}

export class ControlFlowSignal extends Error {
  constructor(
    public type: 'exit' | 'return' | 'continue' | 'break' | 'goto',
    public value?: VbValue
  ) {
    super(type);
  }

  /**
   * Set when the signal carries the answer of the statement that left the procedure, which is what a Function
   * answers with after the text a script handed to `Execute` left it: measured, the function answers the value
   * it had reached rather than one read out of its scope.
   */
  public carriesThisStatement = false;
}

export class GotoSignal extends ControlFlowSignal {
  constructor(public labelName: string) {
    super('goto');
  }
}

function tsTypeAnnotationName(annotation: import('../ast/index.ts').TSTypeAnnotation): string {
  const t = annotation.typeAnnotation;
  switch (t.type) {
    case 'TSStringKeyword':
      return 'string';
    case 'TSBooleanKeyword':
      return 'boolean';
    case 'TSObjectKeyword':
      return 'object';
    case 'TSAnyKeyword':
      return 'variant';
    case 'TSVoidKeyword':
      return 'variant';
    case 'TSNumberKeyword':
      return 'double';
    case 'TSNullKeyword':
      return 'variant';
    case 'TSTypeReference':
      return t.typeName.name;
    case 'TSArrayType':
      return tsTypeAnnotationName({ type: 'TSTypeAnnotation', typeAnnotation: t.elementType });
    default:
      return 'variant';
  }
}

function getTypedDefault(typeName: string, context?: VbContext): VbValue {
  switch (typeName.toLowerCase()) {
    case 'integer':
      return { type: 'Integer', value: 0 };
    case 'long':
      return { type: 'Long', value: 0 };
    case 'longlong':
      return { type: 'LongLong', value: BigInt(0) };
    case 'single':
      return { type: 'Single', value: 0 };
    case 'double':
      return { type: 'Double', value: 0 };
    case 'currency':
      return { type: 'Currency', value: BigInt(0) };
    case 'byte':
      return { type: 'Byte', value: 0 };
    case 'boolean':
      return { type: 'Boolean', value: false };
    case 'string':
      return { type: 'String', value: '' };
    case 'date':
      return { type: 'Date', value: new Date(0) };
    case 'variant':
      return VbEmpty;
    default:
      if (context && context.classRegistry.has(typeName)) {
        const instance = context.classRegistry.createInstance(typeName, []);
        return { type: 'Object', value: instance };
      }
      return { type: 'Object', value: null }; // Object/class → Nothing
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
        case 'VbEraseStatement':
          return this.executeEraseStatement(node);
        case 'VbConstStatement':
          return this.executeConstStatement(node);
        case 'VbForToStatement':
          return this.executeForToStatement(node);
        case 'ForOfStatement':
          return this.executeForEachStatement(node);
        case 'VbDoLoopStatement':
          return this.executeDoLoopStatement(node);
        case 'WhileStatement':
          return this.executeWhileStatement(node);
        case 'VbSelectCaseStatement':
          return this.executeSelectCaseStatement(node);
        case 'WithStatement':
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
        case 'VbGotoStatement':
          return this.executeGotoStatement(node);
        case 'VbLabelStatement':
        case 'EmptyStatement':
          return VbEmpty;
        case 'TSEnumDeclaration':
          return this.executeEnumStatement(node);
        case 'VbTypeStatement':
          return this.executeTypeStatement(node);
        case 'ReturnStatement':
          throw new ControlFlowSignal('return');
        default:
          throw new Error(`Unknown statement type: ${(node as Statement).type}`);
      }
    } catch (error) {
      if (error instanceof ControlFlowSignal) {
        throw error;
      }
      // Stamp the failing statement's position on the way out, but only if nothing has yet.
      // Every enclosing statement catches the same error again, so the first frame to see it --
      // the innermost one -- wins, and that is where the real engine points: at the statement
      // that failed, not at the block containing it.
      const at = (node as Statement).loc?.start;
      const stampPosition = (target: { line?: number; column?: number }): void => {
        if (at && target.line === undefined) {
          target.line = at.line;
          target.column = at.column;
        }
      };

      if (error instanceof VbError) {
        stampPosition(error);
        this.context.setError(error);
        if (this.context.onErrorResumeNext) {
          return VbEmpty;
        }
        throw error;
      }
      const vbError = createVbErrorFromHost(error);
      stampPosition(vbError);
      this.context.setError(vbError);
      if (this.context.onErrorResumeNext) {
        return VbEmpty;
      }
      throw vbError;
    }
  }

  private executeExpressionStatement(node: ExpressionStatement): VbValue {
    const result = this.exprEvaluator.evaluate(node.expression);
    // A whole statement throws away the answer of a call it writes, so an instance only that answer held is
    // let go of here: a fresh cscript runs the Terminate event of the instance `FreshMaker()` alone answers
    // with at that statement. A bare name is left alone, since what it names lives somewhere that keeps its
    // own reference and letting go of this statement's one is not the same thing.
    const expression = node.expression as { type: string; callee?: { type?: string } };
    const isCall =
      expression.type === 'CallExpression' && expression.callee?.type !== 'VbWithObject';
    if (isCall) {
      releaseValue(result);
    }
    return result;
  }

  private executeBlockStatement(node: BlockStatement): VbValue {
    let result: VbValue = VbEmpty;
    for (const stmt of node.body) {
      result = this.execute(stmt);
      this.leaveProcedureIfExecuteLeftIt(result);
    }
    return result;
  }

  /**
   * Leaves the procedure when the text a script handed to `Execute` or `ExecuteGlobal` raised the 1039 an Exit
   * in it raises, which is what Windows does: measured, a Sub whose `Execute "Exit Sub"` answered 1039 does not
   * run the rest of its own statements, and the value a Function had reached is the answer it gives.
   *
   * The mark is the interpreter's, so a `Run` that stands in for this engine's own `Execute` call is what sets
   * it, and it is cleared here before the signal goes out, since the caller of this procedure has nothing to do
   * with it. The signal carries the answer as it stands, which is what the procedure gives back.
   */
  private leaveProcedureIfExecuteLeftIt(answer: VbValue): void {
    if (!this.context.leftProcedureThroughExecute) return;
    this.context.leftProcedureThroughExecute = false;
    const target = this.context.getCurrentCallTarget();
    if (target === null) return;
    this.context.setExitFlag(target);
    const signal = new ControlFlowSignal('exit', answer);
    signal.carriesThisStatement = true;
    throw signal;
  }

  private executeIfStatement(node: IfStatement): VbValue {
    const test = this.exprEvaluator.evaluate(node.test);

    if (toConditionBoolean(test)) {
      return this.execute(node.consequent);
    } else if (node.alternate) {
      return this.execute(node.alternate);
    }

    return VbEmpty;
  }

  private executeDimStatement(node: VbDimStatement): VbValue {
    for (const decl of node.declarations) {
      let value: VbValue = VbEmpty;

      if (decl.isArray) {
        // Dim a() makes a dynamic array with no size yet: IsArray answers True, UBound raises error 9, and a
        // later ReDim gives it its bounds. Measured on a fresh cscript, as is VarType 8204 for it. A bound is
        // a whole number, so a script that writes a fraction gets the nearest one.
        const bounds = (decl.arrayBounds ?? []).map(b => {
          const boundValue = this.exprEvaluator.evaluate(b);
          return toWholeNumber(toNumber(boundValue));
        });
        const arr = createVbArray(bounds);
        value = { type: 'Array', value: arr };
      } else if (decl.init) {
        value = this.exprEvaluator.evaluate(decl.init);
      } else if (decl.typeAnnotation) {
        value = getTypedDefault(tsTypeAnnotationName(decl.typeAnnotation), this.context);
      }

      // A variable Dim wrote with parentheses is an array for the rest of the run, and giving it a whole value
      // afterwards is an error: a fresh cscript answers `Dim a(1)` with `a = 5` and with `a = Array(1, 2)`
      // alike with error 13, and `Dim a()` with `a = Array(1, 2)` with error 9.
      this.context.declareVariable(decl.id.name, value, { isArray: decl.isArray });
    }

    return VbEmpty;
  }

  private executeEnumStatement(node: TSEnumDeclaration): VbValue {
    let nextValue = BigInt(0);
    for (const member of node.members) {
      let memberValue: VbValue;
      if (member.initializer !== null) {
        const evaluated = this.exprEvaluator.evaluate(member.initializer);
        const numVal = toNumber(evaluated);
        nextValue = BigInt(Math.trunc(numVal));
        memberValue = { type: 'Long', value: Number(nextValue) };
      } else {
        memberValue = { type: 'Long', value: Number(nextValue) };
      }
      this.context.globalScope.declare(member.id.name, memberValue, { isConst: true });
      nextValue++;
    }
    return VbEmpty;
  }

  private executeTypeStatement(node: VbTypeStatement): VbValue {
    const cls = new VbClass(node.name.name);

    cls.initializer = instance => {
      for (const member of node.members) {
        const defaultValue = member.typeAnnotation
          ? getTypedDefault(tsTypeAnnotationName(member.typeAnnotation), this.context)
          : VbEmpty;
        instance.setProperty(member.id.name, defaultValue);
      }
    };

    this.context.classRegistry.register(cls);
    return VbEmpty;
  }

  private executeReDimStatement(node: VbReDimStatement): VbValue {
    for (const decl of node.declarations) {
      if (!decl.isArray) continue;

      const bounds = (decl.arrayBounds ?? []).map(b => {
        const boundValue = this.exprEvaluator.evaluate(b);
        return toWholeNumber(toNumber(boundValue));
      });
      const existingVar = this.context.currentScope.get(decl.id.name);

      if (existingVar && existingVar.value.type === 'Array') {
        const arr = existingVar.value.value as VbArray;
        // A fixed array cannot be sized again, which a fresh cscript answers with error 10: Dim a(2), Dim a(2)
        // then Erase a, and Dim a(2) then ReDim Preserve a(3) are all error 10.
        if (!arr.isDynamic()) {
          throw createVbError(10, 'This array is fixed or temporarily locked');
        }
        arr.redim(bounds, node.preserve);
      } else {
        // A ReDim always makes a dynamic array, so that Erase frees it later, and the name stays an ordinary
        // Variant: a fresh cscript lets `Dim x` take an array afterwards, whatever size ReDim gave it first.
        const arr = createVbArray([]);
        arr.redim(bounds, false);
        this.context.setVariable(decl.id.name, { type: 'Array', value: arr });
      }
    }

    return VbEmpty;
  }

  private executeEraseStatement(node: VbEraseStatement): VbValue {
    // Erase takes one name, which a fresh cscript shows for `Erase a, b` with error 450 and for `Erase(a)`,
    // where the parentheses turn the name into a value, with error 13.
    if (node.targets.length !== 1) {
      throw createVbError(
        VbErrorCodes.InvalidQualifier,
        'Wrong number of arguments or invalid property assignment: Erase'
      );
    }
    if (node.parenthesized) {
      throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch: Erase');
    }

    const target = node.targets[0]!;
    const value = this.readEraseTarget(target);

    // Erasing something that is not an array is a type mismatch, which a fresh cscript shows for Erase x.
    if (value.type !== 'Array') {
      throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch: Erase');
    }

    const arr = value.value as VbArray;
    arr.erase(arr.isDynamic());

    return VbEmpty;
  }

  /**
   * The value an Erase names, which is a variable or an element of an array and nothing else.
   *
   * `Erase a(0)` frees the array the element holds and leaves a(0) as an array with no size, which a fresh
   * cscript shows with UBound(a(0)) error 9 and IsEmpty(a(0)) False. A call that is no variable, such as
   * `Erase Array(1, 2)`, is error 13.
   */
  private readEraseTarget(target: Expression): VbValue {
    if (target.type === 'Identifier') {
      const variable = this.context.currentScope.get(target.name);
      // A name nothing holds is read the way every other name is: error 500 with Option Explicit, and the
      // Empty it stands for without, which is then the type mismatch Erase answers with.
      return variable ? variable.value : this.context.getVariable(target.name);
    }

    if (
      target.type === 'CallExpression' &&
      target.callee.type === 'Identifier' &&
      target.arguments.length > 0
    ) {
      const container = this.context.currentScope.get(target.callee.name);
      if (container && container.value.type === 'Array') {
        const arr = container.value.value as VbArray;
        const indices = target.arguments.map(argument =>
          toSubscript(this.exprEvaluator.evaluate(argument as Expression))
        );
        return arr.get(indices);
      }
    }

    throw createVbError(VbErrorCodes.TypeMismatch, 'Type mismatch: Erase');
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

    // A bound or a step that is no number is the loop's own error 92, measured on a fresh cscript:
    // "For i = 1 To Null", "For i = 1 To 3 Step Null" and "For i = \"a\" To \"c\"" are all error 92, and not
    // the type mismatch or the invalid use of Null the conversion itself would raise. A text that is a
    // number is read as one, so "For i = 1 To \"3\"" runs three times.
    const readForBound = (value: VbValue): number => {
      try {
        return toNumber(value);
      } catch {
        throw createVbError(92, 'For loop not initialized');
      }
    };

    const startValue = readForBound(init);
    const endValue = readForBound(to);
    const stepValue = readForBound(step);

    this.context.declareVariable(node.left.name, createVbValue(startValue));

    const shouldContinue =
      stepValue > 0
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

  private executeForEachStatement(node: ForOfStatement): VbValue {
    const collection = this.exprEvaluator.evaluate(node.right as Expression);

    if (collection.type !== 'Array' && collection.type !== 'Object') {
      // A value that is no collection is error 451 in Windows, measured on a fresh cscript for a number, a
      // text, an Empty, a Null, a Boolean and a Date alike: "For Each y In 5" answers "Object is not a
      // collection" there, and not the type mismatch this used to raise. An object the engine cannot
      // enumerate answers it too, which the check further down takes care of.
      throw createVbError(451, 'Object is not a collection');
    }

    let items: VbValue[];
    if (collection.type === 'Array') {
      const arr = collection.value as VbArray;
      items = arr.toArray();
    } else {
      const rawObj = collection.value as unknown;
      if (rawObj !== null && typeof rawObj === 'object' && Symbol.iterator in (rawObj as object)) {
        items = [];
        for (const item of rawObj as Iterable<unknown>) {
          if (item === null || item === undefined) {
            items.push(VbEmpty);
          } else if (typeof item === 'boolean') {
            items.push({ type: 'Boolean', value: item });
          } else if (typeof item === 'number') {
            items.push({ type: 'Double', value: item });
          } else if (typeof item === 'string') {
            items.push({ type: 'String', value: item });
          } else {
            items.push({ type: 'Object', value: item });
          }
        }
      } else {
        const obj = collection.value as VbObjectInstance;
        const count = obj.getProperty('Count');
        const countNum = toNumber(count);
        items = [];
        for (let i = 0; i < countNum; i++) {
          items.push(obj.getProperty(String(i)));
        }
      }
    }

    const loopVar = (node.left as import('../ast/index.ts').Identifier).name;
    this.context.declareVariable(loopVar, VbEmpty);

    for (const item of items) {
      this.context.setVariable(loopVar, item);

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

    // For Each leaves its loop variable empty afterwards, which a fresh cscript shows: the variable reads
    // Empty once the loop is over, whether the loop ran or not. A For...Next loop is the other way round and
    // keeps the value that stopped it, so only this loop clears it.
    this.context.setVariable(loopVar, VbEmpty);

    return VbEmpty;
  }

  private executeDoLoopStatement(node: VbDoLoopStatement): VbValue {
    const checkCondition = (): boolean => {
      if (!node.test) return true;
      const testValue = this.exprEvaluator.evaluate(node.test!);
      return toConditionBoolean(testValue);
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

      // An unconditional `Do ... Loop` has nothing to re-check and runs until an Exit Do.
      // Without the node.test guard the absent condition reads as true, the until-branch
      // negates it to false, and the loop leaves after a single pass with nothing raised.
      if (!isPreTest && node.test) {
        const cond = isWhile ? checkCondition() : !checkCondition();
        if (!cond) break;
      }
    }

    return VbEmpty;
  }

  // `While...Wend` - a real, basic VBScript looping construct that never
  // actually worked in this engine at any layer: `Wend` wasn't even a
  // registered lexer keyword (fell through as a plain Identifier, so
  // parseWhileBody() could never recognize it as the loop's terminator and
  // ran to EOF), and even once parsing was fixed there was no executor case
  // for WhileStatement at all. Found via Wine's own vbscript.dll conformance
  // suite, dlls/vbscript/tests/lang.vbs. Unlike Do...Loop, real VBScript has
  // no `Exit While` - a ControlFlowSignal from the body (e.g. `Exit Do`/
  // `Exit For` targeting an enclosing loop, or `Exit Sub`/`Exit Function`)
  // should simply propagate past this loop uncaught, not be special-cased
  // here.
  private executeWhileStatement(node: WhileStatement): VbValue {
    while (toConditionBoolean(this.exprEvaluator.evaluate(node.test))) {
      if (this.context.checkTimeout) this.context.checkTimeout();
      this.execute(node.body);
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

      const matches = this.matchCase(caseClause.test, discriminant, node.discriminant);

      if (matches) {
        for (const stmt of caseClause.consequent) {
          this.execute(stmt);
        }
        return VbEmpty;
      }
    }

    return VbEmpty;
  }

  private matchCase(
    test: Expression | Expression[] | null,
    discriminant: VbValue,
    discriminantNode: Expression
  ): boolean {
    if (!test) return false;

    const tests = Array.isArray(test) ? test : [test];

    for (const t of tests) {
      // Select Case asks the comparison operators' own question of its value, so it reads text case
      // sensitively and raises a type mismatch for a number written out beside a text, exactly as `=` does:
      // measured, `Select Case "b"` matches `Case "b"` and not `Case "B"`, and `Select Case "abc"` with
      // `Case 1` is error 13. A Null answer matches nothing, which is why a Null value and a `Case Null`
      // both fall through to Case Else.
      if (
        t.type === 'BinaryExpression' &&
        'left' in t &&
        (t.left as Expression).type === 'Identifier' &&
        (t.left as Expression as { name: string }).name === '__select_expr__'
      ) {
        const rightNode = t.right as Expression;
        const right = this.exprEvaluator.evaluate(rightNode);
        this.exprEvaluator.checkNumberLiteralBesideText(
          discriminantNode,
          rightNode,
          discriminant,
          right
        );
        const order = this.exprEvaluator.compareValues(discriminant, right);
        if (order === null) continue;

        switch (t.operator) {
          case '==':
            if (order === 0) return true;
            break;
          case '<':
            if (order < 0) return true;
            break;
          case '>':
            if (order > 0) return true;
            break;
          case '<=':
            if (order <= 0) return true;
            break;
          case '>=':
            if (order >= 0) return true;
            break;
          case '!=':
            if (order !== 0) return true;
            break;
        }
      } else {
        const testValue = this.exprEvaluator.evaluate(t);
        this.exprEvaluator.checkNumberLiteralBesideText(
          discriminantNode,
          t,
          discriminant,
          testValue
        );
        if (this.exprEvaluator.compareValues(discriminant, testValue) === 0) {
          return true;
        }
      }
    }

    return false;
  }

  private executeWithStatement(node: WithStatement): VbValue {
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
    // An Exit written in the text a script handed to Execute or ExecuteGlobal cannot leave anything, which a
    // fresh cscript shows with error 1039: `Execute "Exit Sub"` inside a Sub answers 1039 and the Sub carries
    // on, and the same text inside a Function leaves that Function answering Empty. The rest of the handed
    // over text stops there with it.
    if (this.context.insideExecutedText) {
      throw createVbError(VbErrorCodes.CannotExitProcedure, 'Cannot exit a procedure');
    }
    this.context.setExitFlag(
      node.target.toLowerCase() as 'sub' | 'function' | 'property' | 'do' | 'for' | 'select'
    );
    throw new ControlFlowSignal('exit');
  }

  private executeOptionExplicitStatement(_node: VbOptionExplicitStatement): VbValue {
    void _node; // Intentionally unused
    this.context.optionExplicit = true;
    return VbEmpty;
  }

  private executeSubStatement(node: VbSubStatement): VbValue {
    const subName = node.name.name;
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;

    const subFunc = function (args: VbValue[], refs?: VbArgumentRef[]): VbValue {
      self.context.pushScope();
      self.context.pushCall(subName);
      self.context.pushProcedure('sub');

      try {
        self.bindParameters(node.params, args, refs);

        self.executeBlockStatement(node.body);

        self.updateByRefArgs(node.params, args);
      } catch (signal) {
        // `Exit Sub` was never actually handled here - only 'return' was
        // checked, so it always fell through to the `throw signal` below and
        // escaped uncaught past this Sub's own boundary, exactly like the
        // For/Do loops' own exit-flag checks already handle correctly for
        // 'for'/'do'. Found via Wine's own vbscript.dll conformance suite,
        // dlls/vbscript/tests/lang.vbs - a near-universal real VBScript
        // idiom (early-return guard clauses) was completely broken.
        if (
          signal instanceof ControlFlowSignal &&
          ((signal.type === 'exit' && self.context.getExitFlag() === 'sub') ||
            signal.type === 'return')
        ) {
          if (signal.type === 'exit') self.context.clearExitFlag();
          self.updateByRefArgs(node.params, args);
        } else {
          throw signal;
        }
      } finally {
        self.context.popProcedure();
        self.context.popCall();
        self.context.popScope();
      }

      return VbEmpty;
    };

    const params = node.params.map(p => ({
      name: p.name.name,
      byRef: p.byRef,
      isArray: p.isArray,
      isOptional: p.isOptional,
      isParamArray: p.isParamArray,
    }));

    this.context.functionRegistry.register(subName, subFunc, {
      isSub: true,
      params,
      isUserDefined: true,
    });
    return VbEmpty;
  }

  private executeFunctionStatement(node: VbFunctionStatement): VbValue {
    const funcName = node.name.name;
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;

    const func = function (args: VbValue[], refs?: VbArgumentRef[]): VbValue {
      self.context.pushScope();
      self.context.pushCall(funcName);
      self.context.pushProcedure('function');

      self.context.declareVariable(funcName, VbEmpty);

      // Empty until the body answers: a Function that leaves through an error never reaches the return below,
      // so the value here is only ever the answer of a call that finished.
      let result: VbValue = VbEmpty;

      try {
        self.bindParameters(node.params, args, refs);

        self.executeBlockStatement(node.body);

        self.updateByRefArgs(node.params, args);

        result = self.context.getVariable(funcName);
      } catch (signal) {
        // `Exit Function` - same fix and reasoning as executeSubStatement's
        // `Exit Sub` above.
        if (
          signal instanceof ControlFlowSignal &&
          ((signal.type === 'exit' && self.context.getExitFlag() === 'function') ||
            signal.type === 'return')
        ) {
          if (signal.type === 'exit') self.context.clearExitFlag();
          self.updateByRefArgs(node.params, args);
          // An `Execute "Exit Function"` left the function where it stood, so the value the signal carries is
          // the answer: measured, the function answers Empty when it had not written its name yet.
          result = signal.carriesThisStatement
            ? (signal.value ?? VbEmpty)
            : self.context.getVariable(funcName);
        } else {
          throw signal;
        }
      } finally {
        // The answer leaves the procedure, so the reference it holds is kept before the scope below lets go:
        // measured, a `Function FreshMaker()` answering a fresh instance lets the caller name it afterwards,
        // and that instance's Terminate event runs when the caller lets it go, not at the return. The value is
        // marked as one the procedure made, so the place the caller stores it in takes that reference rather
        // than a second one. An Empty answer holds nothing, so a call that left through an error keeps nothing.
        markMadeValue(result);
        self.context.popProcedure();
        self.context.popCall();
        self.context.popScope();
      }

      return result;
    };
    const params = node.params.map(p => ({
      name: p.name.name,
      byRef: p.byRef,
      isArray: p.isArray,
      isOptional: p.isOptional,
      isParamArray: p.isParamArray,
    }));

    this.context.functionRegistry.register(funcName, func, { params, isUserDefined: true });
    return VbEmpty;
  }

  private bindParameters(
    params: import('../ast/index.ts').VbParameter[],
    args: VbValue[],
    refs?: VbArgumentRef[]
  ): void {
    let argIndex = 0;

    for (const param of params) {
      if (param.isParamArray) {
        const restArgs: VbValue[] = [];
        while (argIndex < args.length) {
          restArgs.push(args[argIndex++]!);
        }
        const arr = createVbArray([restArgs.length]);
        restArgs.forEach((v, idx) => arr.set([idx], v));
        this.context.declareVariable(param.name.name, { type: 'Array', value: arr });
      } else {
        let argValue: VbValue;
        // A parameter takes on the array-ness of the variable it was handed, so writing a whole array into it
        // is the error 13 the variable itself would answer with.
        const isArrayParameter = refs?.[argIndex]?.isArrayVariable === true;

        if (argIndex < args.length) {
          argValue = args[argIndex++]!;
        } else if (param.defaultValue) {
          argValue = this.exprEvaluator.evaluate(param.defaultValue);
        } else if (param.isOptional) {
          argValue = VbEmpty;
        } else {
          argValue = VbEmpty;
        }

        if (!param.byRef) {
          // A ByVal parameter holds a copy of its own, which is measured: the instance it is handed is freed
          // when the procedure returns, even though the caller's variable still holds it.
          retainValue(argValue);
          argValue = copyArrayArgument(argValue);
        }

        this.context.declareVariable(param.name.name, argValue, {
          isArray: isArrayParameter,
          isByRef: param.byRef,
        });
      }
    }
  }

  private updateByRefArgs(params: import('../ast/index.ts').VbParameter[], args: VbValue[]): void {
    let argIndex = 0;

    for (const param of params) {
      if (param.isParamArray) {
        break;
      }

      if (argIndex < args.length && param.byRef) {
        args[argIndex] = this.context.getVariable(param.name.name);
      }
      argIndex++;
    }
  }

  private executeClassStatement(node: VbClassStatement): VbValue {
    const className = node.name.name;
    const cls = new VbClass(className);
    // eslint-disable-next-line @typescript-eslint/no-this-alias
    const self = this;
    let initializeFunc: VbProperty | undefined;
    let terminateFunc: VbProperty | undefined;

    for (const member of node.body) {
      if (member.type === 'VbDimStatement') {
        for (const decl of member.declarations) {
          cls.properties.set(decl.id.name.toLowerCase(), {
            name: decl.id.name,
            isPrivate: member.visibility === 'private',
          });
        }
      } else if (member.type === 'VbSubStatement') {
        const memberNode = member;
        const memberName = member.name.name.toLowerCase();
        const memberArity = declaredArity(classParameterInfo(memberNode.params));

        if (memberName === 'class_initialize') {
          initializeFunc = {
            name: 'Class_Initialize',
            // Reachable from inside the class only, which a fresh cscript shows: naming it from script level,
            // with or without parentheses and whether the class wrote it Private or Public, is error 438
            // without running it. Only `New` runs it.
            isPrivate: true,
            get: function (this: VbObjectInstance): VbValue {
              const prevInstance = self.context.currentInstance;
              self.context.currentInstance = this;
              self.context.pushScope();
              try {
                self.executeBlockStatement(memberNode.body);
              } catch (signal) {
                if (!(
                  signal instanceof ControlFlowSignal &&
                  signal.type === 'exit' &&
                  self.context.getExitFlag() === 'sub'
                )) {
                  throw signal;
                }
                self.context.clearExitFlag();
              } finally {
                self.context.popScope();
                self.context.currentInstance = prevInstance;
              }
              return VbEmpty;
            },
          };
          continue;
        }

        if (memberName === 'class_terminate') {
          terminateFunc = {
            name: 'Class_Terminate',
            // Same reading as Initialize above: naming it from script level is error 438 without running it.
            isPrivate: true,
            get: function (this: VbObjectInstance): VbValue {
              const prevInstance = self.context.currentInstance;
              self.context.currentInstance = this;
              self.context.pushScope();
              try {
                self.executeBlockStatement(memberNode.body);
              } catch (signal) {
                if (!(
                  signal instanceof ControlFlowSignal &&
                  signal.type === 'exit' &&
                  self.context.getExitFlag() === 'sub'
                )) {
                  throw signal;
                }
                self.context.clearExitFlag();
              } finally {
                self.context.popScope();
                self.context.currentInstance = prevInstance;
              }
              return VbEmpty;
            },
          };
          continue;
        }

        cls.methods.set(memberName, {
          name: member.name.name,
          isPrivate: memberNode.visibility === 'private',
          func: function (this: VbObjectInstance, ...args: VbValue[]): VbValue {
            checkArgumentRange(member.name.name, memberArity[0], memberArity[1], args.length);
            const prevInstance = self.context.currentInstance;
            self.context.currentInstance = this;
            self.context.pushScope();
            self.context.pushProcedure('sub');
            try {
              self.bindParameters(memberNode.params, args);
              self.executeBlockStatement(memberNode.body);
            } catch (signal) {
              if (!(
                signal instanceof ControlFlowSignal &&
                signal.type === 'exit' &&
                self.context.getExitFlag() === 'sub'
              )) {
                throw signal;
              }
              self.context.clearExitFlag();
            } finally {
              self.context.popProcedure();
              self.context.popScope();
              self.context.currentInstance = prevInstance;
            }
            return VbEmpty;
          },
          isSub: true,
        });
      } else if (member.type === 'VbFunctionStatement') {
        const memberNode = member;
        const memberArity = declaredArity(classParameterInfo(memberNode.params));
        cls.methods.set(member.name.name.toLowerCase(), {
          name: member.name.name,
          isPrivate: memberNode.visibility === 'private',
          func: function (this: VbObjectInstance, ...args: VbValue[]): VbValue {
            checkArgumentRange(member.name.name, memberArity[0], memberArity[1], args.length);
            const prevInstance = self.context.currentInstance;
            self.context.currentInstance = this;
            self.context.pushScope();
            self.context.pushProcedure('function');
            self.context.declareVariable(memberNode.name.name, VbEmpty);
            self.bindParameters(memberNode.params, args);
            let result: VbValue = VbEmpty;
            try {
              self.executeBlockStatement(memberNode.body);
              result = self.context.getVariable(memberNode.name.name);
            } catch (signal) {
              if (
                signal instanceof ControlFlowSignal &&
                ((signal.type === 'exit' && self.context.getExitFlag() === 'function') ||
                  signal.type === 'return')
              ) {
                if (signal.type === 'exit') self.context.clearExitFlag();
                result = signal.carriesThisStatement
                  ? (signal.value ?? VbEmpty)
                  : self.context.getVariable(memberNode.name.name);
              } else {
                throw signal;
              }
            } finally {
              // Same reading as a plain Function's answer above: the answer's own reference is kept, so a
              // member that answers a fresh instance hands the caller a reference that survives the return. An
              // Empty answer holds nothing, so a member that left through an error keeps nothing.
              markMadeValue(result);
              self.context.popProcedure();
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
        const existing = cls.properties.get(propName) || {
          name: member.name.name,
          isPrivate: memberNode.visibility === 'private',
        };
        existing.get = function (this: VbObjectInstance): VbValue {
          const prevInstance = self.context.currentInstance;
          const prevInPropertyGet = self.context.inPropertyGet;
          const prevPropertyGetName = self.context.propertyGetName;
          self.context.currentInstance = this;
          self.context.inPropertyGet = true;
          self.context.propertyGetName = memberNode.name.name.toLowerCase();
          self.context.pushScope();
          self.context.pushProcedure('property');
          self.context.declareVariable(memberNode.name.name, VbEmpty);
          let result: VbValue;
          try {
            self.executeBlockStatement(memberNode.body);
            result = self.context.currentScope.get(memberNode.name.name)?.value ?? VbEmpty;
          } catch (signal) {
            if (!(
              signal instanceof ControlFlowSignal &&
              signal.type === 'exit' &&
              self.context.getExitFlag() === 'property'
            )) {
              throw signal;
            }
            self.context.clearExitFlag();
            result = signal.carriesThisStatement
              ? (signal.value ?? VbEmpty)
              : (self.context.currentScope.get(memberNode.name.name)?.value ?? VbEmpty);
          } finally {
            // Same reading as a plain Function's answer above: a Property Get that answers an object hands the
            // caller a reference of its own, which survives the scope the property ran in.
            markMadeValue(result!);
            self.context.popProcedure();
            self.context.popScope();
            self.context.currentInstance = prevInstance;
            self.context.inPropertyGet = prevInPropertyGet;
            self.context.propertyGetName = prevPropertyGetName;
          }
          return result;
        };
        cls.properties.set(propName, existing);
      } else if (member.type === 'VbPropertyLetStatement') {
        const memberNode = member;
        const propName = member.name.name.toLowerCase();
        const existing = cls.properties.get(propName) || {
          name: member.name.name,
          isPrivate: memberNode.visibility === 'private',
        };
        existing.let = function (this: VbObjectInstance, value: VbValue): void {
          const prevInstance = self.context.currentInstance;
          self.context.currentInstance = this;
          self.context.pushScope();
          self.context.declareVariable(memberNode.name.name, VbEmpty);
          const param = memberNode.params[0];
          if (param) {
            // A Property Let takes its value as a copy of its own, which the parameter lets go of on return.
            retainValue(value);
            self.context.declareVariable(param.name.name, value);
          }
          try {
            self.executeBlockStatement(memberNode.body);
          } catch (signal) {
            if (!(
              signal instanceof ControlFlowSignal &&
              signal.type === 'exit' &&
              self.context.getExitFlag() === 'property'
            )) {
              throw signal;
            }
            self.context.clearExitFlag();
          } finally {
            self.context.popScope();
            self.context.currentInstance = prevInstance;
          }
        };
        cls.properties.set(propName, existing);
      } else if (member.type === 'VbPropertySetStatement') {
        const memberNode = member;
        const propName = member.name.name.toLowerCase();
        const existing = cls.properties.get(propName) || {
          name: member.name.name,
          isPrivate: memberNode.visibility === 'private',
        };
        existing.set = function (this: VbObjectInstance, value: VbValue): void {
          const prevInstance = self.context.currentInstance;
          self.context.currentInstance = this;
          self.context.pushScope();
          self.context.declareVariable(memberNode.name.name, VbEmpty);
          const param = memberNode.params[0];
          if (param) {
            // A Property Set takes its value as a copy of its own, which the parameter lets go of on return.
            retainValue(value);
            self.context.declareVariable(param.name.name, value);
          }
          try {
            self.executeBlockStatement(memberNode.body);
          } catch (signal) {
            if (!(
              signal instanceof ControlFlowSignal &&
              signal.type === 'exit' &&
              self.context.getExitFlag() === 'property'
            )) {
              throw signal;
            }
            self.context.clearExitFlag();
          } finally {
            self.context.popScope();
            self.context.currentInstance = prevInstance;
          }
        };
        cls.properties.set(propName, existing);
      }
    }

    if (initializeFunc) {
      cls.properties.set('class_initialize', initializeFunc);
    }
    if (terminateFunc) {
      cls.properties.set('class_terminate', terminateFunc);
    }
    this.context.classRegistry.register(cls);
    return VbEmpty;
  }

  private executePropertyStatement(_node: Statement): VbValue {
    void _node; // Intentionally unused
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

  private executeGotoStatement(node: VbGotoStatement): VbValue {
    throw new GotoSignal(node.label.name.toLowerCase());
  }
}
