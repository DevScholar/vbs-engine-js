import type { BaseExpression, BaseNode, BasePattern, BaseStatement, BaseDeclaration } from './base.ts';

export interface Identifier extends BaseNode {
  type: 'Identifier';
  name: string;
}

export interface Literal extends BaseExpression {
  type: 'Literal';
  value: string | number | boolean | null | RegExp | bigint | Date;
  raw?: string;
}

export interface ArrayExpression extends BaseExpression {
  type: 'ArrayExpression';
  elements: Array<Expression | SpreadElement | null>;
}

export interface ObjectExpression extends BaseExpression {
  type: 'ObjectExpression';
  properties: Array<Property | SpreadElement>;
}

export interface Property extends BaseNode {
  type: 'Property';
  key: Expression | PrivateIdentifier;
  value: Expression | Pattern;
  kind: 'init' | 'get' | 'set';
  method: boolean;
  shorthand: boolean;
  computed: boolean;
}

export interface FunctionExpression extends BaseExpression {
  type: 'FunctionExpression';
  id: Identifier | null;
  params: Pattern[];
  body: BlockStatement;
  generator: boolean;
  async: boolean;
}

export interface ArrowFunctionExpression extends BaseExpression {
  type: 'ArrowFunctionExpression';
  params: Pattern[];
  body: Expression | BlockStatement;
  expression: boolean;
}

export interface UnaryExpression extends BaseExpression {
  type: 'UnaryExpression';
  operator: UnaryOperator;
  prefix: boolean;
  argument: Expression;
}

export type UnaryOperator = '-' | '+' | '!' | '~' | 'typeof' | 'void' | 'delete' | 'Not';

export interface UpdateExpression extends BaseExpression {
  type: 'UpdateExpression';
  operator: UpdateOperator;
  argument: Expression;
  prefix: boolean;
}

export type UpdateOperator = '++' | '--';

export interface BinaryExpression extends BaseExpression {
  type: 'BinaryExpression';
  operator: BinaryOperator;
  left: Expression;
  right: Expression;
}

export type BinaryOperator =
  | '=='
  | '!='
  | '==='
  | '!=='
  | '<'
  | '<='
  | '>'
  | '>='
  | '<<'
  | '>>'
  | '>>>'
  | '+'
  | '-'
  | '*'
  | '/'
  | '%'
  | '**'
  | '|'
  | '^'
  | '&'
  | 'in'
  | 'instanceof'
  | '='
  | '<>'
  | '\\'
  | 'Mod'
  | 'Is';

export interface AssignmentExpression extends BaseExpression {
  type: 'AssignmentExpression';
  operator: AssignmentOperator;
  left: Pattern | MemberExpression;
  right: Expression;
  isSet?: boolean;
}

export type AssignmentOperator =
  | '='
  | '+='
  | '-='
  | '*='
  | '/='
  | '%='
  | '**='
  | '<<='
  | '>>='
  | '>>>='
  | '|='
  | '^='
  | '&=';

export interface LogicalExpression extends BaseExpression {
  type: 'LogicalExpression';
  operator: LogicalOperator;
  left: Expression;
  right: Expression;
}

export type LogicalOperator = '||' | '&&' | '??' | 'And' | 'Or' | 'Xor' | 'Eqv' | 'Imp';

export interface MemberExpression extends BaseExpression {
  type: 'MemberExpression';
  object: Expression | Super;
  property: Expression | PrivateIdentifier;
  computed: boolean;
  optional: boolean;
}

export interface ConditionalExpression extends BaseExpression {
  type: 'ConditionalExpression';
  test: Expression;
  alternate: Expression;
  consequent: Expression;
}

export interface CallExpression extends BaseExpression {
  type: 'CallExpression';
  callee: Expression | Super;
  arguments: Array<Expression | SpreadElement>;
  optional: boolean;
}

export interface NewExpression extends BaseExpression {
  type: 'NewExpression';
  callee: Expression;
  arguments: Array<Expression | SpreadElement>;
}

export interface SequenceExpression extends BaseExpression {
  type: 'SequenceExpression';
  expressions: Expression[];
}

export interface SpreadElement extends BaseNode {
  type: 'SpreadElement';
  argument: Expression;
}

export interface Super extends BaseNode {
  type: 'Super';
}

export interface ThisExpression extends BaseNode {
  type: 'ThisExpression';
}

export interface PrivateIdentifier extends BaseNode {
  type: 'PrivateIdentifier';
  name: string;
}

export interface ChainExpression extends BaseExpression {
  type: 'ChainExpression';
  expression: CallExpression | MemberExpression;
}

export interface VbEmptyLiteral extends BaseExpression {
  type: 'VbEmptyLiteral';
  value: undefined;
  raw?: string;
}

export interface VbNewExpression extends BaseExpression {
  type: 'VbNewExpression';
  callee: Identifier;
  arguments: Expression[];
}

export interface VbWithObjectExpression extends BaseExpression {
  type: 'VbWithObject';
}

export interface Program extends BaseNode {
  type: 'Program';
  body: Statement[];
  sourceType?: 'script';
}

export interface ExpressionStatement extends BaseStatement {
  type: 'ExpressionStatement';
  expression: Expression;
  directive?: string;
}

export interface BlockStatement extends BaseStatement {
  type: 'BlockStatement';
  body: Statement[];
}

export interface EmptyStatement extends BaseStatement {
  type: 'EmptyStatement';
}

export interface DebuggerStatement extends BaseStatement {
  type: 'DebuggerStatement';
}

export interface WithStatement extends BaseStatement {
  type: 'WithStatement';
  object: Expression;
  body: Statement;
}

export interface ReturnStatement extends BaseStatement {
  type: 'ReturnStatement';
  argument: Expression | null;
}

export interface LabeledStatement extends BaseStatement {
  type: 'LabeledStatement';
  label: Identifier;
  body: Statement;
}

export interface BreakStatement extends BaseStatement {
  type: 'BreakStatement';
  label: Identifier | null;
}

export interface ContinueStatement extends BaseStatement {
  type: 'ContinueStatement';
  label: Identifier | null;
}

export interface IfStatement extends BaseStatement {
  type: 'IfStatement';
  test: Expression;
  consequent: Statement;
  alternate: Statement | null;
}

export interface SwitchStatement extends BaseStatement {
  type: 'SwitchStatement';
  discriminant: Expression;
  cases: SwitchCase[];
}

export interface SwitchCase extends BaseNode {
  type: 'SwitchCase';
  test: Expression | null;
  consequent: Statement[];
}

export interface ThrowStatement extends BaseStatement {
  type: 'ThrowStatement';
  argument: Expression;
}

export interface TryStatement extends BaseStatement {
  type: 'TryStatement';
  block: BlockStatement;
  handler: CatchClause | null;
  finalizer: BlockStatement | null;
}

export interface CatchClause extends BaseNode {
  type: 'CatchClause';
  param: Pattern | null;
  body: BlockStatement;
}

export interface WhileStatement extends BaseStatement {
  type: 'WhileStatement';
  test: Expression;
  body: Statement;
}

export interface DoWhileStatement extends BaseStatement {
  type: 'DoWhileStatement';
  body: Statement;
  test: Expression;
}

export interface ForStatement extends BaseStatement {
  type: 'ForStatement';
  init: VariableDeclaration | Expression | null;
  test: Expression | null;
  update: Expression | null;
  body: Statement;
}

export interface ForInStatement extends BaseStatement {
  type: 'ForInStatement';
  left: VariableDeclaration | Pattern;
  right: Expression;
  body: Statement;
}

export interface ForOfStatement extends BaseStatement {
  type: 'ForOfStatement';
  left: VariableDeclaration | Pattern;
  right: Expression;
  body: Statement;
}

export interface VariableDeclaration extends BaseDeclaration {
  type: 'VariableDeclaration';
  declarations: VariableDeclarator[];
  kind: 'var' | 'let' | 'const';
}

export interface VariableDeclarator extends BaseNode {
  type: 'VariableDeclarator';
  id: Pattern;
  init: Expression | null;
}

export interface VbParameter extends BaseNode {
  type: 'VbParameter';
  name: Identifier;
  byRef: boolean;
  defaultValue?: Expression;
  isArray: boolean;
  isOptional: boolean;
  isParamArray: boolean;
}

export interface VbSubStatement extends BaseStatement {
  type: 'VbSubStatement';
  name: Identifier;
  params: VbParameter[];
  body: BlockStatement;
  visibility: 'public' | 'private' | 'default';
}

export interface VbFunctionStatement extends BaseStatement {
  type: 'VbFunctionStatement';
  name: Identifier;
  params: VbParameter[];
  body: BlockStatement;
  visibility: 'public' | 'private' | 'default';
}

export interface VbVariableDeclarator extends BaseNode {
  type: 'VbVariableDeclarator';
  id: Identifier;
  init: Expression | null;
  isArray: boolean;
  arrayBounds?: Expression[];
  typeAnnotation?: VbTypeAnnotation;
}

export interface VbTypeAnnotation extends BaseNode {
  type: 'VbTypeAnnotation';
  typeName: string;
  isArray: boolean;
}

export interface VbDimStatement extends BaseStatement {
  type: 'VbDimStatement';
  declarations: VbVariableDeclarator[];
  visibility: 'public' | 'private' | 'dim';
}

export interface VbReDimStatement extends BaseStatement {
  type: 'VbReDimStatement';
  declarations: VbVariableDeclarator[];
  preserve: boolean;
}

export interface VbConstStatement extends BaseStatement {
  type: 'VbConstStatement';
  declarations: VbConstDeclarator[];
  visibility: 'public' | 'private';
}

export interface VbConstDeclarator extends BaseNode {
  type: 'VbConstDeclarator';
  id: Identifier;
  init: Expression;
  typeAnnotation?: VbTypeAnnotation;
}

export interface VbForToStatement extends BaseStatement {
  type: 'VbForToStatement';
  left: Identifier;
  init: Expression;
  to: Expression;
  step: Expression | null;
  body: Statement;
}

export interface VbForEachStatement extends BaseStatement {
  type: 'VbForEachStatement';
  left: Identifier;
  right: Expression;
  body: Statement;
}

export interface VbDoLoopStatement extends BaseStatement {
  type: 'VbDoLoopStatement';
  body: Statement;
  test: Expression | null;
  testPosition: 'while-do' | 'do-while' | 'until-do' | 'do-until' | null;
}

export interface VbSelectCaseStatement extends BaseStatement {
  type: 'VbSelectCaseStatement';
  discriminant: Expression;
  cases: VbCaseClause[];
}

export interface VbCaseClause extends BaseNode {
  type: 'VbCaseClause';
  test: Expression | Expression[] | null;
  consequent: Statement[];
  isElse: boolean;
}

export interface VbWithStatement extends BaseStatement {
  type: 'VbWithStatement';
  object: Expression;
  body: Statement;
}

export interface VbOnErrorHandlerStatement extends BaseStatement {
  type: 'VbOnErrorHandlerStatement';
  action: 'resume_next' | 'goto_0' | 'goto_label';
  label?: Identifier;
}

export interface VbExitStatement extends BaseStatement {
  type: 'VbExitStatement';
  target: 'Sub' | 'Function' | 'Property' | 'Do' | 'For' | 'Select';
}

export interface VbOptionExplicitStatement extends BaseStatement {
  type: 'VbOptionExplicitStatement';
}

export interface VbClassStatement extends BaseStatement {
  type: 'VbClassStatement';
  name: Identifier;
  body: VbClassElement[];
}

export type VbClassElement =
  | VbSubStatement
  | VbFunctionStatement
  | VbPropertyGetStatement
  | VbPropertyLetStatement
  | VbPropertySetStatement
  | VbDimStatement
  | VbConstStatement;

export interface VbPropertyGetStatement extends BaseStatement {
  type: 'VbPropertyGetStatement';
  name: Identifier;
  params: VbParameter[];
  body: BlockStatement;
  visibility: 'public' | 'private' | 'default';
}

export interface VbPropertyLetStatement extends BaseStatement {
  type: 'VbPropertyLetStatement';
  name: Identifier;
  params: VbParameter[];
  body: BlockStatement;
  visibility: 'public' | 'private' | 'default';
}

export interface VbPropertySetStatement extends BaseStatement {
  type: 'VbPropertySetStatement';
  name: Identifier;
  params: VbParameter[];
  body: BlockStatement;
  visibility: 'public' | 'private' | 'default';
}

export interface VbCallStatement extends BaseStatement {
  type: 'VbCallStatement';
  callee: Expression;
  arguments: Expression[];
}

export interface VbSetStatement extends BaseStatement {
  type: 'VbSetStatement';
  left: Pattern;
  right: Expression;
}

export interface VbLetStatement extends BaseStatement {
  type: 'VbLetStatement';
  left: Pattern;
  right: Expression;
}

export interface VbOnErrorStatement extends BaseStatement {
  type: 'VbOnErrorStatement';
  handler: 'ResumeNext' | 'GoTo0' | { label: string };
}

export interface VbResumeStatement extends BaseStatement {
  type: 'VbResumeStatement';
  target: 'next' | null | { label: string };
}

export interface VbGotoStatement extends BaseStatement {
  type: 'VbGotoStatement';
  label: Identifier;
}

export interface VbLabelStatement extends BaseStatement {
  type: 'VbLabelStatement';
  label: Identifier;
}

export interface ObjectPattern extends BasePattern {
  type: 'ObjectPattern';
  properties: Array<Property | RestElement>;
}

export interface ArrayPattern extends BasePattern {
  type: 'ArrayPattern';
  elements: Array<Pattern | null>;
}

export interface RestElement extends BasePattern {
  type: 'RestElement';
  argument: Pattern;
}

export interface AssignmentPattern extends BasePattern {
  type: 'AssignmentPattern';
  left: Pattern;
  right: Expression;
}

export type Pattern = Identifier | ObjectPattern | ArrayPattern | RestElement | AssignmentPattern;

export type Expression =
  | Identifier
  | Literal
  | ArrayExpression
  | ObjectExpression
  | FunctionExpression
  | ArrowFunctionExpression
  | UnaryExpression
  | UpdateExpression
  | BinaryExpression
  | AssignmentExpression
  | LogicalExpression
  | MemberExpression
  | ConditionalExpression
  | CallExpression
  | NewExpression
  | SequenceExpression
  | SpreadElement
  | ChainExpression
  | VbDateLiteral
  | VbNothingLiteral
  | VbEmptyLiteral
  | VbNullLiteral
  | VbstringConcatExpression
  | VbIsExpression
  | VbNewExpression
  | VbMeExpression
  | VbIndexExpression
  | VbDefaultPropertyExpression
  | VbWithObjectExpression;

export type Statement =
  | ExpressionStatement
  | BlockStatement
  | EmptyStatement
  | DebuggerStatement
  | WithStatement
  | ReturnStatement
  | LabeledStatement
  | BreakStatement
  | ContinueStatement
  | IfStatement
  | SwitchStatement
  | ThrowStatement
  | TryStatement
  | WhileStatement
  | DoWhileStatement
  | ForStatement
  | ForInStatement
  | ForOfStatement
  | VariableDeclaration
  | VbSubStatement
  | VbFunctionStatement
  | VbClassStatement
  | VbPropertyGetStatement
  | VbPropertyLetStatement
  | VbPropertySetStatement
  | VbDimStatement
  | VbReDimStatement
  | VbConstStatement
  | VbForToStatement
  | VbForEachStatement
  | VbDoLoopStatement
  | VbSelectCaseStatement
  | VbWithStatement
  | VbOnErrorHandlerStatement
  | VbExitStatement
  | VbOptionExplicitStatement
  | VbCallStatement
  | VbSetStatement
  | VbLetStatement
  | VbOnErrorStatement
  | VbResumeStatement
  | VbGotoStatement
  | VbLabelStatement;

export type Declaration = VariableDeclaration | VbSubStatement | VbFunctionStatement | VbClassStatement;

export type VbNode = Expression | Statement | VbParameter | VbVariableDeclarator | VbConstDeclarator | VbTypeAnnotation | VbCaseClause | Property | VariableDeclarator | SwitchCase | CatchClause | Pattern;
