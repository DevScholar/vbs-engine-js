export type { Position, SourceLocation, BaseNode, BaseExpression, BaseStatement, BaseDeclaration, BasePattern } from './base.ts';
export * from './types.ts';

import type { Expression, Statement } from './types.ts';
import type { VbNode } from './types.ts';

export type { VbNode };

export function isExpression(node: VbNode): node is Expression {
  const expressionTypes = [
    'Identifier',
    'Literal',
    'ArrayExpression',
    'ObjectExpression',
    'FunctionExpression',
    'ArrowFunctionExpression',
    'UnaryExpression',
    'UpdateExpression',
    'BinaryExpression',
    'AssignmentExpression',
    'LogicalExpression',
    'MemberExpression',
    'ConditionalExpression',
    'CallExpression',
    'NewExpression',
    'SequenceExpression',
    'SpreadElement',
    'ChainExpression',
    'ThisExpression',
    'VbEmptyLiteral',
    'VbNewExpression',
    'VbWithObject',
  ];
  return expressionTypes.includes(node.type);
}

export function isStatement(node: VbNode): node is Statement {
  const statementTypes = [
    'ExpressionStatement',
    'BlockStatement',
    'EmptyStatement',
    'DebuggerStatement',
    'WithStatement',
    'ReturnStatement',
    'LabeledStatement',
    'BreakStatement',
    'ContinueStatement',
    'IfStatement',
    'SwitchStatement',
    'ThrowStatement',
    'TryStatement',
    'WhileStatement',
    'DoWhileStatement',
    'ForStatement',
    'ForInStatement',
    'ForOfStatement',
    'VariableDeclaration',
    'VbSubStatement',
    'VbFunctionStatement',
    'VbClassStatement',
    'VbPropertyGetStatement',
    'VbPropertyLetStatement',
    'VbPropertySetStatement',
    'VbDimStatement',
    'VbReDimStatement',
    'VbConstStatement',
    'VbForToStatement',
    'VbForEachStatement',
    'VbDoLoopStatement',
    'VbSelectCaseStatement',
    'VbWithStatement',
    'VbOnErrorHandlerStatement',
    'VbExitStatement',
    'VbOptionExplicitStatement',
  ];
  return statementTypes.includes(node.type);
}
