export interface Position {
  line: number;
  column: number;
  offset?: number;
}

export interface SourceLocation {
  start: Position;
  end: Position;
  source?: string | null;
}

export interface BaseNode {
  type: string;
  loc?: SourceLocation | null;
  range?: [number, number];
}

export interface BaseExpression extends BaseNode {
  type: string;
}

export interface BaseStatement extends BaseNode {
  type: string;
}

export interface BaseDeclaration extends BaseStatement {
  type: string;
}

export interface BasePattern extends BaseNode {
  type: string;
}
