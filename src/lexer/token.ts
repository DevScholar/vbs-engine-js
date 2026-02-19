export enum TokenType {
  EOF = 'EOF',
  Newline = 'Newline',
  WhiteSpace = 'WhiteSpace',
  
  Identifier = 'Identifier',
  Unknown = 'Unknown',
  
  StringLiteral = 'StringLiteral',
  NumberLiteral = 'NumberLiteral',
  DateLiteral = 'DateLiteral',
  BooleanLiteral = 'BooleanLiteral',
  NothingLiteral = 'NothingLiteral',
  NullLiteral = 'NullLiteral',
  EmptyLiteral = 'EmptyLiteral',
  
  Plus = 'Plus',
  Minus = 'Minus',
  Asterisk = 'Asterisk',
  Slash = 'Slash',
  Backslash = 'Backslash',
  Caret = 'Caret',
  Mod = 'Mod',
  
  Ampersand = 'Ampersand',
  
  Eq = 'Eq',
  Lt = 'Lt',
  Gt = 'Gt',
  Le = 'Le',
  Ge = 'Ge',
  Ne = 'Ne',
  Is = 'Is',
  
  And = 'And',
  Or = 'Or',
  Not = 'Not',
  Xor = 'Xor',
  Eqv = 'Eqv',
  Imp = 'Imp',
  
  LParen = 'LParen',
  RParen = 'RParen',
  LBrace = 'LBrace',
  RBrace = 'RBrace',
  LBracket = 'LBracket',
  RBracket = 'RBracket',
  
  Comma = 'Comma',
  Colon = 'Colon',
  Dot = 'Dot',
  Bang = 'Bang',
  
  Dim = 'Dim',
  ReDim = 'ReDim',
  Preserve = 'Preserve',
  
  Public = 'Public',
  Private = 'Private',
  
  Sub = 'Sub',
  Function = 'Function',
  End = 'End',
  Exit = 'Exit',
  
  Call = 'Call',
  Set = 'Set',
  Let = 'Let',
  Get = 'Get',
  
  If = 'If',
  Then = 'Then',
  Else = 'Else',
  ElseIf = 'ElseIf',
  
  For = 'For',
  To = 'To',
  Step = 'Step',
  Next = 'Next',
  Each = 'Each',
  In = 'In',
  
  Do = 'Do',
  Loop = 'Loop',
  While = 'While',
  Until = 'Until',
  
  Select = 'Select',
  Case = 'Case',
  
  On = 'On',
  Error = 'Error',
  Resume = 'Resume',
  Goto = 'Goto',
  
  With = 'With',
  
  Class = 'Class',
  Property = 'Property',
  New = 'New',
  
  Const = 'Const',
  
  Option = 'Option',
  Explicit = 'Explicit',
  
  ByRef = 'ByRef',
  ByVal = 'ByVal',
  Optional = 'Optional',
  ParamArray = 'ParamArray',
  
  Rem = 'Rem',
  
  As = 'As',
  
  Integer = 'Integer',
  Long = 'Long',
  Single = 'Single',
  Double = 'Double',
  Currency = 'Currency',
  String = 'String',
  Boolean = 'Boolean',
  Date = 'Date',
  Object = 'Object',
  Variant = 'Variant',
  Byte = 'Byte',
}

export interface TokenLocation {
  line: number;
  column: number;
  offset: number;
}

export interface Token {
  type: TokenType;
  value: string;
  loc: {
    start: TokenLocation;
    end: TokenLocation;
  };
  raw?: string;
}
