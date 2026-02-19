import type { Position, SourceLocation } from '../ast/base.ts';
import type { Token } from '../lexer/index.ts';

export function tokenToPosition(token: Token): Position {
  return {
    line: token.loc.start.line,
    column: token.loc.start.column,
  };
}

export function tokenToEndPosition(token: Token): Position {
  return {
    line: token.loc.end.line,
    column: token.loc.end.column,
  };
}

export function createLocation(start: Token, end: Token): SourceLocation {
  return {
    start: tokenToPosition(start),
    end: tokenToEndPosition(end),
  };
}

export function mergeLocations(start: SourceLocation, end: SourceLocation): SourceLocation {
  return {
    start: start.start,
    end: end.end,
  };
}
