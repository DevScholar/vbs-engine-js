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

/**
 * Creates a SourceLocation from a node with a location property.
 * Used when we need to create a location from an AST node rather than a Token.
 */
export function createLocationFromNode(
  start: { loc?: SourceLocation | null },
  end: { loc?: SourceLocation | null }
): SourceLocation {
  if (!start.loc || !end.loc) {
    return {
      start: { line: 0, column: 0 },
      end: { line: 0, column: 0 },
    };
  }
  return {
    start: start.loc.start,
    end: end.loc.end,
  };
}

/**
 * Creates a SourceLocation from a Token and a node with location.
 * Useful when mixing Token (from lexer) and Expression nodes.
 */
export function createLocationFromTokenAndNode(
  start: Token,
  end: { loc?: SourceLocation | null }
): SourceLocation {
  if (!end.loc) {
    return {
      start: tokenToPosition(start),
      end: { line: 0, column: 0 },
    };
  }
  return {
    start: tokenToPosition(start),
    end: end.loc.end,
  };
}

/**
 * Creates a SourceLocation from a node with location and a Token.
 * Useful when mixing Expression nodes and Token (from lexer).
 */
export function createLocationFromNodeAndToken(
  start: { loc?: SourceLocation | null },
  end: Token
): SourceLocation {
  if (!start.loc) {
    return {
      start: { line: 0, column: 0 },
      end: tokenToEndPosition(end),
    };
  }
  return {
    start: start.loc.start,
    end: tokenToEndPosition(end),
  };
}
