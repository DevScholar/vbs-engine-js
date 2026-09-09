# Fix 8 real parsing/evaluation bugs (closes #1), verified against Wine's vbscript.dll conformance suite

## Summary

Found and fixed 8 real bugs while using this engine to run actual VPX (Visual Pinball) table
scripts — genuine, complex, third-party-authored VBScript, not synthetic test cases. One of them
(nested multi-line `If...End If`) is this repo's own open **#1**. The rest were found the same
way, plus cross-checked against Wine's own `vbscript.dll` test suite
(`dlls/vbscript/tests/lang.vbs`) for real MS-accuracy behavior rather than just internal
self-consistency.

A small harness (`wine-conformance/run-lang-vbs.mjs`) is included — it runs Wine's actual
`lang.vbs` (LGPL 2.1, included verbatim) against this engine using only two host-surface
functions (`ok()`/`getVT()`), and is what caught several of the lexer-level bugs below.

## Fixes

### Closes #1 — Nested multi-line `If...End If` parsing (`control-flow.ts`)

A multi-line `If` nested inside another `If`/`Else` block threw `Unexpected token:
Else`/`ElseIf`. The previous implementation manually re-scanned tokens ahead of a nested `If` to
classify it single- vs multi-line, then tracked a hand-rolled `nestedIfDepth` counter to decide
whether a later `Else`/`ElseIf`/`End If` belonged to the nested `If` or the outer block. Once
`depth > 0`, it fed the bare `Else`/`ElseIf` token straight into `parseStatement()` as if it were
its own statement — but those tokens are never valid statement-starters on their own.

The fix is simpler than what it replaces: `parseStatement()` already dispatches an `If` token to
the real, recursive if-parser, which correctly consumes the nested `If`'s entire structure — its
own `Else`/`ElseIf` chain and matching `End If` — before returning. By the time control comes back
to the outer loop, any nested `If` has already fully consumed its own terminator, so an
`Else`/`ElseIf`/`End If` seen at *this* loop level can only belong to the block currently being
parsed. No lookahead or depth tracking needed at all.

### Statement-call with a parenthesized first argument, followed by more arguments (`expression-parser.ts`)

```vbscript
PlaySound("fx_ballrolling" & b), -1, Vol(BOT(b)), Pan(BOT(b)), 0, ballpitch, 1, 0, AudioFade(BOT(b))
```

A real line from a production table's script. `parseCall()` correctly builds `PlaySound(...)` as a
single-argument `CallExpression` (arrays and calls share identical syntax, so it can't know ahead
of time this is a bare statement call whose first argument merely *happens* to have its own
parens). The following `Comma` was only ever handled as "maybe an assignment (`=`)" or "abandon
and reparse as a binary-operator continuation" — never as "this is actually more arguments to the
same call." Fixed by recognizing a `Comma` immediately after a single-arg `CallExpression` (whose
callee isn't itself another call, to avoid misreading a jagged-array read) as the start of more
statement-call arguments, and combining them into one `CallExpression`.

### Omitted leading statement-call argument (`expression-parser.ts`)

```vbscript
PlayersReel.SetValue, PlayersPlayingGame
```

Also from a real table script. VBScript lets you skip a positional argument with a bare comma —
already supported for parenthesized calls (`Foo(, x)`, via a `VbEmptyLiteral` placeholder in
`parseArguments()`) but never wired up for the no-parens statement-call path, where a leading
`Comma` failed `isStatementCallArgumentStart()` and the whole statement fell through to "not a
call, must be a bare read," leaving the comma and second argument dangling. Fixed by accepting a
leading `Comma` as a valid statement-call start and reusing the same `VbEmptyLiteral` placeholder
in `parseStatementCallArguments()`.

### Type-annotation keywords usable as ordinary identifiers (`expression-parser.ts`, `procedures.ts`)

```vbscript
Function GetHSChar(String, Index)
    ThisChar = Mid(String, Index, 1)
```

`String`/`Integer`/`Long`/`Boolean`/`Date`/`Object`/`Variant`/etc. are reserved here only for this
engine's `Dim x As String` extension (real VBScript has no reserved type names at all), but
shadowing a built-in function name as a variable or parameter is completely normal and common.
Fixed by adding a `parseFlexibleIdentifier()` (reusing the already-existing permissive
`parsePropertyName()` logic) at all three places that previously required the strict `Identifier`
token type: parameter declarations, general identifier references, and `parsePrimary()`'s
dispatch.

### Empty-parens array declarations (`declarations.ts`)

```vbscript
Dim SomeArray()
```

`Dim`/`ReDim` with no bounds yet (to be supplied later via a real `ReDim`) is valid VBScript, but
`parseArrayBounds()` was called unconditionally and threw `Unexpected token: RParen` trying to
parse an expression starting at `)`. Fixed by only calling it when an expression is actually
present.

### Chained/nested array reads (`expression-evaluator.ts`)

```vbscript
obj.Items(0)
arr(0)(4)
```

Only the plain-`Identifier`-callee case handled a call result that turned out to be an array —
`MemberExpression`/`CallExpression` callees threw instead of resolving.

### Hex/octal integer literal semantics (`lexer.ts`)

Verified against Wine's own `vbscript.dll` test suite:

- **Narrowest-width-then-sign-extend**: an unsuffixed hex/octal literal is interpreted using the
  narrowest width its raw bit pattern fits (16-bit if ≤ `0xFFFF`, else 32-bit), then
  reinterpreting the top bit as a sign via two's complement — `&hffff` is `-1`, not `65535`. A
  trailing `&` suffix forces 32-bit (Long) width regardless — `&hffff&` stays `65535`. The
  trailing `&` wasn't even being consumed before, so it got re-tokenized as a stray `Ampersand`
  (string-concat) operator, cascading into unrelated parse errors.
- **Bare `&` + octal digits with no `o`/`O` letter** (`&100` = octal 100 = decimal `64`) was
  entirely unhandled — Wine's suite explicitly documents this form.

### Trailing-decimal-point float literals (`lexer.ts`)

```vbscript
10. = 10.0
```

A digit was already consumed before the `.`, so it's unambiguously part of the number (numbers are
never valid targets of `.property` access) — no digit needs to follow it. Previously required a
digit after the dot, so `10.` left the `.` unconsumed and re-tokenized as a stray member-access
operator.

### Sub/Function declarations aren't hoisted (`interpreter.ts`)

```vbscript
Option Explicit
LoadCoreFiles          ' called here...

Sub LoadCoreFiles      ' ...but not declared until here
    ExecuteGlobal GetTextFile("core.vbs")
End Sub
```

Real VBScript hoists every `Sub`/`Function` declaration in a scope before running
any statement in that scope, so a call can textually precede its own declaration —
extremely common in real table scripts, which often open with a bare call to a
loader routine defined further down the file (this exact pattern, calling a
`LoadCoreFiles`-style Sub to dynamically load `core.vbs`/`controller.vbs`, is a
near-universal convention across VPX tables). This engine's interpreter only
registered a `Sub`/`Function` into its function registry when execution reached it
sequentially, so a forward call threw `Variable is undefined` under `Option
Explicit` — or, without it, silently no-op'd instead of calling the Sub at all,
since an unrecognized bare identifier is otherwise just an implicit variable read.
Fixed with a small pre-pass (`hoistDeclarations()`) that registers every
top-level `Sub`/`Function` before a program's statements start executing —
applied both to whole-program execution and to `ExecuteGlobal`'s dynamic-eval
path, since real tables commonly hit this same ordering inside dynamically
loaded code, not just the top-level table script.

Found the same way as the other real-table-script fixes above — reproduced with
a minimal case first, then confirmed it accounts for the exact
`LoadCoreFiles`/`core.vbs` failure seen running a real production table end to
end.

## Testing

- All fixes verified with minimal reproductions before/after.
- Cross-checked against Wine's real `vbscript.dll` conformance suite where applicable (the
  lexer-level fixes and several parser fixes).
- The full set verified together against a real, complete production VPX table script (not
  included here — third-party content) that exercised all of the statement-call/identifier fixes
  in combination, not just in isolation.
- Existing test suite passes with no regressions.
