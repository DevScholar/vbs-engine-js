import { describe, expect, it } from 'vitest';
import { VbsEngine } from '../core/index.ts';

/**
 * Class_Initialize and Class_Terminate, one measured cell at a time.
 *
 * Every row was read off a fresh Windows cscript on its own, since an instance that outlives a cell is still
 * alive when the next cell starts, so one process per cell is what keeps each trace the cell's own. The probe,
 * its capture and the comparison script live outside this repository, in the temporary probe directory; this
 * file is the durable copy of what they measured.
 *
 * A cell holds the statements it runs, the trace its Initialize and Terminate events made, and the answer it
 * left in evtResult or the error number it ended with.
 */
const CLASSES = `Class Probe
  Public Name
  Private storedLinked
  Private storedMade

  Private Sub Class_Initialize()
    evtTrace = evtTrace & "init;"
  End Sub

  Private Sub Class_Terminate()
    evtTrace = evtTrace & "term(" & Name & ");"
  End Sub

  Public Property Set Linked(other)
    Set storedLinked = other
  End Property

  Public Property Get Linked
    Set Linked = storedLinked
  End Property

  Public Function SelfValue()
    Set SelfValue = Me
  End Function

  Public Sub CallInitialize()
    Class_Initialize
  End Sub

  Public Sub CallTerminate()
    Class_Terminate
  End Sub
End Class

Class Quiet
  Public Name
End Class

Class Failing
  Private Sub Class_Initialize()
    evtTrace = evtTrace & "init;"
    Err.Raise 5
  End Sub
  Private Sub Class_Terminate()
    evtTrace = evtTrace & "term;"
  End Sub
End Class

Class Raising
  Private Sub Class_Terminate()
    evtTrace = evtTrace & "term;"
    Err.Raise 5
  End Sub
End Class

Class CtorTaking
  Public Name
  Private Sub Class_Initialize()
    evtTrace = evtTrace & "init;"
  End Sub
  Private Sub Class_Terminate()
    evtTrace = evtTrace & "term;"
  End Sub
End Class

Class PlainInit
  Public Name
  Public Sub Class_Initialize()
    evtTrace = evtTrace & "init;"
  End Sub
  Public Sub Class_Terminate()
    evtTrace = evtTrace & "term;"
  End Sub
End Class

Function FreshMaker()
  Dim fresh, made
  Set fresh = New Probe
  fresh.Name = "fresh"
  Set made = fresh
  Set FreshMaker = made
End Function

Sub DropLater()
  Dim later
  Set later = New Probe
  later.Name = "later"
End Sub

Sub PlaceInArray(list)
  Dim placed
  Set placed = New Probe
  placed.Name = "placed"
  Set list(0) = placed
End Sub

Sub ReplaceLocal()
  Dim inner
  Set inner = New Probe
  inner.Name = "inner"
  Set inner = New Probe
  inner.Name = "second"
End Sub

Sub CreateFresh()
  Dim fresh
  Set fresh = New Probe
  fresh.Name = "fresh"
End Sub

Sub HoldOnlyLocal()
  Dim local
  Set local = New Probe
  local.Name = "local"
End Sub

Sub ByRefReplace(item)
  Set item = New Probe
  item.Name = "newer"
End Sub

Sub TakeByVal(ByVal item)
  item.Name = "valued"
End Sub

Sub TakeByRef(item)
  item.Name = "byref"
End Sub

Sub ByValReplace(ByVal item)
  Set item = New Probe
  item.Name = "newer"
End Sub

Sub ByRefFresh(item)
  Set item = New Probe
  item.Name = "newer"
End Sub

Sub ByRefKept(item)
  item.Name = "kept"
End Sub

Sub ByValFresh(ByVal item)
  Set item = New Probe
  item.Name = "newer"
End Sub
`;

const CELLS: ReadonlyArray<readonly [string, string, string, string]> = [
  ['create.simple', 'Set probeOne = New Probe\nprobeOne.Name = "one"\nevtResult = ""', 'init;', ''],
  [
    'create.thenNew',
    'Set probeOne = New Probe\nprobeOne.Name = "first"\nSet probeOne = New Probe\nprobeOne.Name = "second"\nevtResult = ""',
    'init;init;term(first);',
    '',
  ],
  [
    'create.newObjectBeforeOldIsReleased',
    'Set probeOne = New Probe\nprobeOne.Name = "one"\nSet probeTwo = probeOne\nSet probeTwo = New Probe\nprobeTwo.Name = "two"\nevtResult = ""',
    'init;init;',
    '',
  ],
  [
    'terminate.class_TerminateAtNothing',
    'Set probeOne = New Probe\nprobeOne.Name = "gone"\nSet probeOne = Nothing\nevtResult = CStr(probeOne Is Nothing)',
    'init;term(gone);',
    'True',
  ],
  [
    'terminate.twoNames',
    'Set probeOne = New Probe\nprobeOne.Name = "both"\nSet probeTwo = probeOne\nSet probeOne = Nothing\nevtResult = probeTwo.Name',
    'init;',
    'both',
  ],
  [
    'terminate.lastReferenceGoes',
    'Set probeOne = New Probe\nprobeOne.Name = "last"\nSet probeTwo = probeOne\nSet probeOne = Nothing\nSet probeTwo = Nothing\nevtResult = ""',
    'init;term(last);',
    '',
  ],
  [
    'terminate.selfAssign',
    'Set probeOne = New Probe\nprobeOne.Name = "same"\nSet probeOne = probeOne\nevtResult = ""',
    'init;',
    '',
  ],
  [
    'terminate.classWithoutTerminate',
    'Set quietOne = New Quiet\nSet quietOne = Nothing\nevtResult = TypeName(quietOne)',
    '',
    'Nothing',
  ],
  [
    'terminate.classWithoutInitialize',
    'Set probeOne = New Quiet\nprobeOne.Name = "quiet"\nSet probeOne = Nothing\nevtResult = "ran"',
    '',
    'ran',
  ],
  [
    'terminate.atScriptEnd',
    'Set probeOne = New Probe\nprobeOne.Name = "alive"\nevtResult = ""',
    'init;',
    '',
  ],
  ['initialize.countOnce', 'Set probeOne = New Probe\nevtResult = "made"', 'init;', 'made'],
  [
    'initialize.meIsTheInstance',
    'Set probeOne = New Probe\nprobeOne.Name = "self"\nevtResult = CStr(probeOne.SelfValue() Is probeOne)',
    'init;',
    'True',
  ],
  [
    'initialize.nothingReturns',
    'Set probeOne = New Probe\nprobeOne.Name = "returns"\nSet probeTwo = probeOne\nSet probeOne = Nothing\nevtResult = probeTwo.Name',
    'init;',
    'returns',
  ],
  [
    'terminate.argumentStaysAlive',
    'Set probeOne = New Probe\nprobeOne.Name = "arg"\nTakeByRef probeOne\nevtResult = probeOne.Name',
    'init;',
    'byref',
  ],
  [
    'terminate.byValArgument',
    'Set probeOne = New Probe\nprobeOne.Name = "val"\nTakeByVal probeOne\nevtResult = probeOne.Name',
    'init;',
    'valued',
  ],
  [
    'terminate.byRefReplace',
    'Set probeOne = New Probe\nprobeOne.Name = "old"\nByRefReplace probeOne\nevtResult = probeOne.Name',
    'init;init;term(old);',
    'newer',
  ],
  ['terminate.localAtSubEnd', 'HoldOnlyLocal\nevtResult = "back"', 'init;term(local);', 'back'],
  [
    'terminate.localReplacedInSub',
    'ReplaceLocal\nevtResult = "back"',
    'init;init;term(inner);term(second);',
    'back',
  ],
  [
    'terminate.returnedToCaller',
    'Set probeOne = FreshMaker()\nprobeOne.Name = "kept"\nSet probeOne = Nothing\nevtResult = "back"',
    'init;term(kept);',
    'back',
  ],
  ['terminate.returnedAndDropped', 'FreshMaker\nevtResult = "back"', 'init;term(fresh);', 'back'],
  [
    'terminate.byRefParameterReplaced',
    'ByRefReplace probeOne\nevtResult = TypeName(probeOne)',
    'init;',
    'Probe',
  ],
  [
    'terminate.byRefParameterHoldsFresh',
    'Set probeOne = New Probe\nprobeOne.Name = "old"\nByRefReplace probeOne\nevtResult = probeOne.Name',
    'init;init;term(old);',
    'newer',
  ],
  [
    'terminate.byValParameterReplaced',
    'Set probeOne = New Probe\nprobeOne.Name = "old"\nByValReplace probeOne\nevtResult = probeOne.Name',
    'init;init;term(newer);',
    'old',
  ],
  [
    'terminate.byValParameterHoldsFresh',
    'ByValReplace probeOne\nevtResult = "back"',
    'init;term(newer);',
    'back',
  ],
  [
    'terminate.byRefParameterFresh',
    'ByRefFresh probeOne\nevtResult = TypeName(probeOne)',
    'init;',
    'Probe',
  ],
  [
    'terminate.byRefParameterKept',
    'Set probeOne = New Probe\nprobeOne.Name = "old"\nByRefKept probeOne\nevtResult = "back"',
    'init;',
    'back',
  ],
  [
    'terminate.byValParameterFresh',
    'ByValFresh probeOne\nevtResult = "back"',
    'init;term(newer);',
    'back',
  ],
  [
    'terminate.nothingIsNoCall',
    'Set probeOne = Nothing\nSet probeOne = Nothing\nevtResult = CStr(probeOne Is Nothing)',
    '',
    'True',
  ],
  [
    'terminate.arrayElementReplaced',
    'Set slotsA(0) = New Probe\nslotsA(0).Name = "slot"\nSet slotsA(0) = New Probe\nslotsA(0).Name = "next"\nSet slotsA(0) = Nothing\nevtResult = ""',
    'init;init;term(slot);term(next);',
    '',
  ],
  [
    'terminate.eraseArray',
    'Set slotsB(0) = New Probe\nslotsB(0).Name = "erased"\nErase slotsB\nevtResult = ""',
    'init;term(erased);',
    '',
  ],
  [
    'terminate.arrayAnswerReassigned',
    'Set slotsC(0) = New Probe\nslotsC(0).Name = "inArray"\nslotsC = Array(1, 2)\nevtResult = CStr(slotsC(0))',
    'init;',
    'ERR 438',
  ],
  [
    'terminate.arrayHoldsAlive',
    'PlaceInArray slotsD\nSet slotsD(0) = Nothing\nevtResult = "back"',
    'init;term(placed);',
    'back',
  ],
  [
    'terminate.arrayCopy',
    'ReDim secondSlots(1)\nSet firstSlots(0) = New Probe\nfirstSlots(0).Name = "copy"\nsecondSlots = firstSlots\nSet firstSlots(0) = Nothing\nSet secondSlots(0) = Nothing\nevtResult = ""',
    'init;term(copy);',
    'ERR 13',
  ],
  [
    'terminate.reDim',
    'ReDim slotsE(1)\nSet slotsE(0) = New Probe\nslotsE(0).Name = "redim"\nReDim slotsE(3)\nevtResult = ""',
    'init;term(redim);',
    '',
  ],
  [
    'terminate.objectIntoArrayElement',
    'Set probeOne = New Probe\nprobeOne.Name = "held"\nSet slotsF(0) = probeOne\nSet probeOne = Nothing\nevtResult = slotsF(0).Name',
    'init;',
    'held',
  ],
  [
    'initialize.failureRunsTerminate',
    'Set failingOne = New Failing\nevtResult = CStr(Err.Number)',
    'init;',
    'ERR 5',
  ],
  [
    'terminate.errorInTerminate',
    'Set raisingOne = New Raising\nraisingOne.Name = "boom"\nSet raisingOne = Nothing\nevtResult = CStr(Err.Number)',
    'term;',
    'ERR 5',
  ],
  [
    'terminate.errorInTerminateDoesNotStopNextCell',
    'Set raisingOne = New Raising\nraisingOne.Name = "boom"\nSet raisingOne = Nothing\nErr.Clear\nSet probeOne = New Probe\nprobeOne.Name = "fine"\nSet probeOne = Nothing\nIf Err.Number = 0 Then evtResult = "fine"',
    'term;init;term(fine);',
    'fine',
  ],
  [
    'initialize.namedFromOutside',
    'Set probeOne = New Probe\nprobeOne.Name = "meth"\nprobeOne.Class_Initialize\nevtResult = "ran"',
    'init;',
    'ERR 438',
  ],
  [
    'initialize.namedFromOutsideWithParens',
    'Set probeOne = New Probe\nprobeOne.Name = "meth"\nprobeOne.Class_Initialize()\nevtResult = "ran"',
    'init;',
    'ERR 438',
  ],
  [
    'terminate.namedFromOutside',
    'Set probeOne = New Probe\nprobeOne.Name = "meth"\nprobeOne.Class_Terminate\nevtResult = "ran"',
    'init;',
    'ERR 438',
  ],
  [
    'initialize.namedInside',
    'Set probeOne = New Probe\nprobeOne.Name = "inside"\nprobeOne.CallInitialize\nevtResult = "ran"',
    'init;init;',
    'ran',
  ],
  [
    'terminate.namedInside',
    'Set probeOne = New Probe\nprobeOne.Name = "inside"\nprobeOne.CallTerminate\nSet probeOne = Nothing\nevtResult = "ran"',
    'init;term(inside);term(inside);',
    'ran',
  ],
  [
    'terminate.propertySetKeepsArgument',
    'Set probeOne = New Probe\nprobeOne.Name = "prop"\nSet probeOne.Linked = New Probe\nevtResult = TypeName(probeOne.Linked)',
    'init;init;',
    'Probe',
  ],
  [
    'terminate.propertySetArgument',
    'Set probeOne = New Probe\nprobeOne.Name = "prop"\nSet probeTwo = New Probe\nprobeTwo.Name = "linked"\nSet probeOne.Linked = probeTwo\nSet probeOne = Nothing\nSet probeTwo = Nothing\nevtResult = ""',
    'init;init;term(prop);term(linked);',
    '',
  ],
  [
    'terminate.returnOfFunctionKept',
    'Set probeOne = New Probe\nprobeOne.Name = "src"\nSet probeTwo = probeOne.SelfValue()\nSet probeOne = Nothing\nevtResult = probeTwo.Name',
    'init;',
    'src',
  ],
  [
    'terminate.temporaryIsNotAHolder',
    'Set probeOne = New Probe\nprobeOne.Name = "temp"\nSet probeZero = New Probe\nprobeZero.Name = "zero"\nSet probeOne = probeOne.SelfValue()\nSet probeOne = Nothing\nevtResult = "back"',
    'init;init;term(temp);',
    'back',
  ],
  [
    'initialize.ctorNoArguments',
    'Set ctorOne = New CtorTaking\nctorOne.Name = "plain"\nSet ctorOne = Nothing\nevtResult = "back"',
    'init;term;',
    'back',
  ],
  [
    'events.publicAccess',
    'Set plainOne = New PlainInit\nplainOne.Name = "public"\nSet plainOne = Nothing\nevtResult = "back"',
    'init;term;',
    'back',
  ],
  [
    'terminate.insideSubGlobalDrop',
    'Set probeOne = New Probe\nprobeOne.Name = "global"\nDropLater\nSet probeOne = Nothing\nevtResult = ""',
    'init;init;term(later);term(global);',
    '',
  ],
  [
    'terminate.variableReassignedToEmpty',
    'Set probeOne = New Probe\nprobeOne.Name = "emp"\nprobeOne = Empty\nevtResult = TypeName(probeOne)',
    'init;term(emp);',
    'Empty',
  ],
  [
    'terminate.variableReassignedByNew',
    'Set probeOne = New Probe\nprobeOne.Name = "one"\nSet probeOne = New Probe\nprobeOne.Name = "two"\nSet probeOne = Nothing\nevtResult = ""',
    'init;init;term(one);term(two);',
    '',
  ],
  [
    'terminate.secondNameDropped',
    'Set probeOne = New Probe\nprobeOne.Name = "one"\nSet probeTwo = New Probe\nprobeTwo.Name = "two"\nSet probeTwo = probeOne\nSet probeOne = Nothing\nevtResult = probeTwo.Name',
    'init;init;term(two);',
    'one',
  ],
  [
    'terminate.insideWith',
    'Set probeOne = New Probe\nprobeOne.Name = "with"\nWith probeOne\n  .Name = "with2"\nEnd With\nSet probeOne = Nothing\nevtResult = ""',
    'init;term(with2);',
    '',
  ],
  [
    'terminate.insideForEach',
    'ReDim slotsG(1)\nSet slotsG(0) = New Probe\nslotsG(0).Name = "loop0"\nSet slotsG(1) = New Probe\nslotsG(1).Name = "loop1"\nFor Each heldLoop In slotsG\n  heldLoop.Name = heldLoop.Name & "!"\nNext\nSet slotsG(0) = Nothing\nSet slotsG(1) = Nothing\nevtResult = ""',
    'init;init;term(loop0!);term(loop1!);',
    '',
  ],
  [
    'terminate.nestedNewAsArgument',
    'Set probeOne = New Probe\nprobeOne.Name = "outer"\nSet probeOne.Linked = New Probe\nevtResult = TypeName(probeOne.Linked)',
    'init;init;',
    'Probe',
  ],
  [
    'terminate.notCalledForQuietAtEnd',
    'Set quietOne = New Quiet\nquietOne.Name = "quiet"\nevtResult = quietOne.Name',
    '',
    'quiet',
  ],
  [
    'terminate.sameNameRecreated',
    'Set probeOne = New Probe\nprobeOne.Name = "re"\nSet probeOne = New Probe\nprobeOne.Name = "re"\nSet probeOne = Nothing\nevtResult = ""',
    'init;init;term(re);term(re);',
    '',
  ],
  ['terminate.localAtFunctionEnd', 'CreateFresh\nevtResult = "back"', 'init;term(fresh);', 'back'],
];

/** Runs every cell in an engine of its own, the way the probe runs every cell in a cscript of its own. */
function runCells(): string[] {
  return CELLS.map(([, code]) => {
    const engine = new VbsEngine();
    engine.executeStatement(`${CLASSES}
Dim evtTrace, evtResult, evtNumber, evtBeforeEnd
Dim slotsA(1), slotsB(1), slotsC(1), slotsD(1), slotsE(), slotsF(1), slotsG()
Dim firstSlots(1), secondSlots()
Dim probeOne, probeTwo, probeZero, failingOne, raisingOne, ctorOne, plainOne, quietOne
Dim heldLoop
On Error Resume Next
evtTrace = ""
evtResult = ""
evtNumber = 0
`);
    engine.executeStatement(`${code}
evtNumber = Err.Number`);
    const trace = String(engine._getVariable('evtTrace').value);
    const answer = String(engine._getVariable('evtResult').value);
    const failure = engine.error;
    const number =
      failure && failure.number !== 0
        ? String(failure.number)
        : String(engine._getVariable('evtNumber').value);
    // The probe ends its script here, which is what runs the Terminate event of an instance the cell never
    // dropped. It runs after the trace above was read, so that trace is what the cell's own statements made.
    engine.releaseGlobals();
    return number !== '0' ? `${trace}|ERR ${number}` : `${trace}|${answer}`;
  });
}

describe('Class_Initialize and Class_Terminate behave like vbscript.dll', () => {
  const answers = runCells();
  CELLS.forEach(([label, , trace, answer], index) => {
    it(`${label} leaves ${trace}|${answer}`, () => {
      expect(answers[index]).toBe(`${trace}|${answer}`);
    });
  });
});
