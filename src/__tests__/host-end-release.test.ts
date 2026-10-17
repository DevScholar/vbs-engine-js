import { describe, expect, it } from "vitest";
import { VbsEngine } from "../core/index.ts";

/**
 * The engine that claims host integration stands for the host, so it lets the script go when that host ends, the
 * way `cscript` runs the Terminate event of an instance the script never dropped as it runs out of script. The
 * host's end is the whole process here, and an `exit` listener runs on `process.emit('exit')` exactly as it does
 * on the real thing, so the test can ask for the end and then carry on.
 *
 * Only one engine per process claims host integration, which is what the first engine of this file gets. Every
 * later engine is isolated, which is what a library the host drives is: it has no end of its own.
 */
const SCRIPT = `Dim log
log = ""
Class Counter
  Private Sub Class_Terminate()
    log = log & "terminate;"
  End Sub
End Class
Set held = New Counter
`;

describe("the end of the host", () => {
  it("lets the script go when the host ends, and stops listening once destroyed", () => {
    const before = process.listenerCount("exit");
    const engine = new VbsEngine();
    expect(process.listenerCount("exit")).toBe(before + 1);

    engine.executeStatement(SCRIPT);
    expect(engine.eval("log")).toBe("");

    process.emit("exit", 0);

    expect(engine.eval("log")).toBe("terminate;");

    // The name is still there and still answers the instance, which is now terminated, and not Nothing: a host
    // that releases the globals and then carries on can read them. Letting the same value go again runs nothing.
    expect(engine.eval("TypeName(held)")).toBe("Counter");
    expect(engine.eval("held Is Nothing")).toBe(false);
    process.emit("exit", 0);
    expect(engine.eval("log")).toBe("terminate;");

    engine.destroy();
    expect(process.listenerCount("exit")).toBe(before);
  });

  it("leaves an isolated engine with no end of its own", () => {
    const before = process.listenerCount("exit");
    const isolated = new VbsEngine();
    expect(process.listenerCount("exit")).toBe(before);

    isolated.executeStatement(SCRIPT);
    process.emit("exit", 0);

    // Nothing released it, so the instance the script still holds has not had its Terminate event run.
    expect(isolated.eval("log")).toBe("");
  });
});
