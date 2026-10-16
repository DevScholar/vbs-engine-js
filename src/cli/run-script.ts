#!/usr/bin/env node
/**
 * run-script — a cscript/wscript-style command-line host for VBScript (and, in a
 * future phase, JScript) built on vbs-engine-js.
 *
 * Usage:
 *   run-script script.vbs [//options] [script-args...]
 *   run-script script.wsf //job:Job1 [//options] [script-args...]
 *
 * Options (WSH-compatible where possible):
 *   //nologo               suppress the logo/banner line
 *   //job:<id>             select a <job> in a .wsf
 *   //B                    batch mode (suppress errors/prompts to a degree)
 *   //I                    interactive mode (default)
 *   //E:<engine>           target engine (vbscript only for now)
 *   --gui-mode=cscript|wscript   controls WScript.Echo (console vs GUI)
 *
 * Named script arguments use WSH's `/name:value` form; unnamed are bare words.
 * `WScript.Quit n` terminates the script and sets the process exit code to n.
 */
import * as fs from 'node:fs';
import * as path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { VbsEngine, type VbsError } from '../core/index.ts';
import { createWScript, WScriptQuitSignal, type WshEnvironment } from '../host/wscript.ts';
import { findBlockStart, parseWsf, type WsfJob } from '../host/wsf.ts';
import { VBS_COMPILE_SOURCE } from '../runtime/errors.ts';
import { formatWshErrorLine } from '../runtime/wsh-report.ts';
import { runJScript } from '../host/jscript.ts';
import { createMsgBox, getDefaultResult } from '../builtins/msgbox.ts';
import { createInputBox } from '../builtins/inputbox.ts';

interface CliOptions {
  file: string;
  jobId?: string;
  nologo: boolean;
  batch: boolean;
  guiMode: 'cscript' | 'wscript';
  unnamedArgs: string[];
  namedArgs: Record<string, string>;
}

const __dirname = path.dirname(fileURLToPath(import.meta.url));

/** Synchronous sleep that blocks the thread (no event-loop release). */
function syncSleep(ms: number): void {
  const sab = new Int32Array(new SharedArrayBuffer(4));
  Atomics.wait(sab, 0, 0, Math.max(0, ms));
}

/**
 * Returns a blocking GUI dialog backend for `wscript`-mode Echo, selected by
 * platform: Windows → PowerShell MessageBox, macOS → osascript, any other
 * (Unix-like) environment → zenity. No availability probing: the choice is made
 * by `process.platform` alone, and a missing backend surfaces its own error at
 * call time (mirroring real wscript, which does not silently degrade).
 */
function resolveGuiDialog(): (message: string, title?: string) => void {
  if (process.platform === 'win32') {
    return (message: string, title?: string): void => {
      const ps =
        `Add-Type -AssemblyName System.Windows.Forms; ` +
        `[void][System.Windows.Forms.MessageBox]::Show(${JSON.stringify(message)}, ${JSON.stringify(title ?? 'Windows Script Host')})`;
      execFileSync('powershell', ['-NoProfile', '-Command', ps], { stdio: 'inherit' });
    };
  }

  if (process.platform === 'darwin') {
    return (message: string, title?: string): void => {
      const alert = title
        ? `display alert ${JSON.stringify(message)} message ${JSON.stringify(title)}`
        : `display alert ${JSON.stringify(message)}`;
      execFileSync('osascript', ['-e', alert], { stdio: 'inherit' });
    };
  }

  // Unix-like, non-macOS.
  return (message: string, title?: string): void => {
    execFileSync(
      'zenity',
      ['--info', '--title', title ?? 'Windows Script Host', '--text', message],
      {
        stdio: 'inherit',
      }
    );
  };
}

/** Synchronous line read from fd 0; returns null on EOF. */
function syncReadLine(): string | null {
  try {
    const chunks: Buffer[] = [];
    for (;;) {
      const buffer = Buffer.alloc(1024);
      let bytesRead: number;
      try {
        bytesRead = fs.readSync(0, buffer, 0, buffer.length, null);
      } catch {
        return chunks.length
          ? Buffer.concat(chunks)
              .toString('utf8')
              .replace(/\r?\n$/, '')
          : null;
      }
      if (bytesRead === 0) {
        break;
      }
      const chunk = buffer.subarray(0, bytesRead);
      chunks.push(chunk);
      if (chunk.includes(10)) {
        break;
      }
    }
    if (chunks.length === 0) return null;
    return Buffer.concat(chunks)
      .toString('utf8')
      .replace(/\r?\n$/, '');
  } catch {
    return null;
  }
}

function parseArgs(argv: string[]): CliOptions {
  const opts: CliOptions = {
    file: '',
    nologo: false,
    batch: false,
    guiMode: 'cscript',
    unnamedArgs: [],
    namedArgs: {},
  };

  let seenFile = false;
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i]!;

    if (!seenFile) {
      if (arg === '//nologo') {
        opts.nologo = true;
        continue;
      }
      if (arg === '//B') {
        opts.batch = true;
        continue;
      }
      if (arg === '//I') {
        opts.batch = false;
        continue;
      }
      if (arg.startsWith('//job:')) {
        opts.jobId = arg.slice('//job:'.length);
        continue;
      }
      if (arg.startsWith('//E:')) {
        // Only VBScript is supported; ignore for now (validate loosely).
        continue;
      }
      if (arg.startsWith('--gui-mode=')) {
        const m = arg.slice('--gui-mode='.length);
        opts.guiMode = m === 'wscript' ? 'wscript' : 'cscript';
        continue;
      }
      if (arg.startsWith('//')) {
        // Unknown //switch: treat as an error rather than silently passing it as a
        // filename, matching cscript's "Unknown switch" behavior.
        console.error(`run-script: unknown switch '${arg}'`);
        process.exit(1);
      }
      opts.file = arg;
      seenFile = true;
      continue;
    }

    // After the script file: named (/name:value) vs unnamed args.
    if (arg.startsWith('/') && arg.includes(':') && arg.length > 2) {
      const colon = arg.indexOf(':');
      const name = arg.slice(1, colon).toLowerCase();
      opts.namedArgs[name] = arg.slice(colon + 1);
    } else {
      opts.unnamedArgs.push(arg);
    }
  }

  if (!opts.file) {
    console.error('Usage: run-script script.vbs [//options] [script-args...]');
    process.exit(1);
  }
  return opts;
}

/** Builds the usage string for a .wsf job from its <runtime> declarations. */
function buildUsage(fileName: string, job: WsfJob): string {
  const lines: string[] = [];
  if (job.description) lines.push(job.description);
  lines.push(`Usage: ${fileName}`);
  for (const p of job.named) {
    const req = p.required ? '' : ' [';
    const close = p.required ? '' : ']';
    lines.push(`  /${p.name}:value${req}${p.helpString ? '  ' + p.helpString : ''}${close}`);
  }
  for (const p of job.unnamed) {
    lines.push(`  <${p.name}>${p.helpString ? '  ' + p.helpString : ''}`);
  }
  if (job.example) lines.push(job.example);
  return lines.join('\n') + '\n';
}

/**
 * Windows Script Host reports a block taken out of a `.wsf` one column further along than the same code in a
 * `.vbs`: measured on four blocks whose statements stood at columns 1, 5, 1 and 9, and came back as 2, 6, 2 and
 * 10. The same adjustment covers a script a job points at with `src`, which was reported at column 2 for a
 * statement standing at column 1 of the file it came from.
 */
const WSF_COLUMN_OFFSET = 1;

function readScriptFile(filePath: string): { kind: 'vbs' | 'wsf'; source: string; abs: string } {
  const abs = path.resolve(filePath);
  if (!fs.existsSync(abs)) {
    console.error(`run-script: file not found: ${filePath}`);
    process.exit(1);
  }
  const source = fs.readFileSync(abs, 'utf8');
  const kind = abs.toLowerCase().endsWith('.wsf') ? 'wsf' : 'vbs';
  return { kind, source, abs };
}

/**
 * Extract the script blocks to run (in document order), given the parsed options.
 * For a `.wsf` this selects the requested job and returns each of its script
 * blocks with its language, so VBScript and JScript can be executed in order with
 * a shared global scope. A plain `.vbs` produces a single VBScript block.
 *
 * Each block carries where it stands in the file, which the engine is told so that a failure is reported at the
 * line Windows Script Host would name. For a block inside the `.wsf` that is the block's own line in the file, and
 * for a script the job points at with `src` it is a line in that other file, which is what `cscript` reported
 * under the `.wsf`'s own name.
 */
interface ScriptBlock {
  language: string;
  code: string;
  /** The file the failure is reported under, which is the `.wsf` or `.vbs` this host was pointed at. */
  name: string;
  /** The line in that file the code begins on, counted from zero the way the engine takes it. */
  startingLineNumber?: number;
  /** What this host adds to every column it reports, which a `.wsf` does and a `.vbs` does not. */
  columnOffset?: number;
}

function resolveSource(cli: CliOptions): {
  blocks: ScriptBlock[];
  job?: WsfJob;
  kind: 'vbs' | 'wsf';
} {
  const { kind, source, abs } = readScriptFile(cli.file);
  if (kind === 'vbs') {
    return { blocks: [{ language: 'VBScript', code: source, name: abs }], kind };
  }

  const doc = parseWsf(source);
  const job = cli.jobId
    ? doc.jobs.find(j => j.id?.toLowerCase() === cli.jobId!.toLowerCase())
    : doc.jobs[0];

  if (!job) {
    console.error(`run-script: job '${cli.jobId}' not found`);
    process.exit(1);
  }

  // Blocks are found in the file one after another, so each search begins where the last one ended.
  let cursor = 0;
  const blocks: ScriptBlock[] = job.scripts.map(s => {
    if (s.code !== undefined) {
      const found = findBlockStart(source, s.code, cursor);
      if (found) cursor = found.next;
      return {
        language: s.language,
        code: s.code,
        name: abs,
        ...(found ? { startingLineNumber: found.line - 1 } : {}),
        columnOffset: WSF_COLUMN_OFFSET,
      };
    }
    if (s.src) {
      const srcPath = path.resolve(path.dirname(abs), s.src);
      if (!fs.existsSync(srcPath)) {
        console.error(`run-script: script src not found: ${s.src}`);
        process.exit(1);
      }
      return {
        language: s.language,
        code: fs.readFileSync(srcPath, 'utf8'),
        name: abs,
        columnOffset: WSF_COLUMN_OFFSET,
      };
    }
    return { language: s.language, code: '', name: abs, columnOffset: WSF_COLUMN_OFFSET };
  });

  return { blocks, job, kind };
}

/**
 * Prints a failure the way Windows Script Host prints it and answers the code the process should leave with.
 * Measured with `cscript`: a failure inside a `.wsf` leaves 1, a `.vbs` leaves 1 for a parse failure, and a `.vbs`
 * that failed while running leaves 0.
 */
function reportFailure(failure: VbsError, scriptName: string, kind: 'vbs' | 'wsf'): never {
  console.error(formatWshErrorLine(failure, scriptName));
  process.exit(kind === 'wsf' ? 1 : failure.source === VBS_COMPILE_SOURCE ? 1 : 0);
}

function buildEnvironment(cli: CliOptions, job: WsfJob | undefined): WshEnvironment {
  const abs = path.resolve(cli.file);
  const scriptName = path.basename(cli.file);
  const usage = job ? buildUsage(scriptName, job) : '';

  return {
    mode: cli.guiMode,
    scriptFullName: abs,
    scriptName,
    interactive: !cli.batch,
    unnamedArgs: cli.unnamedArgs,
    namedArgs: cli.namedArgs,
    usage,
    writeStdout: (t: string) => {
      try {
        fs.writeSync(1, t);
      } catch {
        process.stdout.write(t);
      }
    },
    writeStderr: (t: string) => {
      try {
        fs.writeSync(2, t);
      } catch {
        process.stderr.write(t);
      }
    },
    readLine: syncReadLine,
    sleep: syncSleep,
    showDialog: cli.guiMode === 'wscript' ? resolveGuiDialog() : undefined,
    fullName: fileURLToPath(import.meta.url),
    path: __dirname,
    version: '5.812',
    buildVersion: 5812,
    currentDirectory: process.cwd(),
  };
}

/** Wires up the ActiveXObject/GetObject/Enumerator host objects when available. */
async function wireComHost(): Promise<void> {
  // ActiveXObject/GetObject come from node-ps1-dotnet if the user has it installed
  // (and injects it). We try to load it, but its absence is not fatal.
  try {
    const activexModuleName = '@devscholar/node-ps1-dotnet/activex';
    const mod = await import(activexModuleName);
    const g = globalThis as Record<string, unknown>;
    if (typeof mod.ActiveXObject === 'function' && !g.ActiveXObject) {
      g.ActiveXObject = mod.ActiveXObject;
    }
    if (typeof mod.GetObject === 'function' && !g.GetObject) {
      g.GetObject = mod.GetObject;
    }
    // Enumerator is only meaningful alongside ActiveXObject, matching its COM-only
    // contract. Inject it only when ActiveXObject is present.
    if (typeof mod.Enumerator === 'function' && g.ActiveXObject && !g.Enumerator) {
      g.Enumerator = mod.Enumerator;
    }
  } catch {
    // node-ps1-dotnet not installed — COM objects simply won't be available.
  }
}

async function main(): Promise<void> {
  const cli = parseArgs(process.argv.slice(2));

  if (!cli.nologo) {
    console.error('Microsoft (R) Windows Script Host Version 5.812');
    console.error('Copyright (C) Microsoft Corporation. All rights reserved.\n');
  }

  const { blocks, job, kind } = resolveSource(cli);
  await wireComHost();

  const env = buildEnvironment(cli, job);
  const { object: wscriptObject, getQuitCode } = createWScript(env);

  // Inject the WScript host object into globalThis, where the engine's global
  // fallback (getVariable → globalThis) will pick it up. JScript blocks resolve
  // WScript from the same globalThis.
  (globalThis as Record<string, unknown>).WScript = wscriptObject;

  const engine = new VbsEngine({ mode: 'general' });

  // MsgBox and InputBox are engine built-ins and already show a native GUI dialog
  // in a Node host (their default registration does so). The only thing this CLI
  // host adds is `//B` batch mode, where real WSH suppresses alerts/input prompts
  // entirely: override them to return the default button (MsgBox) / empty string
  // (InputBox) without showing a dialog.
  if (cli.batch) {
    engine._registerFunction(
      'MsgBox',
      createMsgBox({ gui: args => getDefaultResult(args.buttonType, args.defaultButton) })
    );
    engine._registerFunction('InputBox', createInputBox({ gui: () => '' }));
  }

  try {
    // `<object id=... progid=.../>` declarations are hoisted by real WSH into an
    // implicit `Set <id> = CreateObject(<progid>)` before any script runs. Emit the
    // equivalent prologue so a script can use `fso` / `wshshell` directly.
    for (const obj of job?.objects ?? []) {
      if (obj.id && obj.progid) {
        engine.addCode(`Set ${obj.id} = CreateObject(${JSON.stringify(obj.progid)})`);
      }
    }

    // Execute the job's script blocks in document order, so VBScript and JScript
    // share one global scope — matching real WSH mixed-language jobs.
    for (const block of blocks) {
      const lang = block.language.toLowerCase();
      if (lang === 'vbscript' || lang === 'vbs') {
        engine.addCode(block.code, {
          name: block.name,
          ...(block.startingLineNumber !== undefined
            ? { startingLineNumber: block.startingLineNumber }
            : {}),
          ...(block.columnOffset !== undefined ? { columnOffset: block.columnOffset } : {}),
        });
        // A script that fails stops there, the way it does under Windows Script Host, which prints the one line
        // and leaves. `addCode` records the failure rather than throwing it, so it is read back here.
        if (engine.error) reportFailure(engine.error, block.name, kind);
      } else if (lang === 'jscript' || lang === 'javascript' || lang === 'js') {
        runJScript(block.code);
      } else {
        console.error(`run-script: warning: unknown script language '${block.language}'`);
      }
    }
  } catch (err) {
    // WScript.Quit unwinds via this throw path; surface it as the exit code.
    if (err instanceof WScriptQuitSignal) {
      process.exit(getQuitCode() ?? err.exitCode);
    }
    const msg = err instanceof Error ? err.message : String(err);
    console.error(msg);
    process.exit(1);
  }

  const quitCode = getQuitCode();
  if (quitCode !== null) {
    process.exit(quitCode);
  }
  process.exit(0);
}

main().catch(err => {
  console.error(err instanceof Error ? err.message : String(err));
  process.exit(1);
});
