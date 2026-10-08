import { describe, it, expect } from 'vitest';
import { parseWsf } from '../host/wsf.ts';

describe('WSF parser', () => {
  it('parses a single <job> without a <package> wrapper', () => {
    const xml = `<job id="J1"><script language="VBScript">WScript.Echo 1</script></job>`;
    const doc = parseWsf(xml);
    expect(doc.jobs).toHaveLength(1);
    expect(doc.jobs[0]!.id).toBe('J1');
    expect(doc.jobs[0]!.scripts[0]!.code).toBe('WScript.Echo 1');
  });

  it('parses multiple jobs under a <package>', () => {
    const xml = `<package>
      <job id="A"><script language="VBScript">a</script></job>
      <job id="B"><script language="JScript">b</script></job>
    </package>`;
    const doc = parseWsf(xml);
    expect(doc.jobs).toHaveLength(2);
    expect(doc.jobs.map(j => j.id)).toEqual(['A', 'B']);
    expect(doc.jobs[1]!.scripts[0]!.language).toBe('JScript');
  });

  it('extracts CDATA-wrapped script source verbatim', () => {
    const xml = `<?XML version="1.0"?><job id="J"><script language="VBScript"><![CDATA[
WScript.Echo "x"
]]></script></job>`;
    const doc = parseWsf(xml);
    expect(doc.jobs[0]!.scripts[0]!.code).toContain('WScript.Echo "x"');
  });

  it('parses <runtime> named/unnamed params and description/example', () => {
    const xml = `<job id="J">
      <runtime>
        <description>Demo.</description>
        <named name="Server" helpstring="Server name" type="string" required="false"/>
        <unnamed name="Path" helpstring="File path"/>
        <example>Example: x.wsf /Server:prod</example>
      </runtime>
      <script language="VBScript">x</script>
    </job>`;
    const job = parseWsf(xml).jobs[0]!;
    expect(job.description).toBe('Demo.');
    expect(job.example).toBe('Example: x.wsf /Server:prod');
    expect(job.named).toHaveLength(1);
    expect(job.named[0]!.name).toBe('Server');
    expect(job.named[0]!.required).toBe(false);
    expect(job.unnamed).toHaveLength(1);
    expect(job.unnamed[0]!.name).toBe('Path');
  });

  it('parses <object>/<reference>/<resource> blocks', () => {
    const xml = `<job id="J">
      <object id="fso" progid="Scripting.FileSystemObject"/>
      <reference object="Scripting.FileSystemObject"/>
      <resource id="text">hello</resource>
      <script language="VBScript">x</script>
    </job>`;
    const job = parseWsf(xml).jobs[0]!;
    expect(job.objects[0]!.progid).toBe('Scripting.FileSystemObject');
    expect(job.references[0]!.object).toBe('Scripting.FileSystemObject');
    expect(job.resources['text']).toBe('hello');
  });

  it('rejects a document with no job', () => {
    expect(() => parseWsf('<package></package>')).toThrow();
  });
});
