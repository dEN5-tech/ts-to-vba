import ts from 'typescript';
import { VbaEmitter } from '../src/emitter/vbaEmitter';
import { detectRuntimeFeatures } from '../src/runtime/featureDetection';

describe('error handling mapping', () => {
  it('emits label-based On Error flow for try/catch', () => {
    const source = `
      export function run(): void {
        try {
          throw 'boom';
        } catch (e) {
          console.log(e);
        }
      }
    `;
    const sourceFile = ts.createSourceFile('main.ts', source, ts.ScriptTarget.Latest, true);
    const emitter = new VbaEmitter({ namespacePrefix: 'TS_', moduleName: 'main' });
    const output = emitter.emit(sourceFile);

    expect(output).toContain('On Error GoTo TS_CATCH_1');
    expect(output).toContain('TS_CATCH_1:');
    expect(output).toContain('TS_PushError Err.Number, Err.Description');
    expect(output).toContain('Err.Raise vbObjectError + 513, "TSTVBA", CStr(');
  });

  it('detects error.stack runtime feature for try/catch and throw', () => {
    const source = `
      try {
        throw 'x';
      } catch (e) {
        console.log(e);
      }
    `;
    const sourceFile = ts.createSourceFile('sample.ts', source, ts.ScriptTarget.Latest, true);
    const features = detectRuntimeFeatures(sourceFile);

    expect(features.has('error.stack')).toBe(true);
  });
});
