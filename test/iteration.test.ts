import ts from 'typescript';
import { VbaEmitter } from '../src/emitter/vbaEmitter';
import { detectRuntimeFeatures } from '../src/runtime/featureDetection';
import { renderRuntimeFragments } from '../src/runtime/fragments';

describe('iteration bridge', () => {
  it('emits for...of bridge for arrays and collections', () => {
    const source = `
      export function run(items: string[]): void {
        for (const item of items) {
          console.log(item);
        }
      }
    `;
    const sourceFile = ts.createSourceFile('main.ts', source, ts.ScriptTarget.Latest, true);
    const emitter = new VbaEmitter({ namespacePrefix: 'TS_', moduleName: 'main' });
    const output = emitter.emit(sourceFile);

    expect(output).toContain('If IsArray(items) And TS_HasArrayBounds(items) Then');
    expect(output).toContain('For ts_i_1 = LBound(items) To UBound(items)');
    expect(output).toContain('For Each item In items');
  });

  it('injects iterator runtime feature when for...of is present', () => {
    const source = `
      for (const x of arr) {
        console.log(x);
      }
    `;
    const sourceFile = ts.createSourceFile('sample.ts', source, ts.ScriptTarget.Latest, true);
    const features = detectRuntimeFeatures(sourceFile);
    const runtime = renderRuntimeFragments(features);

    expect(features.has('iterator.protocol')).toBe(true);
    expect(runtime).toContain('TS_HasArrayBounds');
  });
});
