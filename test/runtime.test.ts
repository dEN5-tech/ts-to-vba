import ts from 'typescript';
import { detectRuntimeFeatures } from '../src/runtime/featureDetection';
import { renderRuntimeFragments } from '../src/runtime/fragments';

describe('runtime injection', () => {
  it('detects console.log and array.push and renders only required fragments', () => {
    const source = `
      const arr: number[] = [];
      arr.push(1);
      console.log(arr.length);
    `;
    const sourceFile = ts.createSourceFile('sample.ts', source, ts.ScriptTarget.Latest, true);

    const features = detectRuntimeFeatures(sourceFile);
    const runtime = renderRuntimeFragments(features);

    expect(features.has('console.log')).toBe(true);
    expect(features.has('array.push')).toBe(true);
    expect(runtime).toContain('TS_ConsoleLog');
    expect(runtime).toContain('TS_ArrayPush');
  });
});
