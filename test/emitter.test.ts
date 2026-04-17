import ts from 'typescript';
import { VbaEmitter } from '../src/emitter/vbaEmitter';

describe('VbaEmitter namespacing', () => {
  it('prefixes function names with namespace + module name', () => {
    const source = `export function init(userName: string): void {}`;
    const sourceFile = ts.createSourceFile('main.ts', source, ts.ScriptTarget.Latest, true);

    const emitter = new VbaEmitter({ namespacePrefix: 'TS_', moduleName: 'main' });
    const output = emitter.emit(sourceFile);

    expect(output).toContain('Public Sub TS_main_init(ByVal userName As Variant)');
  });

  it('maps as-expression/globalThis and truthy checks to VBA-safe forms', () => {
    const source = `
      export function run(productName: string): void {
        const sheet: any = (globalThis as any).ActiveSheet;
        if (productName) {
          console.log(productName);
        }
      }
    `;
    const sourceFile = ts.createSourceFile('main.ts', source, ts.ScriptTarget.Latest, true);

    const emitter = new VbaEmitter({ namespacePrefix: 'TS_', moduleName: 'main' });
    const output = emitter.emit(sourceFile);

    expect(output).toContain('Set sheet = ActiveSheet');
    expect(output).toContain('If productName <> "" Then');
    expect(output).toContain('Call TS_ConsoleLog(productName)');
  });

  it('maps string includes() to InStr-based condition', () => {
    const source = `
      export function run(value: string): void {
        if (value.includes("X")) {
          console.log(value);
        }
      }
    `;
    const sourceFile = ts.createSourceFile('main.ts', source, ts.ScriptTarget.Latest, true);

    const emitter = new VbaEmitter({ namespacePrefix: 'TS_', moduleName: 'main' });
    const output = emitter.emit(sourceFile);

    expect(output).toContain('If (InStr(1, CStr(value), CStr("X")) > 0) Then');
  });

  it('does not emit top-level assignments for initialized globals', () => {
    const source = `
      let globalObserver: any = null;
      export function run(): void {
        if (globalObserver) {
          console.log(globalObserver);
        }
      }
    `;
    const sourceFile = ts.createSourceFile('main.ts', source, ts.ScriptTarget.Latest, true);

    const emitter = new VbaEmitter({ namespacePrefix: 'TS_', moduleName: 'main' });
    const output = emitter.emit(sourceFile);

    expect(output).toContain('Public globalObserver As Variant');
    expect(output).not.toContain('globalObserver = null');
  });

  it('emits object method calls and object truthiness checks', () => {
    const source = `
      let globalObserver: any;

      export function run(controller: any, user: any): void {
        controller.render(user);
      }

      export function handleSheetChange(target: any): void {
        if (globalObserver) {
          globalObserver.onCellChange(target);
        }
      }
    `;
    const sourceFile = ts.createSourceFile('main.ts', source, ts.ScriptTarget.Latest, true);

    const emitter = new VbaEmitter({ namespacePrefix: 'TS_', moduleName: 'main' });
    const output = emitter.emit(sourceFile);

    expect(output).toContain('Call controller.render(user)');
    expect(output).toContain('If Not globalObserver Is Nothing Then');
    expect(output).toContain('Call globalObserver.onCellChange(target)');
  });
});
