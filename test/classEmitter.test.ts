import ts from 'typescript';
import { emitClassModules } from '../src/emitter/classEmitter';

describe('class module mirroring', () => {
  it('emits VBA class module for TS class', () => {
    const source = `
      export class User {
        public name: string;
        constructor(name: string) {
          this.name = name;
        }
        public greet(): string {
          return this.name;
        }
      }
    `;
    const sourceFile = ts.createSourceFile('main.ts', source, ts.ScriptTarget.Latest, true);

    const modules = emitClassModules(sourceFile, { namespacePrefix: 'TS_', moduleName: 'main' });

    expect(modules).toHaveLength(1);
    expect(modules[0].fileName).toBe('TS_main_User.cls');
    expect(modules[0].content).toContain('Public Sub Init(ByVal name As Variant)');
    expect(modules[0].content).toContain('Public Function greet() As Variant');
    expect(modules[0].content).toContain('m_name = name');
    expect(modules[0].content).toContain('ts_return = m_name');
  });

  it('flattens inheritance members for derived class', () => {
    const source = `
      export class BaseTransaction {
        constructor(public id: string, public amount: number) {}
        public getFormattedType(): string {
          return "GENERIC";
        }
      }

      export class Income extends BaseTransaction {
        constructor(id: string, amount: number, public source: string) {
          super(id, amount);
        }

        public getFormattedType(): string {
          return "INCOME";
        }
      }
    `;

    const sourceFile = ts.createSourceFile('main.ts', source, ts.ScriptTarget.Latest, true);
    const modules = emitClassModules(sourceFile, { namespacePrefix: 'TS_', moduleName: 'main' });
    const income = modules.find((m) => m.fileName === 'TS_main_Income.cls');

    expect(income).toBeDefined();
    expect(income?.content).toContain('Private m_id As Variant');
    expect(income?.content).toContain('Private m_amount As Variant');
    expect(income?.content).toContain('Private m_source As Variant');
    expect(income?.content).toContain('m_id = id');
    expect(income?.content).toContain('m_amount = amount');
    expect(income?.content).toContain('m_source = source');
    expect(income?.content).toContain('Public Function getFormattedType() As Variant');
    expect(income?.content).toContain('ts_return = "I" & "N" & "C" & "O" & "M" & "E"');
  });

  it('emits strict VBA class header with correct VB_Name as first line block', () => {
    const source = `
      export class ReportController {
        public run(): void {}
      }
    `;

    const sourceFile = ts.createSourceFile('presentation.ts', source, ts.ScriptTarget.Latest, true);
    const modules = emitClassModules(sourceFile, { namespacePrefix: 'TS_', moduleName: 'presentation' });
    const report = modules.find((m) => m.fileName === 'TS_presentation_ReportController.cls');

    expect(report).toBeDefined();
    const content = report?.content ?? '';

    expect(content.startsWith('VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  \'True\r\nEND\r\n')).toBe(true);
    expect(content).toContain('Attribute VB_Name = "TS_presentation_ReportController"');
    expect(content).toContain('Option Explicit');
  });
});
