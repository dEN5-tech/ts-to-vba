import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { transpileProject } from '../src/transpile';
import { TstVbaConfig } from '../src/types';

describe('generated files bundle (.bas + 3 .cls)', () => {
  it('creates MyProject.bas and 3 class modules that are listed in manifest', () => {
    const outDir = fs.mkdtempSync(path.join(os.tmpdir(), 'tstvba-files-'));

    const config: TstVbaConfig = {
      projectFilePath: path.resolve('tsconfig.json'),
      entry: path.resolve('example/main.ts'),
      outDir,
      tstvbaOptions: {
        targetApplication: 'Excel',
        moduleStyle: 'StandardModule',
        vbaLibraryPath: './lib/vbalib.bas',
        namespacePrefix: 'TS_',
        emitSourceMaps: true,
        bundle: true,
        outputFileName: 'MyProject.bas',
      },
    };

    const result = transpileProject(config);
    const basFile = path.join(outDir, 'MyProject.bas');
    const classFiles = [
      path.join(outDir, 'TS_main_BaseTransaction.cls'),
      path.join(outDir, 'TS_main_Income.cls'),
      path.join(outDir, 'TS_main_Expense.cls'),
    ];

    expect(fs.existsSync(basFile)).toBe(true);
    for (const classFile of classFiles) {
      expect(fs.existsSync(classFile)).toBe(true);
      expect(result.generatedFiles).toContain(classFile);
    }

    const manifest = JSON.parse(fs.readFileSync(result.manifestPath, 'utf8')) as { files: string[] };
    expect(manifest.files).toContain(basFile);
    for (const classFile of classFiles) {
      expect(manifest.files).toContain(classFile);
    }
  });
});
