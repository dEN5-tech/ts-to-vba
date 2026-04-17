import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { transpileProject } from '../src/transpile';
import { TstVbaConfig } from '../src/types';

describe('transpileProject multi-file class emission', () => {
  it('includes imported class modules in output and manifest', () => {
    const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'tstvba-'));
    const srcDir = path.join(tmpDir, 'src');
    const domainDir = path.join(srcDir, 'domain');
    const outDir = path.join(tmpDir, 'dist-vba');
    const projectFilePath = path.join(tmpDir, 'tsconfig.tstvba.json');
    const entry = path.join(srcDir, 'main.ts');

    fs.mkdirSync(domainDir, { recursive: true });
    fs.writeFileSync(
      entry,
      [
        "import { UserProfile } from './domain/UserProfile';",
        '',
        'export function run(): void {',
        "  const user = new UserProfile('Ada');",
        '  console.log(user.name);',
        '}',
        '',
      ].join('\n'),
      'utf8',
    );

    fs.writeFileSync(
      path.join(domainDir, 'UserProfile.ts'),
      [
        'export class UserProfile {',
        '  constructor(public name: string) {}',
        '}',
        '',
      ].join('\n'),
      'utf8',
    );

    fs.writeFileSync(projectFilePath, '{}', 'utf8');

    const config: TstVbaConfig = {
      projectFilePath,
      entry,
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
    const classFile = path.join(outDir, 'TS_domain_UserProfile.cls');

    expect(fs.existsSync(result.outputPath)).toBe(true);
    expect(fs.existsSync(classFile)).toBe(true);
    expect(result.generatedFiles).toContain(classFile);

    const manifestRaw = fs.readFileSync(result.manifestPath, 'utf8');
    const manifest = JSON.parse(manifestRaw) as { files: string[] };
    expect(manifest.files).toContain(result.outputPath);
    expect(manifest.files).toContain(classFile);
  });
});
