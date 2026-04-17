import fs from 'node:fs';
import path from 'node:path';
import ts from 'typescript';
import { TstVbaConfig } from './types';

type TsConfigLike = {
  compilerOptions?: {
    outDir?: string;
  };
  tstvbaOptions?: Partial<TstVbaConfig['tstvbaOptions']>;
};

const defaultOptions: TstVbaConfig['tstvbaOptions'] = {
  entry: 'example/main.ts',
  targetApplication: 'Excel',
  moduleStyle: 'StandardModule',
  vbaLibraryPath: './lib/vbalib.bas',
  namespacePrefix: 'TS_',
  emitSourceMaps: true,
  bundle: true,
  outputFileName: 'MyProject.bas',
};

function formatTsError(error: ts.Diagnostic): string {
  return ts.flattenDiagnosticMessageText(error.messageText, '\n');
}

function resolveProjectFilePath(projectPath: string): string {
  const absPath = path.resolve(projectPath);

  if (fs.existsSync(absPath) && fs.statSync(absPath).isDirectory()) {
    const candidates = ['tsconfig.tstvba.json', 'tsconfig.json'];
    for (const candidate of candidates) {
      const candidatePath = path.join(absPath, candidate);
      if (fs.existsSync(candidatePath)) {
        return candidatePath;
      }
    }
    throw new Error(`No tsconfig file found in directory: ${absPath}`);
  }

  return absPath;
}

export function loadConfig(configPath: string): TstVbaConfig {
  const absPath = resolveProjectFilePath(configPath);
  if (!fs.existsSync(absPath)) {
    throw new Error(`Config file not found: ${absPath}`);
  }

  const readResult = ts.readConfigFile(absPath, ts.sys.readFile);
  if (readResult.error) {
    throw new Error(`Failed to read tsconfig: ${formatTsError(readResult.error)}`);
  }

  const parsed = (readResult.config as TsConfigLike) ?? {};
  const projectDir = path.dirname(absPath);
  const mergedOptions = {
    ...defaultOptions,
    ...(parsed.tstvbaOptions ?? {}),
  };

  const entry = path.resolve(projectDir, mergedOptions.entry ?? defaultOptions.entry ?? 'example/main.ts');
  const outDir = path.resolve(projectDir, parsed.compilerOptions?.outDir ?? 'dist-vba');

  return {
    projectFilePath: absPath,
    entry,
    outDir,
    tstvbaOptions: mergedOptions,
  };
}

export function createDefaultProjectConfig(): Record<string, unknown> {
  return {
    compilerOptions: {
      target: 'ESNext',
      module: 'CommonJS',
      strict: true,
      outDir: 'dist-vba',
    },
    tstvbaOptions: defaultOptions,
  };
}
