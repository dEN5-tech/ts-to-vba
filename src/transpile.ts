import fs from 'node:fs';
import path from 'node:path';
import ts from 'typescript';
import { emitClassModules } from './emitter/classEmitter';
import { VbaEmitter } from './emitter/vbaEmitter';
import { RuntimeFeature } from './runtime/featureDetection';
import { detectRuntimeFeatures } from './runtime/featureDetection';
import { renderRuntimeSections } from './runtime/fragments';
import { TstVbaConfig } from './types';

const TSTVBA_BUILD_VERSION = '1.0.0-template';

export interface TranspileResult {
  outputPath: string;
  vbaCode: string;
  generatedFiles: string[];
  manifestPath: string;
}

function toVbaSafeName(name: string): string {
  const sanitized = name.replace(/[^A-Za-z0-9_]/g, '_');
  if (!sanitized) {
    return 'Module';
  }

  if (/^[0-9]/.test(sanitized)) {
    return `M_${sanitized}`;
  }

  return sanitized;
}

function ensureCleanOutputDir(outDir: string): void {
  if (fs.existsSync(outDir)) {
    fs.rmSync(outDir, { recursive: true, force: true });
  }
  fs.mkdirSync(outDir, { recursive: true });
}

function createBuildMetadataHeader(config: TstVbaConfig): string {
  const timestamp = new Date().toISOString();
  return [
    "' =============================================",
    "' DO NOT EDIT - GENERATED CODE (TSTVBA)",
    `' Version: ${TSTVBA_BUILD_VERSION}`,
    `' GeneratedAt: ${timestamp}`,
    `' Project: ${config.projectFilePath}`,
    "' =============================================",
    '',
  ].join('\r\n');
}

function sanitizeVbaCode(code: string): string {
  const normalized = code
    .replace(/\u00A0/g, ' ')
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2013\u2014]/g, '-')
    .replace(/\r\n|\r|\n/g, '\r\n');

  let asciiOnly = '';
  for (const ch of normalized) {
    const codePoint = ch.charCodeAt(0);
    const isAllowedControl = codePoint === 9 || codePoint === 10 || codePoint === 13;
    const isPrintableAscii = codePoint >= 32 && codePoint <= 126;

    if (isAllowedControl || isPrintableAscii) {
      asciiOnly += ch;
    }
  }

  return asciiOnly;
}

function splitDeclarationsAndProcedures(vbaCode: string): { declarations: string; procedures: string } {
  const lines = vbaCode.split(/\r\n|\n|\r/);
  const declarations: string[] = [];
  const procedures: string[] = [];

  let inProcedure = false;
  for (const line of lines) {
    const trimmed = line.trim();
    if (/^Option\s+Explicit$/i.test(trimmed)) {
      continue;
    }

    if (/^(Public|Private)\s+(Sub|Function)\b/i.test(trimmed)) {
      inProcedure = true;
      procedures.push(line);
      continue;
    }

    if (/^End\s+(Sub|Function)$/i.test(trimmed)) {
      procedures.push(line);
      inProcedure = false;
      continue;
    }

    if (inProcedure) {
      procedures.push(line);
      continue;
    }

    if (/^(Public|Private)\b/i.test(trimmed) || trimmed.length > 0) {
      declarations.push(line);
    }
  }

  return {
    declarations: declarations.join('\r\n'),
    procedures: procedures.join('\r\n'),
  };
}

function resolveRelativeImport(fromFile: string, specifier: string): string | null {
  if (!specifier.startsWith('.')) {
    return null;
  }

  const base = path.resolve(path.dirname(fromFile), specifier);
  const candidates = [
    base,
    `${base}.ts`,
    `${base}.tsx`,
    path.join(base, 'index.ts'),
    path.join(base, 'index.tsx'),
  ];

  for (const candidate of candidates) {
    if (!fs.existsSync(candidate) || !fs.statSync(candidate).isFile()) {
      continue;
    }

    if (/\.d\.ts$/i.test(candidate)) {
      continue;
    }

    if (/\.tsx?$/i.test(candidate)) {
      return path.resolve(candidate);
    }
  }

  return null;
}

function collectReachableSourceFiles(entryPath: string): ts.SourceFile[] {
  const ordered: ts.SourceFile[] = [];
  const visited = new Set<string>();
  const stack = [path.resolve(entryPath)];

  while (stack.length) {
    const filePath = stack.pop() as string;
    if (visited.has(filePath) || !fs.existsSync(filePath)) {
      continue;
    }

    visited.add(filePath);
    const sourceText = fs.readFileSync(filePath, 'utf8');
    const sourceFile = ts.createSourceFile(filePath, sourceText, ts.ScriptTarget.Latest, true);
    ordered.push(sourceFile);

    for (const stmt of sourceFile.statements) {
      const importLike =
        (ts.isImportDeclaration(stmt) || ts.isExportDeclaration(stmt)) &&
        stmt.moduleSpecifier &&
        ts.isStringLiteral(stmt.moduleSpecifier)
          ? stmt.moduleSpecifier.text
          : null;

      if (!importLike) {
        continue;
      }

      const resolved = resolveRelativeImport(filePath, importLike);
      if (resolved && !visited.has(resolved)) {
        stack.push(resolved);
      }
    }
  }

  return ordered;
}

function getModuleNameForSourceFile(entryPath: string, sourceFilePath: string): string {
  const entryDir = path.dirname(path.resolve(entryPath));
  const relativeDir = path.dirname(path.relative(entryDir, path.resolve(sourceFilePath)));

  if (relativeDir === '.') {
    return toVbaSafeName(path.basename(entryPath, path.extname(entryPath)));
  }

  return toVbaSafeName(relativeDir.split(path.sep).join('_'));
}

export function transpileProject(config: TstVbaConfig): TranspileResult {
  if (!fs.existsSync(config.entry)) {
    throw new Error(`Entry file not found: ${config.entry}`);
  }

  const allSourceFiles = collectReachableSourceFiles(config.entry);
  const sourceFile = allSourceFiles[0];
  if (!sourceFile) {
    throw new Error(`Unable to parse entry file: ${config.entry}`);
  }

  const moduleName = getModuleNameForSourceFile(config.entry, sourceFile.fileName);
  const emitter = new VbaEmitter({
    namespacePrefix: config.tstvbaOptions.namespacePrefix,
    moduleName,
  });
  const emittedCode = emitter.emit(sourceFile);
  const emittedSections = splitDeclarationsAndProcedures(emittedCode);

  const runtimeFeatures = new Set<RuntimeFeature>();
  for (const sf of allSourceFiles) {
    const features = detectRuntimeFeatures(sf);
    for (const feature of features) {
      runtimeFeatures.add(feature);
    }
  }
  const runtimeSections = renderRuntimeSections(runtimeFeatures);

  let externalVbaLib = '';
  const vbaLibPath = path.resolve(path.dirname(config.projectFilePath), config.tstvbaOptions.vbaLibraryPath);
  if (fs.existsSync(vbaLibPath)) {
    externalVbaLib = [
      "' ===== External VBALib =====",
      fs.readFileSync(vbaLibPath, 'utf8'),
      '',
    ].join('\r\n');
  }

  const metadataHeader = createBuildMetadataHeader(config);
  const vbaCode = [
    metadataHeader,
    'Option Explicit',
    '',
    runtimeSections.declarations,
    emittedSections.declarations,
    externalVbaLib,
    runtimeSections.procedures,
    emittedSections.procedures,
  ]
    .filter(Boolean)
    .join('\r\n');

  const classModules = allSourceFiles.flatMap((sf) =>
    emitClassModules(sf, {
      namespacePrefix: config.tstvbaOptions.namespacePrefix,
      moduleName: getModuleNameForSourceFile(config.entry, sf.fileName),
    }),
  );

  ensureCleanOutputDir(config.outDir);

  const generatedFiles: string[] = [];
  const outputPath = path.join(config.outDir, config.tstvbaOptions.outputFileName);
  fs.writeFileSync(outputPath, sanitizeVbaCode(vbaCode), { encoding: 'ascii' });
  generatedFiles.push(outputPath);

  for (const classModule of classModules) {
    const classPath = path.join(config.outDir, classModule.fileName);
    fs.writeFileSync(classPath, sanitizeVbaCode(classModule.content), { encoding: 'ascii' });
    generatedFiles.push(classPath);
  }

  const manifestPath = path.join(config.outDir, 'tstvba-manifest.json');
  const manifest = {
    version: TSTVBA_BUILD_VERSION,
    generatedAt: new Date().toISOString(),
    entry: config.entry,
    outDir: config.outDir,
    outputFileName: config.tstvbaOptions.outputFileName,
    runtimeFeatures: Array.from(runtimeFeatures).sort(),
    files: generatedFiles,
  };
  fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2), 'utf8');

  return { outputPath, vbaCode, generatedFiles, manifestPath };
}
