#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import { Command } from 'commander';
import { createDefaultProjectConfig, loadConfig } from './config';
import { transpileProject } from './transpile';

const DEFAULT_PROJECT_FILE = 'tsconfig.json';

function writeDefaultProjectConfig(targetPath: string): void {
  const absPath = path.resolve(targetPath);
  if (fs.existsSync(absPath)) {
    throw new Error(`Config file already exists: ${absPath}`);
  }

  fs.writeFileSync(absPath, JSON.stringify(createDefaultProjectConfig(), null, 2), 'utf8');
}

function ensureDefaultVbaLib(): string {
  const libDir = path.resolve('lib');
  const libPath = path.join(libDir, 'vbalib.bas');
  if (!fs.existsSync(libDir)) {
    fs.mkdirSync(libDir, { recursive: true });
  }

  if (!fs.existsSync(libPath)) {
    const stub = [
      "Attribute VB_Name = \"VbaLib\"",
      "Option Explicit",
      "",
      "' Placeholder for JS/VBA polyfills (Array.push, String.split, etc.)",
      'Public Function JsVersion() As String',
      '    JsVersion = "vbalib-template-0.1"',
      'End Function',
      '',
    ].join('\r\n');
    fs.writeFileSync(libPath, stub, 'utf8');
  }

  return libPath;
}

function runBuild(projectPath: string): void {
  const config = loadConfig(projectPath);
  const result = transpileProject(config);
  console.log(`Generated VBA: ${result.outputPath}`);
  console.log(`Manifest: ${result.manifestPath}`);
  console.log(`Bundle files (${result.generatedFiles.length}):`);
  for (const filePath of result.generatedFiles) {
    console.log(` - ${filePath}`);
  }
}

function runWatch(projectPath: string): void {
  const resolvedProjectPath = path.resolve(projectPath);
  runBuild(resolvedProjectPath);
  console.log(`Watching for changes: ${resolvedProjectPath}`);

  fs.watch(path.dirname(resolvedProjectPath), { recursive: true }, (_event, filename) => {
    if (!filename) {
      return;
    }

    if (!/\.(ts|json)$/i.test(filename)) {
      return;
    }

    try {
      runBuild(resolvedProjectPath);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.error(`Watch build failed: ${message}`);
    }
  });
}

function run(): void {
  const program = new Command();

  program
    .name('tstvba')
    .description('TypeScript to VBA Transpiler')
    .version('1.0.0')
    .option('-p, --project <path>', 'Path to tsconfig file or project directory', '.')
    .option('-w, --watch', 'Watch for TS/JSON changes and rebuild')
    .option('--init', 'Create default tsconfig.tstvba.json + vbalib stub');

  program
    .command('init')
    .description('Create default tsconfig with tstvbaOptions + vbalib stub')
    .option('-p, --project <path>', 'Target config file to create', 'tsconfig.tstvba.json')
    .action((options: { project: string }) => {
      writeDefaultProjectConfig(options.project);
      const libPath = ensureDefaultVbaLib();
      console.log(`Created ${path.resolve(options.project)}`);
      console.log(`Prepared ${libPath}`);
    });

  program
    .command('check')
    .description('Validate project config and print resolved values')
    .option('-p, --project <path>', 'Path to tsconfig file or project directory', DEFAULT_PROJECT_FILE)
    .action((options: { project: string }) => {
      const config = loadConfig(options.project);
      console.log(`Config is valid: ${config.projectFilePath}`);
      console.log(`Entry: ${config.entry}`);
      console.log(`OutDir: ${config.outDir}`);
      console.log(`Output file: ${config.tstvbaOptions.outputFileName}`);
    });

  program.action((options: { project: string; watch?: boolean; init?: boolean }) => {
    if (options.init) {
      writeDefaultProjectConfig('tsconfig.tstvba.json');
      const libPath = ensureDefaultVbaLib();
      console.log(`Created ${path.resolve('tsconfig.tstvba.json')}`);
      console.log(`Prepared ${libPath}`);
      return;
    }

    if (options.watch) {
      runWatch(options.project);
      return;
    }

    runBuild(options.project);
  });

  if (process.argv.length <= 2) {
    runBuild('.');
    return;
  }

  program.parse(process.argv);
}

run();
