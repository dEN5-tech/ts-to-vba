# TypeScript to VBA (TSTVBA)

Transpile modern TypeScript into VBA modules (`.bas` + `.cls`) for Excel/Office automation.

## Installation

```bash
npm install -g typescript-to-vba
```

## Quick start

```bash
tstvba --init
tstvba check -p .
tstvba -p .
```

Local development usage (without global install):

```bash
npm install
npm run build
node dist/cli.js --init
node dist/cli.js check -p .
node dist/cli.js -p .
```

Generated VBA output: `dist/vba/MyProject.bas`

Build now also generates a bundle manifest: `dist/vba/tstvba-manifest.json`.

## Config model

`tstvbaOptions` are stored directly inside `tsconfig.json`:

```json
{
  "compilerOptions": {
    "target": "ESNext",
    "module": "CommonJS",
    "strict": true,
    "outDir": "dist/vba"
  },
  "tstvbaOptions": {
    "entry": "example/main.ts",
    "targetApplication": "Excel",
    "moduleStyle": "StandardModule",
    "vbaLibraryPath": "./lib/vbalib.bas",
    "namespacePrefix": "TS_",
    "emitSourceMaps": true,
    "bundle": true,
    "outputFileName": "MyProject.bas"
  }
}
```

## Commands

- `node dist/cli.js -p .` — build using project directory (`tsconfig.json` or `tsconfig.tstvba.json`)
- `node dist/cli.js --watch -p .` — watch mode
- `node dist/cli.js --init` — create `tsconfig.tstvba.json` + `lib/vbalib.bas` stub
- `node dist/cli.js check -p .` — validate and print resolved config

## Namespacing strategy (VBA global scope)

Template now uses **Static Prefixing** for exported function declarations:

- TS: `export function init()` in `main.ts`
- VBA: `Public Sub TS_main_init()`

Formula: `<namespacePrefix><moduleName>_<functionName>`

This avoids collisions between modules with identical function names.

## Object bridge: Class Module Mirroring

Current template includes first implementation of **TS class -> VBA `.cls`** generation:

- each top-level `class` generates a dedicated `.cls` file in output directory;
- generated class name uses namespacing pattern (`<namespacePrefix><moduleName>_<ClassName>`);
- constructor is mapped to `Public Sub Init(...)`;
- class fields become `Private m_<field>` + `Property Get/Let` accessors;
- methods are emitted as `Public Sub`/`Public Function` stubs.

### Import note for `.cls` files

Class module files contain metadata header (`VERSION/BEGIN/END` + `Attribute ...`) and must be imported via **VBE -> File -> Import File...**.

- If you copy/paste manually into a class code window, do **not** paste header lines above `Option Explicit`.
- Use generated `.cls` files from `dist/vba` directly for reliable class properties (`MultiUse`, `VB_Name`, etc.).

## Error handling bridge: Try/Catch -> On Error

Template now includes first pass of exception mapping:

- `try/catch/finally` is emitted using unique label blocks (`TS_CATCH_n`, `TS_FINALLY_n`, `TS_ENDTRY_n`);
- `throw` is mapped to `Err.Raise vbObjectError + 513, "TSTVBA", ...`;
- runtime auto-injects `error.stack` helpers (`TS_PushError`, `TS_LastErrorMessage`, `TS_ClearError`) when try/throw is detected.

## Iteration bridge: For...Of

Template now includes first pass of advanced iteration mapping:

- `for...of` emits hybrid VBA loop strategy:
  - arrays: `For i = LBound(...) To UBound(...)`;
  - non-array collections: `For Each item In collection`;
- runtime auto-injects `iterator.protocol` helper `TS_HasArrayBounds` when `ForOfStatement` is detected.

## Packaging strategy (Production bundle)

Current build pipeline includes packaging polish:

- `dist/vba` is cleaned before each transpile run;
- generated `.bas` file gets metadata header:
  - `DO NOT EDIT - GENERATED CODE`
  - compiler version
  - build timestamp
  - project path
- build writes `tstvba-manifest.json` with runtime features and all generated output files.

## Project structure

- `src/cli.ts` — CLI entrypoint
- `src/config.ts` — tsconfig + `tstvbaOptions` loader
- `src/transpile.ts` — transpilation pipeline template
- `src/emitter/vbaEmitter.ts` — basic AST emitter skeleton
