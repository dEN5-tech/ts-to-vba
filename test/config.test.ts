import { createDefaultProjectConfig } from '../src/config';

describe('config helpers', () => {
  it('creates default config with required fields', () => {
    const cfg = createDefaultProjectConfig() as {
      compilerOptions: { outDir: string };
      tstvbaOptions: { targetApplication: string; outputFileName: string };
    };

    expect(cfg.compilerOptions.outDir).toBeTruthy();
    expect(cfg.tstvbaOptions.targetApplication).toBe('Excel');
    expect(cfg.tstvbaOptions.outputFileName).toBe('MyProject.bas');
  });
});
