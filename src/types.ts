export type TargetApplication = 'Excel' | 'Access' | 'Word';
export type ModuleStyle = 'StandardModule' | 'ClassModule';

export interface TstVbaOptions {
  entry?: string;
  targetApplication: TargetApplication;
  moduleStyle: ModuleStyle;
  vbaLibraryPath: string;
  namespacePrefix: string;
  emitSourceMaps: boolean;
  bundle: boolean;
  outputFileName: string;
}

export interface TstVbaConfig {
  projectFilePath: string;
  entry: string;
  outDir: string;
  tstvbaOptions: TstVbaOptions;
}
