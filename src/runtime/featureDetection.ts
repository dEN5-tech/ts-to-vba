import ts from 'typescript';

export type RuntimeFeature = 'console.log' | 'array.push' | 'error.stack' | 'iterator.protocol';

export function detectRuntimeFeatures(sourceFile: ts.SourceFile): Set<RuntimeFeature> {
  const features = new Set<RuntimeFeature>();

  const visit = (node: ts.Node): void => {
    if (ts.isTryStatement(node) || ts.isThrowStatement(node)) {
      features.add('error.stack');
    }

    if (ts.isForOfStatement(node)) {
      features.add('iterator.protocol');
    }

    if (ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression)) {
      const left = node.expression.expression;
      const right = node.expression.name.text;

      if (ts.isIdentifier(left) && left.text === 'console' && right === 'log') {
        features.add('console.log');
      }

      if (right === 'push') {
        features.add('array.push');
      }
    }

    ts.forEachChild(node, visit);
  };

  visit(sourceFile);
  return features;
}
