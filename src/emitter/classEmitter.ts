import ts from 'typescript';

export interface GeneratedClassModule {
  className: string;
  fileName: string;
  content: string;
}

interface ClassEmitterOptions {
  namespacePrefix: string;
  moduleName: string;
}

interface ClassInfo {
  declaration: ts.ClassDeclaration;
  className: string;
  constructorDecl?: ts.ConstructorDeclaration;
}

function toVbaSafeName(name: string): string {
  const sanitized = name.replace(/[^A-Za-z0-9_]/g, '_');
  if (!sanitized) {
    return 'GeneratedClass';
  }

  if (/^[0-9]/.test(sanitized)) {
    return `C_${sanitized}`;
  }

  return sanitized;
}

function emitExpression(expression: ts.Expression): string {
  if (expression.kind === ts.SyntaxKind.ThisKeyword) {
    return 'Me';
  }

  if (ts.isIdentifier(expression)) {
    return expression.text;
  }

  if (ts.isStringLiteral(expression) || ts.isNoSubstitutionTemplateLiteral(expression)) {
    return emitStringLiteralValue(expression.text);
  }

  if (ts.isNumericLiteral(expression)) {
    return expression.text;
  }

  if (ts.isPropertyAccessExpression(expression)) {
    const owner = emitExpression(expression.expression);
    if (owner === 'this' || owner === 'Me') {
      return `m_${expression.name.text}`;
    }
    return `${owner}.${expression.name.text}`;
  }

  if (ts.isTemplateExpression(expression)) {
    const parts: string[] = [];

    if (expression.head.text) {
      parts.push(emitStringLiteralValue(expression.head.text));
    }

    for (const span of expression.templateSpans) {
      parts.push(`CStr(${emitExpression(span.expression)})`);
      if (span.literal.text) {
        parts.push(emitStringLiteralValue(span.literal.text));
      }
    }

    return parts.join(' & ');
  }

  if (ts.isBinaryExpression(expression)) {
    return `${emitExpression(expression.left)} ${mapBinaryOperator(expression.operatorToken.kind)} ${emitExpression(
      expression.right,
    )}`;
  }

  if (ts.isCallExpression(expression) && ts.isPropertyAccessExpression(expression.expression)) {
    return `${emitExpression(expression.expression.expression)}.${expression.expression.name.text}(${expression.arguments
      .map((arg) => emitExpression(arg))
      .join(', ')})`;
  }

  return expression.getText();
}

function emitStringLiteralValue(value: string): string {
  if (!value.length) {
    return '""';
  }

  const parts: string[] = [];
  for (const char of Array.from(value)) {
    const codePoint = char.codePointAt(0) ?? 0;
    if (codePoint >= 32 && codePoint <= 126 && char !== '"') {
      parts.push(`"${char}"`);
    } else if (char === '"') {
      parts.push('ChrW$(34)');
    } else {
      parts.push(`ChrW$(${codePoint})`);
    }
  }

  return parts.join(' & ');
}

function mapBinaryOperator(kind: ts.SyntaxKind): string {
  switch (kind) {
    case ts.SyntaxKind.EqualsEqualsEqualsToken:
    case ts.SyntaxKind.EqualsEqualsToken:
      return '=';
    case ts.SyntaxKind.ExclamationEqualsEqualsToken:
    case ts.SyntaxKind.ExclamationEqualsToken:
      return '<>';
    case ts.SyntaxKind.AmpersandAmpersandToken:
      return 'And';
    case ts.SyntaxKind.BarBarToken:
      return 'Or';
    case ts.SyntaxKind.GreaterThanToken:
      return '>';
    case ts.SyntaxKind.GreaterThanEqualsToken:
      return '>=';
    case ts.SyntaxKind.LessThanToken:
      return '<';
    case ts.SyntaxKind.LessThanEqualsToken:
      return '<=';
    case ts.SyntaxKind.PlusToken:
      return '+';
    case ts.SyntaxKind.MinusToken:
      return '-';
    case ts.SyntaxKind.AsteriskToken:
      return '*';
    case ts.SyntaxKind.SlashToken:
      return '/';
    default:
      return '=';
  }
}

function emitStatements(block: ts.Block | undefined, indentLevel = 1): string[] {
  if (!block) {
    return [];
  }

  const lines: string[] = [];
  for (const statement of block.statements) {
    lines.push(...emitStatement(statement, indentLevel));
  }

  return lines;
}

function emitStatement(statement: ts.Statement, indentLevel: number): string[] {
  const indent = '    '.repeat(indentLevel);

  if (
    ts.isExpressionStatement(statement) &&
    ts.isBinaryExpression(statement.expression) &&
    statement.expression.operatorToken.kind === ts.SyntaxKind.EqualsToken
  ) {
    return [`${indent}${emitExpression(statement.expression.left)} = ${emitExpression(statement.expression.right)}`];
  }

  if (ts.isReturnStatement(statement) && statement.expression) {
    return [`${indent}ts_return = ${emitExpression(statement.expression)}`];
  }

  if (ts.isIfStatement(statement)) {
    return emitIfStatement(statement, indentLevel, false);
  }

  if (ts.isBlock(statement)) {
    return emitStatements(statement, indentLevel);
  }

  return [];
}

function emitIfStatement(node: ts.IfStatement, indentLevel: number, isElseIf: boolean): string[] {
  const indent = '    '.repeat(indentLevel);
  const header = isElseIf ? `ElseIf ${emitExpression(node.expression)} Then` : `If ${emitExpression(node.expression)} Then`;
  const lines: string[] = [`${indent}${header}`];

  lines.push(...emitThenOrElseBody(node.thenStatement, indentLevel + 1));

  if (node.elseStatement) {
    if (ts.isIfStatement(node.elseStatement)) {
      lines.push(...emitIfStatement(node.elseStatement, indentLevel, true));
    } else {
      lines.push(`${indent}Else`);
      lines.push(...emitThenOrElseBody(node.elseStatement, indentLevel + 1));
    }
  }

  if (!isElseIf) {
    lines.push(`${indent}End If`);
  }

  return lines;
}

function emitThenOrElseBody(statement: ts.Statement, indentLevel: number): string[] {
  if (ts.isBlock(statement)) {
    return emitStatements(statement, indentLevel);
  }

  return emitStatement(statement, indentLevel);
}

function collectCtorParameterPropertyNames(constructorDecl: ts.ConstructorDeclaration | undefined): string[] {
  if (!constructorDecl) {
    return [];
  }

  const names: string[] = [];
  for (const param of constructorDecl.parameters) {
    if (!ts.isIdentifier(param.name)) {
      continue;
    }

    const hasAccessModifier = (ts.getCombinedModifierFlags(param) &
      (ts.ModifierFlags.Public | ts.ModifierFlags.Private | ts.ModifierFlags.Protected)) !==
      0;
    if (hasAccessModifier) {
      names.push(param.name.text);
    }
  }

  return names;
}

function getCtorParameterNames(constructorDecl: ts.ConstructorDeclaration | undefined): string[] {
  if (!constructorDecl) {
    return [];
  }

  return constructorDecl.parameters
    .map((p) => p.name)
    .filter(ts.isIdentifier)
    .map((id) => id.text);
}

function collectClassPropertyNames(statement: ts.ClassDeclaration): string[] {
  const names = new Set<string>();

  for (const member of statement.members) {
    if (ts.isPropertyDeclaration(member) && ts.isIdentifier(member.name)) {
      names.add(member.name.text);
    }

    if (ts.isConstructorDeclaration(member)) {
      for (const param of member.parameters) {
        if (!ts.isIdentifier(param.name)) {
          continue;
        }

        const hasAccessModifier = (ts.getCombinedModifierFlags(param) &
          (ts.ModifierFlags.Public | ts.ModifierFlags.Private | ts.ModifierFlags.Protected)) !==
          0;
        if (hasAccessModifier) {
          names.add(param.name.text);
        }
      }
    }
  }

  return Array.from(names);
}

function getExtendsBaseName(statement: ts.ClassDeclaration): string | null {
  const clause = statement.heritageClauses?.find((item) => item.token === ts.SyntaxKind.ExtendsKeyword);
  const firstType = clause?.types[0];
  if (!firstType) {
    return null;
  }

  if (ts.isIdentifier(firstType.expression)) {
    return firstType.expression.text;
  }

  return null;
}

function getInheritanceChain(className: string, classIndex: Map<string, ClassInfo>): ClassInfo[] {
  const chain: ClassInfo[] = [];
  const visited = new Set<string>();

  let current: ClassInfo | undefined = classIndex.get(className);
  while (current && !visited.has(current.declaration.name?.text ?? '')) {
    const currentName = current.declaration.name?.text ?? '';
    visited.add(currentName);
    chain.unshift(current);
    const baseName = getExtendsBaseName(current.declaration);
    current = baseName ? classIndex.get(baseName) : undefined;
  }

  return chain;
}

function collectFlattenedPropertyNames(chain: ClassInfo[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];

  for (const info of chain) {
    for (const propertyName of collectClassPropertyNames(info.declaration)) {
      if (!seen.has(propertyName)) {
        seen.add(propertyName);
        result.push(propertyName);
      }
    }
  }

  return result;
}

function collectFlattenedMethods(chain: ClassInfo[]): ts.MethodDeclaration[] {
  const methods = new Map<string, ts.MethodDeclaration>();

  for (const info of chain) {
    for (const member of info.declaration.members) {
      if (!ts.isMethodDeclaration(member) || !member.name || !ts.isIdentifier(member.name)) {
        continue;
      }

      methods.set(member.name.text, member);
    }
  }

  return Array.from(methods.values());
}

function getSuperCall(constructorDecl: ts.ConstructorDeclaration | undefined): ts.CallExpression | undefined {
  if (!constructorDecl?.body) {
    return undefined;
  }

  const first = constructorDecl.body.statements[0];
  if (!first || !ts.isExpressionStatement(first) || !ts.isCallExpression(first.expression)) {
    return undefined;
  }

  if (first.expression.expression.kind !== ts.SyntaxKind.SuperKeyword) {
    return undefined;
  }

  return first.expression;
}

export function emitClassModules(
  sourceFile: ts.SourceFile,
  options: ClassEmitterOptions,
): GeneratedClassModule[] {
  const modules: GeneratedClassModule[] = [];

  const classIndex = new Map<string, ClassInfo>();
  for (const statement of sourceFile.statements) {
    if (!ts.isClassDeclaration(statement) || !statement.name) {
      continue;
    }

    const originalName = statement.name.text;
    const className = toVbaSafeName(`${options.namespacePrefix}${options.moduleName}_${originalName}`);
    const constructorDecl = statement.members.find(
      (member): member is ts.ConstructorDeclaration => ts.isConstructorDeclaration(member),
    );

    classIndex.set(originalName, {
      declaration: statement,
      className,
      constructorDecl,
    });
  }

  for (const statement of sourceFile.statements) {
    if (!ts.isClassDeclaration(statement) || !statement.name) {
      continue;
    }

    const originalName = statement.name.text;
    const classInfo = classIndex.get(originalName);
    if (!classInfo) {
      continue;
    }

    const className = classInfo.className;
    const chain = getInheritanceChain(originalName, classIndex);
    const properties = collectFlattenedPropertyNames(chain);
    const methods = collectFlattenedMethods(chain);
    const constructorDecl = classInfo.constructorDecl;
    const effectiveCtor = constructorDecl ?? [...chain].reverse().find((item) => item.constructorDecl)?.constructorDecl;

    const lines: string[] = [
      'VERSION 1.0 CLASS',
      'BEGIN',
      "  MultiUse = -1  'True",
      'END',
      `Attribute VB_Name = "${className}"`,
      'Attribute VB_GlobalNameSpace = False',
      'Attribute VB_Creatable = False',
      'Attribute VB_PredeclaredId = False',
      'Attribute VB_Exposed = False',
      'Option Explicit',
      '',
    ];

    for (const propertyName of properties) {
      lines.push(`Private m_${propertyName} As Variant`);
    }

    if (properties.length) {
      lines.push('');
    }

    if (effectiveCtor) {
      const ctorParams = effectiveCtor.parameters
        .map((p) => `ByVal ${p.name.getText()} As Variant`)
        .join(', ');
      lines.push(`Public Sub Init(${ctorParams})`);
      for (const paramPropName of collectCtorParameterPropertyNames(effectiveCtor)) {
        lines.push(`    m_${paramPropName} = ${paramPropName}`);
      }

      if (constructorDecl) {
        const baseName = getExtendsBaseName(statement);
        const baseInfo = baseName ? classIndex.get(baseName) : undefined;
        const superCall = getSuperCall(constructorDecl);
        if (baseInfo && superCall) {
          const basePropNames = collectCtorParameterPropertyNames(baseInfo.constructorDecl);
          const baseParamNames = getCtorParameterNames(baseInfo.constructorDecl);

          for (let i = 0; i < basePropNames.length && i < superCall.arguments.length; i++) {
            const arg = superCall.arguments[i];
            const fallbackName = baseParamNames[i] ?? basePropNames[i];
            lines.push(`    m_${basePropNames[i]} = ${emitExpression(arg) || fallbackName}`);
          }
        }
      }

      const ctorLines = emitStatements(constructorDecl?.body ?? effectiveCtor.body);
      if (ctorLines.length) {
        lines.push(...ctorLines);
      }
      lines.push('End Sub');
      lines.push('');
    }

    for (const propertyName of properties) {
      lines.push(`Public Property Get ${propertyName}() As Variant`);
      lines.push(`    ${propertyName} = m_${propertyName}`);
      lines.push('End Property');
      lines.push('');
      lines.push(`Public Property Let ${propertyName}(ByVal value As Variant)`);
      lines.push(`    m_${propertyName} = value`);
      lines.push('End Property');
      lines.push('');
    }

    for (const method of methods) {
      if (!method.name || !ts.isIdentifier(method.name)) {
        continue;
      }
      const methodName = method.name.text;
      const params = method.parameters
        .map((p) => `ByVal ${p.name.getText()} As Variant`)
        .join(', ');
      const isSub = method.type?.kind === ts.SyntaxKind.VoidKeyword;

      if (isSub) {
        lines.push(`Public Sub ${methodName}(${params})`);
        const methodLines = emitStatements(method.body);
        if (methodLines.length) {
          lines.push(...methodLines);
        }
        lines.push('End Sub');
      } else {
        lines.push(`Public Function ${methodName}(${params}) As Variant`);
        lines.push('    Dim ts_return As Variant');
        const methodLines = emitStatements(method.body);
        if (methodLines.length) {
          lines.push(...methodLines);
        }
        lines.push(`    ${methodName} = ts_return`);
        lines.push('End Function');
      }
      lines.push('');
    }

    modules.push({
      className,
      fileName: `${className}.cls`,
      content: lines.join('\r\n'),
    });
  }

  return modules;
}
