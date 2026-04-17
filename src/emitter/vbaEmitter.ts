import ts from 'typescript';

interface VbaEmitterOptions {
  namespacePrefix: string;
  moduleName: string;
}

export class VbaEmitter {
  private indentLevel = 0;
  private out: string[] = [];
  private tryCounter = 0;
  private forOfCounter = 0;
  private currentProcedureVariables: Map<string, string> | null = null;
  private currentProcedureParameters: Set<string> | null = null;
  private currentProcedureObjectArrays: Set<string> | null = null;
  private currentProcedureObjectVariables: Set<string> | null = null;
  private knownObjectVariables = new Set<string>();

  constructor(private readonly options: VbaEmitterOptions) {}

  public emit(sourceFile: ts.SourceFile): string {
    this.visit(sourceFile);
    return this.out.join('');
  }

  private writeLine(text = ''): void {
    this.out.push(`${'    '.repeat(this.indentLevel)}${text}\r\n`);
  }

  private visit(node: ts.Node): void {
    switch (node.kind) {
      case ts.SyntaxKind.SourceFile:
        this.visitSourceFile(node as ts.SourceFile);
        break;
      case ts.SyntaxKind.FunctionDeclaration:
        this.visitFunctionDeclaration(node as ts.FunctionDeclaration);
        break;
      case ts.SyntaxKind.ClassDeclaration:
        break;
      case ts.SyntaxKind.IfStatement:
        this.visitIfStatement(node as ts.IfStatement);
        break;
      case ts.SyntaxKind.TryStatement:
        this.visitTryStatement(node as ts.TryStatement);
        break;
      case ts.SyntaxKind.ForOfStatement:
        this.visitForOfStatement(node as ts.ForOfStatement);
        break;
      case ts.SyntaxKind.ForStatement:
        this.visitForStatement(node as ts.ForStatement);
        break;
      case ts.SyntaxKind.VariableStatement:
        this.visitVariableStatement(node as ts.VariableStatement);
        break;
      case ts.SyntaxKind.ExpressionStatement:
        this.visitExpressionStatement(node as ts.ExpressionStatement);
        break;
      case ts.SyntaxKind.Block:
        this.visitBlock(node as ts.Block);
        break;
      case ts.SyntaxKind.ThrowStatement:
        this.visitThrowStatement(node as ts.ThrowStatement);
        break;
      default:
        ts.forEachChild(node, (child) => this.visit(child));
    }
  }

  private visitSourceFile(node: ts.SourceFile): void {
    this.writeLine('Option Explicit');
    this.writeLine();

    for (const stmt of node.statements) {
      if (ts.isClassDeclaration(stmt) && stmt.name) {
        this.emitClassFactory(stmt);
      }
    }

    for (const stmt of node.statements) {
      if (
        ts.isFunctionDeclaration(stmt) ||
        ts.isVariableStatement(stmt) ||
        ts.isClassDeclaration(stmt)
      ) {
        this.visit(stmt);
      }
    }
  }

  private visitBlock(node: ts.Block): void {
    node.statements.forEach((stmt) => this.visit(stmt));
  }

  private visitIfStatement(node: ts.IfStatement): void {
    const condition = this.emitCondition(node.expression);
    this.writeLine(`If ${condition} Then`);
    this.indentLevel++;
    this.visit(node.thenStatement);
    this.indentLevel--;

    if (node.elseStatement) {
      this.writeLine('Else');
      this.indentLevel++;
      this.visit(node.elseStatement);
      this.indentLevel--;
    }

    this.writeLine('End If');
  }

  private visitFunctionDeclaration(node: ts.FunctionDeclaration): void {
    if (!node.name) {
      return;
    }

    const functionName = this.composeFunctionName(node.name.text);
    const parameters = node.parameters
      .map((p) => {
        const parameterName = p.name.getText();
        return `ByVal ${parameterName} As Variant`;
      })
      .join(', ');

    const isSub = node.type?.kind === ts.SyntaxKind.VoidKeyword;
    if (isSub) {
      this.writeLine(`Public Sub ${functionName}(${parameters})`);
    } else {
      this.writeLine(`Public Function ${functionName}(${parameters}) As Variant`);
    }

    const previousProcedureVariables = this.currentProcedureVariables;
    const previousProcedureParameters = this.currentProcedureParameters;
    const previousProcedureObjectArrays = this.currentProcedureObjectArrays;
    const previousProcedureObjectVariables = this.currentProcedureObjectVariables;
    this.currentProcedureVariables = new Map<string, string>();
    this.currentProcedureParameters = new Set(
      node.parameters
        .map((p) => p.name.getText())
        .filter((name) => name.length > 0),
    );
    this.currentProcedureObjectArrays = new Set<string>();
    this.currentProcedureObjectVariables = new Set<string>();

    const bodyInsertIndex = this.out.length;
    this.indentLevel++;
    if (node.body) {
      this.visit(node.body);
    }
    this.indentLevel--;

    const declarations = Array.from(this.currentProcedureVariables.entries())
      .sort((a, b) => a[0].localeCompare(b[0]))
      .map(([name, vbaType]) => `${'    '}Dim ${name} As ${vbaType}\r\n`);
    if (declarations.length) {
      this.out.splice(bodyInsertIndex, 0, ...declarations);
    }

    this.currentProcedureVariables = previousProcedureVariables;
    this.currentProcedureParameters = previousProcedureParameters;
    this.currentProcedureObjectArrays = previousProcedureObjectArrays;
    this.currentProcedureObjectVariables = previousProcedureObjectVariables;

    this.writeLine(isSub ? 'End Sub' : 'End Function');
    this.writeLine();
  }

  private visitTryStatement(node: ts.TryStatement): void {
    const blockId = ++this.tryCounter;
    const catchLabel = `TS_CATCH_${blockId}`;
    const finallyLabel = `TS_FINALLY_${blockId}`;
    const exitLabel = `TS_ENDTRY_${blockId}`;

    this.writeLine('On Error GoTo ' + catchLabel);
    this.visit(node.tryBlock);
    this.writeLine('On Error GoTo 0');
    this.writeLine('GoTo ' + finallyLabel);
    this.writeLine(catchLabel + ':');

    this.indentLevel++;
    this.writeLine('TS_PushError Err.Number, Err.Description');
    const catchVar = node.catchClause?.variableDeclaration?.name;
    if (catchVar && ts.isIdentifier(catchVar)) {
      this.registerProcedureVariable(catchVar.text, 'Variant');
      this.writeLine(`${catchVar.text} = TS_LastErrorMessage()`);
    }
    if (node.catchClause?.block) {
      this.visit(node.catchClause.block);
    }
    this.indentLevel--;

    this.writeLine(finallyLabel + ':');
    this.indentLevel++;
    if (node.finallyBlock) {
      this.visit(node.finallyBlock);
    }
    this.writeLine('TS_ClearError');
    this.indentLevel--;

    this.writeLine(exitLabel + ':');
  }

  private visitThrowStatement(node: ts.ThrowStatement): void {
    const message = node.expression ? this.emitExpression(node.expression) : '"Error"';
    this.writeLine(`Err.Raise vbObjectError + 513, "TSTVBA", CStr(${message})`);
  }

  private visitVariableStatement(node: ts.VariableStatement): void {
    for (const declaration of node.declarationList.declarations) {
      this.visitVariableDeclaration(declaration, node.declarationList.flags, node.parent);
    }
  }

  private visitVariableDeclaration(
    node: ts.VariableDeclaration,
    flags: ts.NodeFlags,
    parentNode: ts.Node,
  ): void {
    const name = node.name.getText();
    const isConst = (flags & ts.NodeFlags.Const) !== 0;
    const isTopLevel = parentNode.kind === ts.SyntaxKind.SourceFile;
    const initializer = node.initializer ? this.emitExpression(node.initializer) : undefined;

    if (this.currentProcedureVariables) {
      this.registerProcedureVariable(name, 'Variant');
      if (this.isLikelyObjectDeclaration(node)) {
        this.registerObjectVariable(name);
      }
      if (initializer) {
        const assignmentPrefix = node.initializer && this.needsSetAssignment(node.initializer) ? 'Set ' : '';
        if (assignmentPrefix) {
          this.registerObjectVariable(name);
        }
        this.writeLine(`${assignmentPrefix}${name} = ${initializer}`);
      }
      return;
    }

    if (isTopLevel && isConst && initializer) {
      this.writeLine(`Public Const ${name} = ${initializer}`);
      return;
    }

    if (this.isLikelyObjectDeclaration(node)) {
      this.registerObjectVariable(name);
    }

    if (isTopLevel) {
      this.writeLine(`Public ${name} As Variant`);
    } else {
      this.writeLine(`Dim ${name} As Variant`);
    }

    if (isTopLevel) {
      return;
    }

    if (initializer) {
      const assignmentPrefix = node.initializer && this.needsSetAssignment(node.initializer) ? 'Set ' : '';
      this.writeLine(`${assignmentPrefix}${name} = ${initializer}`);
    }
  }

  private visitExpressionStatement(node: ts.ExpressionStatement): void {
    const expression = node.expression;

    if (ts.isPostfixUnaryExpression(expression)) {
      const operand = this.emitExpression(expression.operand);
      if (expression.operator === ts.SyntaxKind.PlusPlusToken) {
        this.writeLine(`${operand} = ${operand} + 1`);
        return;
      }
      if (expression.operator === ts.SyntaxKind.MinusMinusToken) {
        this.writeLine(`${operand} = ${operand} - 1`);
        return;
      }
    }

    if (ts.isCallExpression(expression) && ts.isPropertyAccessExpression(expression.expression)) {
      const target = expression.expression.expression;
      const method = expression.expression.name.text;

      if (ts.isIdentifier(target) && target.text === 'console' && method === 'log') {
        const rendered = this.renderConsoleLogArgs(expression.arguments);
        this.writeLine(`Call TS_ConsoleLog(${rendered})`);
        return;
      }

      if (method === 'push' && expression.arguments.length >= 1) {
        const arrayTarget = this.emitExpression(target);
        const pushedValue = this.emitExpression(expression.arguments[0]);

        if (ts.isIdentifier(target) && this.needsSetAssignment(expression.arguments[0])) {
          this.currentProcedureObjectArrays?.add(target.text);
        }

        this.writeLine(`Call TS_ArrayPush(${arrayTarget}, ${pushedValue})`);
        return;
      }

      const callTarget = this.emitExpression(target);
      const args = expression.arguments.map((arg) => this.emitExpression(arg)).join(', ');
      this.writeLine(`Call ${callTarget}.${method}(${args})`);
      return;
    }

    if (ts.isCallExpression(expression) && ts.isIdentifier(expression.expression)) {
      const functionName = expression.expression.text;
      const args = expression.arguments.map((arg) => this.emitExpression(arg)).join(', ');
      this.writeLine(`Call ${functionName}(${args})`);
      return;
    }

    if (ts.isBinaryExpression(expression) && expression.operatorToken.kind === ts.SyntaxKind.EqualsToken) {
      const left = this.emitExpression(expression.left);
      const right = this.emitExpression(expression.right);
      const assignmentPrefix =
        ts.isIdentifier(expression.left) && this.needsSetAssignment(expression.right) ? 'Set ' : '';
      if (assignmentPrefix && ts.isIdentifier(expression.left)) {
        this.registerObjectVariable(expression.left.text);
      }
      this.writeLine(`${assignmentPrefix}${left} = ${right}`);
      return;
    }

    if (ts.isBinaryExpression(expression) && expression.operatorToken.kind === ts.SyntaxKind.PlusEqualsToken) {
      const left = this.emitExpression(expression.left);
      const right = this.emitExpression(expression.right);
      this.writeLine(`${left} = ${left} + ${right}`);
      return;
    }

    if (ts.isBinaryExpression(expression) && expression.operatorToken.kind === ts.SyntaxKind.MinusEqualsToken) {
      const left = this.emitExpression(expression.left);
      const right = this.emitExpression(expression.right);
      this.writeLine(`${left} = ${left} - ${right}`);
      return;
    }
  }

  private visitForOfStatement(node: ts.ForOfStatement): void {
    const loopId = ++this.forOfCounter;
    const iterableText = this.emitExpression(node.expression);
    const indexName = `ts_i_${loopId}`;
    const valueName = this.resolveForOfVariableName(node.initializer);
    const valueTarget = this.resolveForOfAssignmentTarget(node.initializer);

    if (valueName) {
      this.registerProcedureVariable(valueName, 'Variant');
    }

    this.registerProcedureVariable(indexName, 'Long');
    this.writeLine(`If IsArray(${iterableText}) And TS_HasArrayBounds(${iterableText}) Then`);
    this.indentLevel++;
    this.writeLine(`For ${indexName} = LBound(${iterableText}) To UBound(${iterableText})`);
    this.indentLevel++;
    const setPrefix = this.shouldUseSetForForOfValue(iterableText, valueTarget) ? 'Set ' : '';
    this.writeLine(`${setPrefix}${valueTarget} = ${iterableText}(${indexName})`);
    this.visit(node.statement);
    this.indentLevel--;
    this.writeLine('Next ' + indexName);
    this.indentLevel--;
    this.writeLine('Else');
    this.indentLevel++;
    this.writeLine(`For Each ${valueTarget} In ${iterableText}`);
    this.indentLevel++;
    this.visit(node.statement);
    this.indentLevel--;
    this.writeLine('Next ' + valueTarget);
    this.indentLevel--;
    this.writeLine('End If');
  }

  private visitForStatement(node: ts.ForStatement): void {
    const loopVarName = this.resolveForLoopVariableName(node.initializer);
    if (!loopVarName || !node.condition || !ts.isBinaryExpression(node.condition)) {
      if (node.initializer) {
        this.visit(node.initializer);
      }
      if (node.statement) {
        this.visit(node.statement);
      }
      return;
    }

    this.registerProcedureVariable(loopVarName, 'Long');

    const startExpr = this.resolveForLoopStartExpression(node.initializer);
    const endExpr = this.resolveForLoopEndExpression(node.condition);

    this.writeLine(`For ${loopVarName} = ${startExpr} To ${endExpr}`);
    this.indentLevel++;
    this.visit(node.statement);
    this.indentLevel--;
    this.writeLine(`Next ${loopVarName}`);
  }

  private resolveForLoopVariableName(initializer: ts.ForInitializer | undefined): string | null {
    if (!initializer) {
      return null;
    }

    if (ts.isVariableDeclarationList(initializer)) {
      const firstDecl = initializer.declarations[0];
      if (firstDecl && ts.isIdentifier(firstDecl.name)) {
        return firstDecl.name.text;
      }
      return null;
    }

    if (ts.isBinaryExpression(initializer) && initializer.operatorToken.kind === ts.SyntaxKind.EqualsToken) {
      if (ts.isIdentifier(initializer.left)) {
        return initializer.left.text;
      }
    }

    return null;
  }

  private resolveForLoopStartExpression(initializer: ts.ForInitializer | undefined): string {
    if (!initializer) {
      return '0';
    }

    if (ts.isVariableDeclarationList(initializer)) {
      const firstDecl = initializer.declarations[0];
      if (firstDecl?.initializer) {
        return this.emitExpression(firstDecl.initializer);
      }
      return '0';
    }

    if (ts.isBinaryExpression(initializer) && initializer.operatorToken.kind === ts.SyntaxKind.EqualsToken) {
      return this.emitExpression(initializer.right);
    }

    return '0';
  }

  private resolveForLoopEndExpression(condition: ts.BinaryExpression): string {
    const rhs = this.emitExpression(condition.right);
    if (condition.operatorToken.kind === ts.SyntaxKind.LessThanToken) {
      return `(${rhs}) - 1`;
    }

    return rhs;
  }

  private resolveForOfVariableName(initializer: ts.ForInitializer): string | null {
    if (ts.isVariableDeclarationList(initializer)) {
      const firstDecl = initializer.declarations[0];
      if (firstDecl && ts.isIdentifier(firstDecl.name)) {
        return firstDecl.name.text;
      }
    }
    return null;
  }

  private resolveForOfAssignmentTarget(initializer: ts.ForInitializer): string {
    if (ts.isVariableDeclarationList(initializer)) {
      const firstDecl = initializer.declarations[0];
      if (firstDecl) {
        return firstDecl.name.getText();
      }
      return 'ts_item';
    }

    return initializer.getText();
  }

  private shouldUseSetForForOfValue(iterableText: string, valueTarget: string): boolean {
    if (!this.currentProcedureObjectArrays) {
      return false;
    }

    if (!/^[_A-Za-z][_A-Za-z0-9]*$/.test(valueTarget)) {
      return false;
    }

    if (/^[_A-Za-z][_A-Za-z0-9]*$/.test(iterableText)) {
      return this.currentProcedureObjectArrays.has(iterableText);
    }

    return false;
  }

  private renderConsoleLogArgs(argumentsList: ts.NodeArray<ts.Expression>): string {
    if (!argumentsList.length) {
      return '""';
    }

    return argumentsList
      .map((arg) => this.emitExpression(arg))
      .join(' & " " & ');
  }

  private emitExpression(expression: ts.Expression): string {
    if (ts.isAsExpression(expression)) {
      return this.emitExpression(expression.expression);
    }

    if (ts.isParenthesizedExpression(expression)) {
      return this.emitExpression(expression.expression);
    }

    if (ts.isIdentifier(expression)) {
      return expression.text;
    }

    if (ts.isStringLiteral(expression) || ts.isNoSubstitutionTemplateLiteral(expression)) {
      return this.emitStringLiteralValue(expression.text);
    }

    if (ts.isNumericLiteral(expression)) {
      return expression.text;
    }

    if (ts.isArrayLiteralExpression(expression)) {
      const elements = expression.elements.map((e) => this.emitExpression(e as ts.Expression)).join(', ');
      return `Array(${elements})`;
    }

    if (ts.isTemplateExpression(expression)) {
      const parts: string[] = [];
      if (expression.head.text) {
        parts.push(this.emitStringLiteralValue(expression.head.text));
      }

      for (const span of expression.templateSpans) {
        parts.push(`CStr(${this.emitExpression(span.expression)})`);
        if (span.literal.text) {
          parts.push(this.emitStringLiteralValue(span.literal.text));
        }
      }

      return parts.join(' & ');
    }

    if (ts.isPrefixUnaryExpression(expression) && expression.operator === ts.SyntaxKind.ExclamationToken) {
      return `Not (${this.emitExpression(expression.operand)})`;
    }

    if (ts.isBinaryExpression(expression)) {
      const left = this.emitExpression(expression.left);
      const right = this.emitExpression(expression.right);
      const op = this.mapBinaryOperator(expression.operatorToken.kind);
      return `${left} ${op} ${right}`;
    }

    if (ts.isPropertyAccessExpression(expression)) {
      const ownerNode = this.unwrapExpression(expression.expression);
      if (ts.isIdentifier(ownerNode) && ownerNode.text === 'globalThis') {
        return expression.name.text;
      }

      if (expression.name.text === 'length') {
        const target = this.emitExpression(expression.expression);
        return `(UBound(${target}) - LBound(${target}) + 1)`;
      }

      const owner = this.emitExpression(expression.expression);

      return `${owner}.${expression.name.text}`;
    }

    if (ts.isCallExpression(expression) && ts.isPropertyAccessExpression(expression.expression)) {
      const method = expression.expression.name.text;
      if (method === 'includes' && expression.arguments.length >= 1) {
        const haystack = this.emitExpression(expression.expression.expression);
        const needle = this.emitExpression(expression.arguments[0]);
        return `(InStr(1, CStr(${haystack}), CStr(${needle})) > 0)`;
      }

      const callTarget = this.emitExpression(expression.expression.expression);
      const args = expression.arguments.map((arg) => this.emitExpression(arg)).join(', ');
      return `${callTarget}.${method}(${args})`;
    }

    if (ts.isCallExpression(expression) && ts.isIdentifier(expression.expression)) {
      const functionName = expression.expression.text;
      const args = expression.arguments.map((arg) => this.emitExpression(arg)).join(', ');
      return `${functionName}(${args})`;
    }

    if (ts.isElementAccessExpression(expression)) {
      return `${this.emitExpression(expression.expression)}(${this.emitExpression(expression.argumentExpression)})`;
    }

    if (ts.isNewExpression(expression) && ts.isIdentifier(expression.expression)) {
      const className = this.composeClassName(expression.expression.text);
      const args = (expression.arguments ?? []).map((arg) => this.emitExpression(arg)).join(', ');
      return `New_${className}(${args})`;
    }

    return expression.getText();
  }

  private emitCondition(expression: ts.Expression): string {
    if (ts.isAsExpression(expression)) {
      return this.emitCondition(expression.expression);
    }

    if (ts.isParenthesizedExpression(expression)) {
      return this.emitCondition(expression.expression);
    }

    if (ts.isPrefixUnaryExpression(expression) && expression.operator === ts.SyntaxKind.ExclamationToken) {
      return `Not (${this.emitCondition(expression.operand)})`;
    }

    if (ts.isBinaryExpression(expression)) {
      const opKind = expression.operatorToken.kind;
      if (opKind === ts.SyntaxKind.AmpersandAmpersandToken || opKind === ts.SyntaxKind.BarBarToken) {
        const joiner = opKind === ts.SyntaxKind.AmpersandAmpersandToken ? 'And' : 'Or';
        return `${this.emitCondition(expression.left)} ${joiner} ${this.emitCondition(expression.right)}`;
      }

      return this.emitExpression(expression);
    }

    if (ts.isCallExpression(expression) && ts.isPropertyAccessExpression(expression.expression)) {
      const method = expression.expression.name.text;
      if (method === 'includes') {
        return this.emitExpression(expression);
      }
    }

    if (ts.isIdentifier(expression)) {
      if (this.isObjectVariable(expression.text)) {
        return `Not ${expression.text} Is Nothing`;
      }

      return `${this.emitExpression(expression)} <> ""`;
    }

    if (ts.isPropertyAccessExpression(expression) || ts.isElementAccessExpression(expression)) {
      return `${this.emitExpression(expression)} <> ""`;
    }

    if (ts.isCallExpression(expression)) {
      return `${this.emitExpression(expression)} <> ""`;
    }

    return this.emitExpression(expression);
  }

  private needsSetAssignment(expression: ts.Expression): boolean {
    if (ts.isAsExpression(expression) || ts.isParenthesizedExpression(expression)) {
      return this.needsSetAssignment(expression.expression);
    }

    if (ts.isPropertyAccessExpression(expression)) {
      const objectLikeProps = new Set(['ActiveSheet', 'ThisWorkbook', 'ActiveWorkbook', 'Range', 'Cells']);
      return objectLikeProps.has(expression.name.text);
    }

    if (ts.isElementAccessExpression(expression)) {
      return false;
    }

    return ts.isNewExpression(expression);
  }

  private unwrapExpression(expression: ts.Expression): ts.Expression {
    let current = expression;
    while (ts.isAsExpression(current) || ts.isParenthesizedExpression(current)) {
      current = current.expression;
    }
    return current;
  }

  private registerProcedureVariable(name: string, vbaType: string): void {
    if (!this.currentProcedureVariables) {
      return;
    }

    if (this.currentProcedureParameters?.has(name)) {
      return;
    }

    const existingType = this.currentProcedureVariables.get(name);
    if (!existingType) {
      this.currentProcedureVariables.set(name, vbaType);
      return;
    }

    if (existingType !== 'Long' && vbaType === 'Long') {
      this.currentProcedureVariables.set(name, 'Long');
    }
  }

  private registerObjectVariable(name: string): void {
    if (!name) {
      return;
    }

    this.currentProcedureObjectVariables?.add(name);
    this.knownObjectVariables.add(name);
  }

  private isObjectVariable(name: string): boolean {
    return Boolean(this.currentProcedureObjectVariables?.has(name)) || this.knownObjectVariables.has(name);
  }

  private isLikelyObjectDeclaration(node: ts.VariableDeclaration): boolean {
    if (node.initializer) {
      if (node.initializer.kind === ts.SyntaxKind.NullKeyword) {
        return true;
      }

      if (ts.isNewExpression(node.initializer)) {
        return true;
      }
    }

    if (!node.type) {
      return false;
    }

    switch (node.type.kind) {
      case ts.SyntaxKind.StringKeyword:
      case ts.SyntaxKind.NumberKeyword:
      case ts.SyntaxKind.BooleanKeyword:
      case ts.SyntaxKind.BigIntKeyword:
        return false;
      default:
        return true;
    }
  }

  private emitStringLiteralValue(value: string): string {
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

  private emitClassFactory(classDecl: ts.ClassDeclaration): void {
    if (!classDecl.name) {
      return;
    }

    const className = this.composeClassName(classDecl.name.text);
    const ctor = classDecl.members.find(
      (member): member is ts.ConstructorDeclaration => ts.isConstructorDeclaration(member),
    );
    const params = ctor?.parameters ?? [];
    const paramSignature = params
      .map((p) => `ByVal ${p.name.getText()} As Variant`)
      .join(', ');
    const argPass = params.map((p) => p.name.getText()).join(', ');

    this.writeLine(`Public Function New_${className}(${paramSignature}) As ${className}`);
    this.indentLevel++;
    this.writeLine(`Dim obj As New ${className}`);
    this.writeLine(`Call obj.Init(${argPass})`);
    this.writeLine(`Set New_${className} = obj`);
    this.indentLevel--;
    this.writeLine('End Function');
    this.writeLine();
  }

  private composeClassName(originalName: string): string {
    return `${this.options.namespacePrefix}${this.options.moduleName}_${originalName}`;
  }

  private mapBinaryOperator(kind: ts.SyntaxKind): string {
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

  private composeFunctionName(originalName: string): string {
    return `${this.options.namespacePrefix}${this.options.moduleName}_${originalName}`;
  }
}
