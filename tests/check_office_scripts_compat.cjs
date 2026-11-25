// A focused guard for aliasing patterns seen in Excel, not an Office Scripts compiler.
const fs = require("node:fs");
const path = require("node:path");

const [compiler, declaration, ...files] = process.argv.slice(2);
const compilerDir = path.dirname(fs.realpathSync(compiler));
const ts = require(require.resolve("typescript", { paths: [compilerDir, path.dirname(compiler)] }));
const issues = [];

for (const file of files) {
  const program = ts.createProgram([declaration, file], {
    strict: true, target: ts.ScriptTarget.ES2019, noEmit: true,
  });
  const checker = program.getTypeChecker();
  const source = program.getSourceFile(file);

  function handles(type) {
    if (type.isUnionOrIntersection()) return type.types.flatMap(handles);
    if (checker.isArrayType(type)) return checker.getTypeArguments(type).flatMap(handles);
    const symbol = type.getSymbol();
    if (!symbol) return [];
    if (symbol.declarations?.some(item => ts.isInterfaceDeclaration(item) &&
        item.getSourceFile().fileName === declaration)) {
      return [symbol.getName()];
    }
    return [];
  }
  const native = node => handles(checker.getTypeAtLocation(node));
  const root = node => ts.isIdentifier(node) && node.text === "workbook" &&
    native(node).every(type => type === "Workbook");
  const directApi = expression => ts.isCallExpression(expression) &&
    checker.getResolvedSignature(expression)?.declaration?.getSourceFile().fileName === declaration;
  function issue(node, rule) {
    const { line, character } = source.getLineAndCharacterOfPosition(node.getStart(source));
    issues.push({ file, line: line + 1, column: character + 1, rule });
  }
  function walk(node) {
    if (ts.isPropertySignature(node) && native(node).length) issue(node, "api-in-data-type");
    if ((ts.isPropertyAssignment(node) && native(node.initializer).length) ||
        (ts.isShorthandPropertyAssignment(node) && native(node.name).length) ||
        (ts.isArrayLiteralExpression(node) && node.elements.some(item => native(item).length))) {
      issue(node, "api-in-data");
    }
    if (ts.isParameter(node) && native(node).length && !root(node.name)) issue(node, "api-parameter");
    if (ts.isReturnStatement(node) && node.expression && native(node.expression).length) {
      issue(node, "api-return");
    }
    if (ts.isVariableDeclaration(node) && native(node.name).length &&
        (!node.initializer || !directApi(node.initializer))) issue(node, "api-alias");
    if (ts.isConditionalExpression(node) && native(node).length) issue(node, "api-conditional");
    if (ts.isBinaryExpression(node) && node.operatorToken.kind === ts.SyntaxKind.EqualsToken &&
        native(node.right).length) issue(node, "api-assignment");
    if (ts.isCallExpression(node) &&
        checker.getResolvedSignature(node)?.declaration?.getSourceFile() === source &&
        node.arguments.some(argument => native(argument).length && !root(argument))) {
      issue(node, "api-helper-argument");
    }
    ts.forEachChild(node, walk);
  }
  walk(source);
}

console.log(JSON.stringify(issues, null, 2));
if (issues.length) process.exitCode = 1;
