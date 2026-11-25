// A focused guard for aliasing patterns seen in Excel, not an Office Scripts compiler.
const fs = require("node:fs");
const path = require("node:path");

const [compiler, declaration, ...arguments_] = process.argv.slice(2);
const strict = arguments_.includes("--strict-office");
const files = arguments_.filter(argument => argument !== "--strict-office");
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
  // Strict mode includes ExcelScript data contracts, including type aliases.
  // A local shape is followed through its fields and element types. Host symbols
  // are terminal so traversing a Range does not also visit all its methods.
  function apiTypes(type, at, seen = new Set()) {
    if (seen.has(type)) return [];
    seen.add(type);
    for (const symbol of [type.aliasSymbol, type.getSymbol()]) {
      if (symbol?.declarations?.some(item => item.getSourceFile().fileName === declaration &&
          (ts.isInterfaceDeclaration(item) || ts.isTypeAliasDeclaration(item)))) {
        return [symbol.getName()];
      }
    }
    if (type.isUnionOrIntersection()) return type.types.flatMap(item => apiTypes(item, at, seen));
    if (checker.isArrayType(type) || checker.isTupleType(type)) {
      return checker.getTypeArguments(type).flatMap(item => apiTypes(item, at, seen));
    }
    if ((type.flags & ts.TypeFlags.Object) === 0) return [];
    const fields = checker.getPropertiesOfType(type).filter(field =>
      field.declarations?.some(item => item.getSourceFile() === source));
    const index = checker.getIndexTypeOfType(type, ts.IndexKind.String);
    return [
      ...fields.flatMap(field => apiTypes(checker.getTypeOfSymbolAtLocation(field, at), at, seen)),
      ...(index ? apiTypes(index, at, seen) : []),
    ];
  }
  const native = node => {
    const type = checker.getTypeAtLocation(node);
    return strict ? apiTypes(type, node) : handles(type);
  };
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
    if (ts.isParameter(node) && native(node).length && !root(node.name) &&
        !(strict && !node.type && ts.isArrowFunction(node.parent))) issue(node, "api-parameter");
    if (ts.isReturnStatement(node) && node.expression && native(node.expression).length) {
      issue(node, "api-return");
    }
    if (ts.isVariableDeclaration(node) && native(node.name).length &&
        (!node.initializer || !directApi(node.initializer))) issue(node, "api-alias");
    if (strict && ts.isCallExpression(node) && directApi(node)) {
      for (const argument of node.arguments) {
        if (native(argument).length &&
            (ts.isPropertyAccessExpression(argument) || ts.isElementAccessExpression(argument))) {
          issue(argument, "api-property-argument");
        }
      }
    }
    if (strict && node.kind === ts.SyntaxKind.AnyKeyword) issue(node, "explicit-any");
    if (strict && ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression) &&
        checker.isArrayType(checker.getTypeAtLocation(node.expression.expression)) &&
        ["map", "filter", "forEach", "find", "findIndex", "some", "every", "sort", "reduce", "flatMap"]
          .includes(node.expression.name.text) && node.arguments.length &&
        !ts.isArrowFunction(node.arguments[0])) issue(node.arguments[0], "array-callback");
    if (strict && ts.isCallExpression(node) && ts.isPropertyAccessExpression(node.expression) &&
        ["push", "unshift", "splice"].includes(node.expression.name.text) &&
        node.arguments.some(argument => native(argument).length)) issue(node, "api-array-write");
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
