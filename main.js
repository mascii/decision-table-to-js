const DC_INPUT_STRING = "don't care"; // 入力配列内で Don't Care を表す文字列
const DONT_CARE_SYMBOL = Symbol("DON'T CARE"); // 内部処理用のシンボル

class Leaf {
  /**
   * @param {string | typeof DONT_CARE_SYMBOL} value
   */
  constructor(value) {
    this.value = value;
  }
}

class Branch {
  /**
   * @param {number} varIndex
   * @param {Leaf | Branch} high
   * @param {Leaf | Branch} low
   */
  constructor(varIndex, high, low) {
    this.varIndex = varIndex;
    this.high = high;
    this.low = low;
  }
}

/**
 * ★最重要: 評価関数 (Cost Function)
 * コード生成器が `if (A && B)` のように結合することをシミュレートし、
 * 最終的に出力される `return` 文の数を計算する。
 *
 * @param {Leaf | Branch} node
 * @returns {number}
 */
function calculateOptimizedReturnCount(node) {
  // 葉ノードなら return 文は1つ
  if (node instanceof Leaf) {
    return 1;
  }

  // 分岐ノードの場合、&& 結合できるかチェック
  // 条件:
  // 1. 自分のElse(low)が葉である
  // 2. 子供(high)も分岐であり、そのElse(child.low)が自分のElseと同じ値である
  const canMerge =
    node.low instanceof Leaf &&
    node.high instanceof Branch &&
    node.high.low instanceof Leaf &&
    node.high.low.value === node.low.value;
  if (canMerge) {
    // 結合できる場合 (例: if (A && B) ... else X)
    // 親のElse(X)は子供のElse(X)と共通化されるため、
    // ここでコストを加算せず、子供(High)のコストだけを返す
    return calculateOptimizedReturnCount(node.high);
  }

  // 結合できない場合
  // High側のリターン数 + Low側のリターン数
  return (
    calculateOptimizedReturnCount(node.high) +
    calculateOptimizedReturnCount(node.low)
  );
}

/**
 * ノードの等価性チェック
 * @param {Leaf | Branch} n1
 * @param {Leaf | Branch} n2
 * @returns {boolean}
 */
function isNodesEqual(n1, n2) {
  if (n1 instanceof Leaf && n2 instanceof Leaf) {
    return n1.value === n2.value;
  }
  if (n1 instanceof Branch && n2 instanceof Branch) {
    return (
      n1.varIndex === n2.varIndex &&
      isNodesEqual(n1.high, n2.high) &&
      isNodesEqual(n1.low, n2.low)
    );
  }
  return false;
}

/**
 * MTBDDの簡約化（Reduction）処理
 * @param {number} varIndex - 現在の分岐が評価する変数のインデックス
 * @param {Leaf | Branch} high - 変数が True の場合の子ノード
 * @param {Leaf | Branch} low - 変数が False の場合の子ノード
 * @returns {Leaf | Branch} 簡約化されたノード
 */
function reduce(varIndex, high, low) {
  const isHighDC = high instanceof Leaf && high.value === DONT_CARE_SYMBOL;
  const isLowDC = low instanceof Leaf && low.value === DONT_CARE_SYMBOL;

  // Don't Care (DC) の吸収: 両方がDCの場合、DCを返します。
  if (isHighDC && isLowDC) return high;
  // Don't Care (DC) の吸収: DCではない側のノードを返します（DCは何でも良いため、分岐を減らす方に合わせます）。
  if (isHighDC) return low;
  if (isLowDC) return high;
  // 冗長な分岐の削除: HighとLowが全く同じ構造を持つ場合、分岐は不要なため一方の子ノードを返します。
  if (isNodesEqual(high, low)) return high;

  // 統合不可な場合: 新しい Branch ノードを作成して返します。
  return new Branch(varIndex, high, low);
}

/**
 * 再帰的な木の構築 (変数順序対応版)
 * @param {(string | typeof DONT_CARE_SYMBOL)[]} outputs
 * @param {number[]} order - 変数の評価順序
 * @param {number} depth
 * @returns {Leaf | Branch} - 構築されたMTBDDのルートノード
 */
function buildMTBDD(outputs, order, depth = 0) {
  if (outputs.length === 1) {
    return new Leaf(outputs[0]);
  }

  const mid = outputs.length / 2;
  const currentVarIndex = order[depth]; // 現在の深さで評価すべき変数ID

  const highPart = outputs.slice(0, mid);
  const lowPart = outputs.slice(mid);

  const highNode = buildMTBDD(highPart, order, depth + 1);
  const lowNode = buildMTBDD(lowPart, order, depth + 1);

  return reduce(currentVarIndex, highNode, lowNode);
}

/**
 * テーブルを指定された変数順序(order)に従って並べ替える
 * @param {(string | typeof DONT_CARE_SYMBOL)[]} originalTable
 * @param {number} n
 * @param {number[]} order
 * @returns {(string | typeof DONT_CARE_SYMBOL)[]}
 */
function reorderTable(originalTable, n, order) {
  return Array.from({ length: originalTable.length }, (_, i) => {
    let originalIdx = 0;
    // i は「並べ替え後のツリー上のパス」 (0=All False, 1=...True)
    for (let depth = 0; depth < n; depth++) {
      // i の (n-1-depth) ビット目は、変数 order[depth] の値
      const bitVal = (i >> (n - 1 - depth)) & 1;

      if (bitVal === 1) {
        // 元のインデックス上の正しい位置にビットを立てる
        originalIdx |= 1 << (n - 1 - order[depth]);
      }
    }
    return originalTable[originalIdx];
  });
}

/**
 * 0..n-1 の順列生成 (Generator)
 * @param {number} n
 */
function* permute(n) {
  const permutation = Array.from({ length: n }, (_, i) => i);
  /** @type {number[]} */
  const c = Array(n).fill(0);
  let i = 1;

  yield permutation.slice();
  while (i < n) {
    if (c[i] < i) {
      const k = i % 2 && c[i];
      [permutation[i], permutation[k]] = [permutation[k], permutation[i]];
      c[i]++;
      i = 1;
      yield permutation.slice();
    } else {
      c[i] = 0;
      i++;
    }
  }
}

/**
 * 全順列のMTBDDを生成・分析する
 * @param {string[]} decisionTable
 * @returns {Array<{ root: Leaf | Branch, order: number[], id: number, score: number }>}
 */
function analyzeAllMTBDDs(decisionTable) {
  const length = decisionTable.length;

  // バリデーション
  if (length === 0 || (length & (length - 1)) !== 0) {
    throw new Error(
      `配列の長さ(${length})は 2のn乗 (2, 4, 8, 16...) である必要があります。`,
    );
  }

  const n = Math.log2(length);

  // 入力データの変換
  const inputs = decisionTable.map((val) =>
    val === DC_INPUT_STRING ? DONT_CARE_SYMBOL : val,
  );

  const results = [];
  let id = 1;

  // 全順列を総当たり
  for (const order of permute(n)) {
    const table = reorderTable(inputs, n, order);
    const root = buildMTBDD(table, order, 0);
    const score = calculateOptimizedReturnCount(root);

    results.push({
      id: id++,
      order,
      root,
      score,
    });
  }

  return results;
}

/**
 * 全順列を生成し、最小スコア（リターン数）を持つMTBDDのみを抽出する
 * @param {string[]} decisionTable
 * @returns {Array<{ root: Leaf | Branch, order: number[], id: number, score: number }>}
 */
function getOptimalMTBDDs(decisionTable) {
  const allMTBDDs = analyzeAllMTBDDs(decisionTable);
  const minScore = Math.min(...allMTBDDs.map((r) => r.score));
  return allMTBDDs.filter((r) => r.score === minScore);
}

/**
 * @param {any[]} range - Google Spreadsheets の範囲指定(決定表の出力部分)
 * @param {string[]} names - 関数名と引数名
 * @returns {string[]} - 生成されたJavaScriptコードの配列
 * @customfunction
 */
function generateJS(range, ...names) {
  const [funcName = "decideLogic", ...argNames] = names;
  const decisionTable = range.flatMap(String);
  const bestMTBDDs = getOptimalMTBDDs(decisionTable);

  /** @type {Map<string, number>} */
  const uniqueCodeMap = new Map();
  /** @type {string[]} */
  const results = [];

  for (const { root, id } of bestMTBDDs) {
    /**
     * 文法を統一した表示関数 (再帰)
     * @param {Leaf | Branch} node
     * @param {number} indent
     * @returns {string}
     */
    const generateBody = (node, indent = 0) => {
      const spaces = "  ".repeat(indent);

      if (node instanceof Leaf) {
        const valStr =
          node.value === DONT_CARE_SYMBOL ? "null" : `'${node.value}'`;
        return `${spaces}return ${valStr};\n`;
      }

      const getVarName = (/** @type {number} */ idx) => {
        return argNames[idx] ?? `input[${idx}]`;
      };

      if (node.low instanceof Leaf) {
        const fallbackValue = node.low.value;
        /** @type {{index: number, name: string}[]} */
        const conditions = [];
        /** @type {Leaf | Branch} */
        let current = node;

        while (
          current instanceof Branch &&
          current.low instanceof Leaf &&
          current.low.value === fallbackValue
        ) {
          conditions.push({
            index: current.varIndex,
            name: getVarName(current.varIndex),
          });
          current = current.high;
        }

        if (conditions.length > 0) {
          // ★varIndex でソートして正規化 (A && B == B && A)
          conditions.sort((a, b) => a.index - b.index);
          const condStr = conditions.map((c) => c.name).join(" && ");
          let buffer = "";

          buffer += `${spaces}if (${condStr}) {\n`;
          buffer += generateBody(current, indent + 1);
          buffer += `${spaces}}\n`;
          const elseStr =
            fallbackValue === DONT_CARE_SYMBOL ? "null" : `'${fallbackValue}'`;
          buffer += `${spaces}return ${elseStr};\n`;
          return buffer;
        }
      }

      let buffer = "";
      buffer += `${spaces}if (${getVarName(node.varIndex)}) {\n`;
      buffer += generateBody(node.high, indent + 1);
      buffer += `${spaces}}\n`;
      buffer += generateBody(node.low, indent);
      return buffer;
    };

    const body = generateBody(root, 1);

    // ヘッダ以外の部分で重複チェックを行う
    if (!uniqueCodeMap.has(body)) {
      uniqueCodeMap.set(body, id);

      const args = argNames.length > 0 ? argNames.join(", ") : "...input";
      const fullCode = `// ID: ${id}
function ${funcName}(${args}) {
${body}}`;
      results.push(fullCode);
    }
  }

  return results;
}

/**
 * @param {any[]} range - Google Spreadsheets の範囲指定(決定表の出力部分)
 * @param {string[]} names - 関数名と引数名
 * @returns {string[]} - 生成されたMermaidフローチャートの配列
 * @customfunction
 */
function generateMermaid(range, ...names) {
  const [funcName = "decideLogic", ...argNames] = names;
  const decisionTable = range.flatMap(String);
  const bestMTBDDs = getOptimalMTBDDs(decisionTable);

  /** @type {Map<string, number>} */
  const uniqueGraphMap = new Map();
  /** @type {string[]} */
  const results = [];

  for (const { root: node, id } of bestMTBDDs) {
    const lines = ["graph TD"];
    const nodeMap = new Map();
    let idCounter = 0;

    function getId(n) {
      if (!nodeMap.has(n)) {
        nodeMap.set(n, `N${idCounter++}`);
      }
      return nodeMap.get(n);
    }

    const rootId = getId(node);
    const startLabel = funcName || "Start";
    lines.push(`    Start(["${startLabel}"]) --> ${rootId}`);

    const visited = new Set();

    function traverse(n) {
      if (visited.has(n)) return;
      visited.add(n);

      const nodeId = getId(n);

      if (n instanceof Leaf) {
        const valStr = n.value === DONT_CARE_SYMBOL ? "null" : `"${n.value}"`;
        lines.push(`    ${nodeId}[${valStr}]`);
      } else {
        let merged = false;
        if (n.low instanceof Leaf) {
          const fallbackValue = n.low.value;
          const conditions = [];
          let current = n;

          while (
            current instanceof Branch &&
            current.low instanceof Leaf &&
            current.low.value === fallbackValue
          ) {
            conditions.push({
              index: current.varIndex,
              name: argNames[current.varIndex] ?? `input[${current.varIndex}]`,
            });
            current = current.high;
          }

          if (conditions.length > 1) {
            merged = true;
            // ★varIndex でソートして正規化
            conditions.sort((a, b) => a.index - b.index);
            const label = conditions.map((c) => c.name).join(" && ");
            lines.push(`    ${nodeId}{"${label}"}`);

            const highId = getId(current);
            traverse(current);
            lines.push(`    ${nodeId} -->|True| ${highId}`);

            const lowId = getId(n.low);
            traverse(n.low);
            lines.push(`    ${nodeId} -.->|False| ${lowId}`);
          }
        }

        if (!merged) {
          const varName = argNames[n.varIndex] ?? `input[${n.varIndex}]`;
          lines.push(`    ${nodeId}{"${varName}"}`);

          const highId = getId(n.high);
          traverse(n.high);
          lines.push(`    ${nodeId} -->|True| ${highId}`);

          const lowId = getId(n.low);
          traverse(n.low);
          lines.push(`    ${nodeId} -.->|False| ${lowId}`);
        }
      }
    }

    traverse(node);
    const body = lines.join("\n");

    if (!uniqueGraphMap.has(body)) {
      uniqueGraphMap.set(body, id);
      const fullCode = `%% ID: ${id}\n${body}`;
      results.push(fullCode);
    }
  }

  return results;
}

/**
 * 指定された入力数に基づき、すべてのブール値の組み合わせを持つ決定表の入力部分を生成します。
 * スプレッドシート上で =decisionTable(4) のように使用することを想定しています。
 *
 * @param {number} n - 入力（変数）の数。生成される行数は 2^n となります
 * @param {any} [trueValue=true] - 真を表す値（任意）
 * @param {any} [falseValue=false] - 偽を表す値（任意）
 * @returns {any[][]} すべての組み合わせを含む 2次元配列
 * @customfunction
 */
function decisionTable(n, trueValue = true, falseValue = false) {
  if (n <= 0) return [[]];

  const rowCount = 2 ** n;
  const table = [];

  for (let i = 0; i < rowCount; i++) {
    const row = [];
    for (let j = n - 1; j >= 0; j--) {
      row.push(!((i >> j) & 1) ? trueValue : falseValue);
    }
    table.push(row);
  }

  return table;
}
