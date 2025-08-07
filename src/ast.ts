import { intToColChars } from './util';

export namespace AST {
  export class Env {
    static readonly type: 'Env' = 'Env';
    readonly tag = Env.type;
    readonly path: string;
    readonly workbookName: string;
    readonly worksheetName: string;

    constructor(path: string, workbookName: string, worksheetName: string) {
      this.path = path;
      this.workbookName = workbookName;
      this.worksheetName = worksheetName;
    }

    equals(e: Env): boolean {
      return this.path === e.path && this.workbookName === e.workbookName && this.worksheetName === e.worksheetName;
    }

    isEmpty() {
      return this.worksheetName === '' && this.workbookName === '' && this.path === '';
    }
  }

  export type AddressMode = 'absolute' | 'relative';
  const singleQuoteRe = /'/g;

  const getEnvPrefix = (env: Env): string => {
    if (env.isEmpty()) {
      return '';
    }

    const { worksheetName, workbookName, path } = env;
    let prefix = '';
    if (path) {
      prefix += path;
    }

    if (workbookName) {
      prefix += '[' + workbookName + ']';
    }

    if (worksheetName) {
      prefix += worksheetName.replace(singleQuoteRe, `''`);
    }

    return `'${prefix}'!`;
  };

  const stringifyR1C1Item = (char: 'R' | 'C', value: number, mode: AddressMode): string => {
    let s = char;
    const relative = mode === 'relative';

    if (value !== 0) {
      if (relative) {
        s += '[';
      }
      s += value.toString();
      if (relative) {
        s += ']';
      }
    }

    return s;
  };

  const stringifyR1C1 = (address: Address): string =>
    stringifyR1C1Item('R', address.row, address.rowMode) + stringifyR1C1Item('C', address.column, address.colMode);

  const getA1Prefix = (mode: AddressMode): string => (mode === 'absolute' ? '$' : '');

  const stringifyA1Item = (value: string | number, mode: AddressMode) => getA1Prefix(mode) + value.toString();

  const stringifyA1 = (address: Address): string =>
    stringifyA1Item(intToColChars(address.column), address.colMode) + stringifyA1Item(address.row, address.rowMode);

  export class R1C1Pair {
    readonly type = 'R1C1';
    readonly value: number;
    readonly mode: AddressMode;

    constructor();

    constructor(value: number, mode: AddressMode);

    constructor(value?: number, mode?: AddressMode) {
      this.value = value === undefined ? 0 : value;
      this.mode = mode === undefined ? 'relative' : mode;
    }
  }

  export class A1Pair {
    readonly type = 'A1';
    readonly value: number;
    readonly mode: AddressMode;

    constructor(value: number, mode: AddressMode) {
      this.value = value === undefined ? 0 : value;
      this.mode = mode === undefined ? 'relative' : mode;
    }
  }

  type AddressType = 'r1c1' | 'a1';

  export class Address {
    readonly row: number;
    readonly column: number;
    readonly rowMode: AddressMode;
    readonly colMode: AddressMode;
    readonly env: Env;

    #type: AddressType;

    static r1c1(row: R1C1Pair, column: R1C1Pair, env: Env): Address {
      return new Address(row.value, column.value, row.mode, column.mode, env, 'r1c1');
    }

    static a1(row: A1Pair, column: A1Pair, env: Env): Address {
      return new Address(row.value, column.value, row.mode, column.mode, env, 'a1');
    }

    private constructor(
      row: number,
      column: number,
      rowMode: AddressMode,
      colMode: AddressMode,
      env: Env,
      type: AddressType
    ) {
      this.row = row;
      this.column = column;
      this.rowMode = rowMode;
      this.colMode = colMode;
      this.env = env;
      this.#type = type;
    }

    toString(): string {
      return '(' + this.column.toString() + ',' + this.row.toString() + ')';
    }

    toFormula(): string {
      return this.#type === 'r1c1' ? this.#toR1C1Ref() : this.#toA1Ref();
    }

    /**
     * Returns a copy of this Address but with an updated Env.
     * @param env An Env object.
     * @returns An Address.
     */
    copyWithNewEnv(env: Env) {
      return new Address(this.row, this.column, this.rowMode, this.colMode, env, this.#type);
    }

    #toA1Ref(): string {
      return getEnvPrefix(this.env) + stringifyA1(this);
    }

    #toR1C1Ref(): string {
      return getEnvPrefix(this.env) + stringifyR1C1(this);
    }
  }

  type Region = readonly [Address, Address];

  export class Range {
    static readonly type: 'Range' = 'Range';
    readonly type = Range.type;
    readonly regions: readonly Region[];

    constructor(regions: readonly Region[]) {
      this.regions = regions;
    }

    /**
     * Merge this range and another range into a discontiguous range.
     * @param r The other range.
     * @returns A discontiguous range.
     */
    merge(r: Range): Range {
      return new Range(this.regions.concat(r.regions));
    }

    /**
     * Returns a copy of this Range but with an updated Env.
     * @param env An Env object.
     * @returns A Range.
     */
    copyWithNewEnv(env: Env): Range {
      return new Range(this.regions.map(([tl, br]) => [tl.copyWithNewEnv(env), br.copyWithNewEnv(env)]));
    }

    toString(): string {
      const sregs = this.regions.map(([tl, br]) => tl.toString() + ':' + br.toString());
      return 'List(' + sregs.join(',') + ')';
    }

    toFormula(): string {
      return this.regions.map(([tl, br]) => tl.toFormula() + ':' + br.toFormula()).join(',');
    }
  }

  export interface BaseExpression {
    /**
     * Returns the type tag for the expression subtype,
     * for use in pattern-matching expressions. Also
     * available as a static property on ReferenceExpr
     * types.
     */
    readonly type: string;

    /**
     * Generates a valid Excel formula from this expression.
     */
    toFormula(): string;

    /**
     * Pretty-prints the AST as a string.  Note that this
     * does not produce a valid Excel formula.
     */
    toString(): string;
  }

  export class ReferenceRange implements BaseExpression {
    static readonly type: 'ReferenceRange' = 'ReferenceRange';
    readonly type = ReferenceRange.type;
    readonly range: Range;

    constructor(env: Env, range: Range) {
      this.range = range.copyWithNewEnv(env);
    }

    toFormula(): string {
      return this.range.toFormula();
    }

    toString(): string {
      return 'ReferenceRange(' + this.range.toString() + ')';
    }
  }

  export class ReferenceAddress implements BaseExpression {
    static readonly type: 'ReferenceAddress' = 'ReferenceAddress';
    readonly type = ReferenceAddress.type;
    readonly address: Address;

    constructor(env: Env, address: Address) {
      this.address = address.copyWithNewEnv(env);
    }

    toString(): string {
      return 'ReferenceAddress(' + this.address.toString() + ')';
    }

    toFormula(): string {
      return this.address.toFormula();
    }
  }

  export class ReferenceNamed implements BaseExpression {
    static readonly type: 'ReferenceNamed' = 'ReferenceNamed';
    readonly type = ReferenceNamed.type;
    readonly varName: string;

    constructor(varName: string) {
      this.varName = varName;
    }

    toString(): string {
      return 'ReferenceName(' + this.varName + ')';
    }

    toFormula(): string {
      return this.varName;
    }
  }

  export class FixedArity {
    static readonly type: 'FixedArity' = 'FixedArity';
    readonly num: number;

    constructor(num: number) {
      this.num = num;
    }
  }

  export class LowBoundArity {
    static readonly type: 'LowBoundArity' = 'LowBoundArity';
    readonly num: number;

    constructor(num: number) {
      this.num = num;
    }
  }

  class VarArgsArity {
    static readonly type: 'VarArgsArity' = 'VarArgsArity';
  }

  export const VarArgsArityInst = new VarArgsArity();

  export type Arity = FixedArity | LowBoundArity | VarArgsArity;

  export class FunctionApplication implements BaseExpression {
    static readonly type: 'FunctionApplication' = 'FunctionApplication';
    readonly type = FunctionApplication.type;
    readonly name: string;
    readonly args: readonly Expression[];
    readonly arity: Arity;

    constructor(name: string, args: readonly Expression[], arity: Arity) {
      this.name = name;
      this.args = args;
      this.arity = arity;
    }

    toString(): string {
      return 'Function[' + this.name + ',' + this.arity + '](' + this.args.map(arg => arg.toFormula).join(',') + ')';
    }

    toFormula(): string {
      return this.name + '(' + this.args.map(arg => arg.toFormula()).join(',') + ')';
    }
  }

  export class NumberLiteral implements BaseExpression {
    static readonly type: 'NumberLiteral' = 'NumberLiteral';
    readonly type = NumberLiteral.type;
    readonly value: number;

    constructor(value: number) {
      this.value = value;
    }

    toString(): string {
      return 'Number(' + this.value + ')';
    }

    toFormula(): string {
      return this.value.toString();
    }
  }

  export class StringLiteral implements BaseExpression {
    static readonly type: 'StringLiteral' = 'StringLiteral';
    readonly type = StringLiteral.type;
    readonly value: string;

    constructor(value: string) {
      this.value = value;
    }

    toString(): string {
      return 'String(' + this.value + ')';
    }

    toFormula(): string {
      return '"' + this.value + '"';
    }
  }

  export class BooleanLiteral implements BaseExpression {
    static readonly type: 'BooleanLiteral' = 'BooleanLiteral';
    readonly type = BooleanLiteral.type;
    readonly value: boolean;

    constructor(value: boolean) {
      this.value = value;
    }

    toString(): string {
      return 'Boolean(' + this.value + ')';
    }

    toFormula(): string {
      return this.value.toString().toUpperCase();
    }
  }

  // this should only ever be instantiated by
  // the reserved words class, which is designed
  // to fail
  export class PoisonPill implements BaseExpression {
    static readonly type: 'PoisonPill' = 'PoisonPill';
    readonly type = PoisonPill.type;

    toString(): string {
      throw new Error('This object should never appear in an AST.');
    }

    toFormula(): string {
      throw new Error('This object should never appear in an AST.');
    }
  }

  export class ParensExpr implements BaseExpression {
    static readonly type: 'ParensExpr' = 'ParensExpr';
    readonly type = ParensExpr.type;
    readonly expr: Expression;

    constructor(expr: Expression) {
      this.expr = expr;
    }

    toString(): string {
      return 'Parens(' + this.expr.toString() + ')';
    }

    toFormula(): string {
      return '(' + this.expr.toFormula() + ')';
    }
  }

  export class BinOpExpr implements BaseExpression {
    static readonly type: 'BinOpExpr' = 'BinOpExpr';
    readonly type = BinOpExpr.type;
    readonly op: string;
    readonly exprL: Expression;
    readonly exprR: Expression;

    constructor(op: string, exprL: Expression, exprR: Expression) {
      this.op = op;
      this.exprR = exprR;
      this.exprL = exprL;
    }

    toFormula(): string {
      return this.exprL.toFormula() + ' ' + this.op + ' ' + this.exprR.toFormula();
    }

    toString(): string {
      return 'BinOpExpr(' + this.op.toString() + ',' + this.exprL.toFormula() + ',' + this.exprR.toFormula() + ')';
    }
  }

  export class UnaryOpExpr implements BaseExpression {
    static readonly type: 'UnaryOpExpr' = 'UnaryOpExpr';
    readonly type = UnaryOpExpr.type;
    readonly op: string;
    readonly expr: Expression;

    constructor(op: string, expr: Expression) {
      this.op = op;
      this.expr = expr;
    }

    toFormula(): string {
      return this.op + this.expr.toFormula();
    }

    toString(): string {
      return 'UnaryOpExpr(' + this.op.toString() + ',' + this.expr.toFormula() + ')';
    }
  }

  export type Expression =
    | ReferenceRange
    | ReferenceAddress
    | ReferenceNamed
    | FunctionApplication
    | NumberLiteral
    | StringLiteral
    | BooleanLiteral
    | BinOpExpr
    | UnaryOpExpr
    | ParensExpr;
}
