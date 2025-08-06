export namespace AST {
  const intToColChars = (dividend: number): string => {
    let quot = Math.floor(dividend / 26);
    const rem = dividend % 26;
    if (rem === 0) {
      quot -= 1;
    }
    const ltr = rem === 0 ? 'Z' : String.fromCharCode(64 + rem);
    return quot === 0 ? ltr : intToColChars(quot) + ltr;
  };

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

  export interface AbsoluteAddressMode {
    readonly type: 'AbsoluteAddress';
  }

  export interface RelativeAddressMode {
    readonly type: 'RelativeAddress';
  }

  export const AbsoluteAddress: AbsoluteAddressMode = {
    type: 'AbsoluteAddress'
  };

  export const RelativeAddress: RelativeAddressMode = {
    type: 'RelativeAddress'
  };

  export type AddressMode = AbsoluteAddressMode | RelativeAddressMode;

  const singleQuoteRe = /'/g;
  const getA1Prefix = (mode: AddressMode): string => (mode.type === 'AbsoluteAddress' ? '$' : '');

  export class Address {
    static readonly type: 'Address' = 'Address';
    readonly type = Address.type;
    readonly row: number;
    readonly column: number;
    readonly rowMode: AddressMode;
    readonly colMode: AddressMode;
    readonly env: Env;

    constructor(row: number, column: number, rowMode: AddressMode, colMode: AddressMode, env: Env) {
      this.row = row;
      this.column = column;
      this.rowMode = rowMode;
      this.colMode = colMode;
      this.env = env;
    }

    get path() {
      return this.env.path;
    }

    get workbookName() {
      return this.env.workbookName;
    }

    get worksheetName() {
      return this.env.worksheetName;
    }

    equals(a: Address) {
      // note that we explicitly do not compare paths since two different workbooks
      // can have the "same address" but will live at different paths
      return (
        this.row === a.row &&
        this.column === a.column &&
        this.worksheetName === a.worksheetName &&
        this.workbookName === a.workbookName
      );
    }

    toString(): string {
      return '(' + this.column.toString() + ',' + this.row.toString() + ')';
    }

    /**
     * Pretty-prints an address in A1 format.
     */
    toFormula(r1c1: boolean = false): string {
      return r1c1 ? this.toR1C1Ref() : this.toA1Ref();
    }

    /**
     * Returns a copy of this Address but with an updated Env.
     * @param env An Env object.
     * @returns An Address.
     */
    copyWithNewEnv(env: Env) {
      return new Address(this.row, this.column, this.rowMode, this.colMode, env);
    }

    #toEnvPrefix(): string {
      if (this.env.isEmpty()) {
        return '';
      }

      const { worksheetName, workbookName, path } = this.env;
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
    }

    toA1Ref(): string {
      return this.#toEnvPrefix() + this.toA1RefOnly();
    }

    toR1C1Ref(): string {
      return this.#toEnvPrefix() + this.toR1C1RefOnly();
    }

    toA1RefOnly(): string {
      const colPrefix = getA1Prefix(this.colMode);
      const rowPrefix = getA1Prefix(this.rowMode);

      return colPrefix + intToColChars(this.column) + rowPrefix + this.row.toString();
    }

    toR1C1RefOnly(): string {
      return 'R' + this.row + 'C' + this.column;
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
     * Returns true if the range object represents a contiguous range.
     */
    get isContiguous(): boolean {
      return this.regions.length === 1;
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

    toFormula(r1c1: boolean = false): string {
      return this.regions.map(([tl, br]) => tl.toFormula(r1c1) + ':' + br.toFormula(r1c1)).join(',');
    }
  }

  export interface IExpr {
    /**
     * Returns the type tag for the expression subtype,
     * for use in pattern-matching expressions. Also
     * available as a static property on ReferenceExpr
     * types.
     */
    readonly type: string;

    /**
     * Generates a valid Excel formula from this expression.
     * @param r1c1 If true, returns a formula with R1C1 references,
     * otherwise returns a formula with A1 references.
     */
    toFormula(r1c1: boolean): string;

    /**
     * Pretty-prints the AST as a string.  Note that this
     * does not produce a valid Excel formula.
     */
    toString(): string;
  }

  export class ReferenceRange implements IExpr {
    static readonly type: 'ReferenceRange' = 'ReferenceRange';
    readonly type = ReferenceRange.type;
    readonly range: Range;

    constructor(env: Env, range: Range) {
      this.range = range.copyWithNewEnv(env);
    }

    toFormula(r1c1: boolean = false): string {
      return this.range.toFormula(r1c1);
    }

    toString(): string {
      return 'ReferenceRange(' + this.range.toString() + ')';
    }
  }

  export class ReferenceAddress implements IExpr {
    static readonly type: 'ReferenceAddress' = 'ReferenceAddress';
    readonly type = ReferenceAddress.type;
    readonly address: Address;

    constructor(env: Env, address: Address) {
      this.address = address.copyWithNewEnv(env);
    }

    toString(): string {
      return 'ReferenceAddress(' + this.address.toString() + ')';
    }

    toFormula(r1c1: boolean = false): string {
      return this.address.toFormula(r1c1);
    }
  }

  export class ReferenceNamed implements IExpr {
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

  export class FunctionApplication implements IExpr {
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

    toFormula(r1c1: boolean = false): string {
      return this.name + '(' + this.args.map(arg => arg.toFormula(r1c1)).join(',') + ')';
    }
  }

  export class NumberLiteral implements IExpr {
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

  export class StringLiteral implements IExpr {
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

  export class BooleanLiteral implements IExpr {
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
  export class PoisonPill implements IExpr {
    static readonly type: 'PoisonPill' = 'PoisonPill';
    readonly type = PoisonPill.type;

    toString(): string {
      throw new Error('This object should never appear in an AST.');
    }

    toFormula(): string {
      throw new Error('This object should never appear in an AST.');
    }
  }

  export class ParensExpr implements IExpr {
    static readonly type: 'ParensExpr' = 'ParensExpr';
    readonly type = ParensExpr.type;
    readonly expr: Expression;

    constructor(expr: Expression) {
      this.expr = expr;
    }

    toString(): string {
      return 'Parens(' + this.expr.toString() + ')';
    }

    toFormula(r1c1: boolean = false): string {
      return '(' + this.expr.toFormula(r1c1) + ')';
    }
  }

  export class BinOpExpr implements IExpr {
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

    toFormula(r1c1: boolean = false): string {
      return this.exprL.toFormula(r1c1) + ' ' + this.op + ' ' + this.exprR.toFormula(r1c1);
    }

    toString(): string {
      return (
        'BinOpExpr(' + this.op.toString() + ',' + this.exprL.toFormula(false) + ',' + this.exprR.toFormula(false) + ')'
      );
    }
  }

  export class UnaryOpExpr implements IExpr {
    static readonly type: 'UnaryOpExpr' = 'UnaryOpExpr';
    readonly type = UnaryOpExpr.type;
    readonly op: string;
    readonly expr: Expression;

    constructor(op: string, expr: Expression) {
      this.op = op;
      this.expr = expr;
    }

    toFormula(r1c1: boolean = false): string {
      return this.op + this.expr.toFormula(r1c1);
    }

    toString(): string {
      return 'UnaryOpExpr(' + this.op.toString() + ',' + this.expr.toFormula(false) + ')';
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
