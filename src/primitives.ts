import { AST } from './ast';
import type { CharUtil as CU } from 'parsecco';
import { Primitives as P } from 'parsecco';

export namespace Primitives {
  /**
   * TODO remove: this is a stub until parsecco supports parsing with user state.
   */
  export const EnvStub = new AST.Env('', '', '');

  /**
   * Parse an Excel integer.
   */
  export const signedInt = P.choices(
    // leading + sign
    P.pipe2<CU.CharStream, number, number>(P.char('+'))(P.integer)((_sign, num) => num),
    // leading - sign
    P.pipe2<CU.CharStream, number, number>(P.char('-'))(P.integer)((_sign, num) => -num),
    // no leading sign
    P.integer
  );

  /**
   * Parses a `p`, preceeded and suceeded with whitespace. Returns
   * only the result of `p`.
   * @param p A parser
   */
  export function wsPad<T>(p: P.IParser<T>): P.IParser<T> {
    return P.between<CU.CharStream, CU.CharStream, T>(P.ws)(P.ws)(p);
  }

  /**
   * Parses a comma surrounded by optional whitespace.
   */
  export const comma = wsPad(P.char(','));
}
