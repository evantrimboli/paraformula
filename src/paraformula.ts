import type { AST } from './ast';
import { Primitives as P, CharUtil as CU } from 'parsecco';
import { Expression as PE } from './expression';
import { Range as PR } from './range';

export namespace Paraformula {
  /**
   * Top-level grammar definition.
   */
  export const grammar: P.IParser<AST.Expression> = P.right<CU.CharStream, AST.Expression>(P.char('='))(
    PE.expr(PR.rangeAny)
  );

  const processOutput = (output: P.Outcome<AST.Expression>): AST.Expression => {
    switch (output.tag) {
      case 'success':
        return output.result;
      case 'failure':
        throw new Error('Unable to parse input: ' + output.error_msg);
    }
  };

  /**
   * Parses an Excel formula and returns an AST.  Throws an
   * exception if the input is invalid.
   * @param input A formula string
   */
  export function parse(input: string): AST.Expression {
    const cs = new CU.CharStream(input);
    const it = grammar(cs);
    const elem = it.next();
    if (elem.done) {
      return processOutput(elem.value);
    }
    throw new Error('This should never happen.');
  }

  /**
   * Parses an Excel formula and returns an AST.  Throws an
   * exception if the input is invalid. Yieldable.
   * @param input A formula string
   */
  export function* yieldableParse(input: string): Generator<undefined, AST.Expression, undefined> {
    const cs = new CU.CharStream(input);
    const output = yield* grammar(cs);
    return processOutput(output);
  }
}
