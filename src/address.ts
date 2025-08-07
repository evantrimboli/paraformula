import { AST } from './ast';
import { Primitives as P, CharUtil as CU } from 'parsecco';
import { Primitives as PP } from './primitives';
import { colCharsToInt } from './util';

export namespace Address {
  const setupAddressRowOrColParser = (char: string) => {
    const charMatcher = P.char(char);

    const emptyPair = new AST.R1C1Pair();

    const relative = P.pipe<number, AST.R1C1Pair>(
      P.between<CU.CharStream, CU.CharStream, number>(P.str(char + '['))(P.char(']'))(PP.signedInt)
    )(n => new AST.R1C1Pair(n, 'relative'));

    const absolute = P.pipe2<CU.CharStream, number, AST.R1C1Pair>(charMatcher)(P.integer)(
      (_, n) => new AST.R1C1Pair(n, 'absolute')
    );

    const charOnly = P.pipe<CU.CharStream, AST.R1C1Pair>(charMatcher)(() => emptyPair);

    return P.choices(relative, absolute, charOnly);
  };

  export const r1c1Row = setupAddressRowOrColParser('R');
  export const r1c1Col = setupAddressRowOrColParser('C');

  export const r1c1Address = P.pipe2<AST.R1C1Pair, AST.R1C1Pair, AST.Address>(r1c1Row)(r1c1Col)((row, col) =>
    AST.Address.r1c1(row, col, PP.EnvStub)
  );

  const a1Mode = P.choice<AST.AddressMode>(P.pipe<CU.CharStream, AST.AddressMode>(P.char('$'))(() => 'absolute'))(
    P.pipe<undefined, AST.AddressMode>(P.ok(undefined))(() => 'relative')
  );

  const a1ColPart = P.pipe<CU.CharStream[], CU.CharStream>(P.many1(P.upper))(CU.CharStream.concat);

  /**
   * Parses the column component of an A1 address, including address mode.
   */
  const a1Col = P.pipe2<AST.AddressMode, CU.CharStream, AST.A1Pair>(a1Mode)(a1ColPart)(
    (mode, col) => new AST.A1Pair(colCharsToInt(col.toString()), mode)
  );

  /**
   * Parses the row component of an A1 address, including address mode.
   */
  const a1Row = P.pipe2<AST.AddressMode, number, AST.A1Pair>(a1Mode)(P.integer)(
    (mode, row) => new AST.A1Pair(row, mode)
  );

  /**
   * Parses an A1 address, with address modes.
   */
  export const a1Address = P.pipe2<AST.A1Pair, AST.A1Pair, AST.Address>(a1Col)(a1Row)((col, row) =>
    AST.Address.a1(row, col, PP.EnvStub)
  );

  /**
   * Parses either an A1 or R1C1 address.
   */
  export const addrAny = P.choice(r1c1Address)(a1Address);
}
