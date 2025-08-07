import { AST } from './ast';
import type { CharUtil as CU } from 'parsecco';
import { Primitives as P } from 'parsecco';
import { Primitives as PP } from './primitives';
import { Address as PA } from './address';

export namespace Range {
  /**
   * Parses an A1 range suffix.
   */
  export const rangeA1Suffix = P.right<CU.CharStream, AST.Address>(P.str(':'))(PA.a1Address);

  /**
   * Parses an R1C1 range suffix.
   */
  export const rangeR1C1Suffix = P.right<CU.CharStream, AST.Address>(P.str(':'))(PA.r1c1Address);

  /**
   * Parses an A1-style contiguous range.
   */
  export const rangeA1Contig = P.pipe2<AST.Address, AST.Address, AST.Range>(PA.a1Address)(rangeA1Suffix)(
    (a1, a2) => new AST.Range([[a1, a2]])
  );

  /**
   * Parses an R1C1-style contiguous range.
   */
  export const rangeR1C1Contig = P.pipe2<AST.Address, AST.Address, AST.Range>(PA.r1c1Address)(rangeR1C1Suffix)(
    (a1, a2) => new AST.Range([[a1, a2]])
  );

  /**
   * Parses a discontiguous A1-style range list.
   */
  export const rangeA1Discontig = P.pipe2<AST.Range[], AST.Range, AST.Range>(
    // recursive case
    P.many1(P.left<AST.Range, CU.CharStream>(rangeA1Contig)(PP.comma))
  )(
    // base case
    rangeA1Contig
  )(
    // reducer
    (rs, r) => rs.reduce((acc, r) => acc.merge(r)).merge(r)
  );

  /**
   * Parses a discontiguous R1C1-style range list.
   */
  export const rangeR1C1Discontig = P.pipe2<AST.Range[], AST.Range, AST.Range>(
    // recursive case
    P.many1(P.left<AST.Range, CU.CharStream>(rangeR1C1Contig)(PP.comma))
  )(
    // base case
    rangeR1C1Contig
  )(
    // reducer
    (rs, r) => rs.reduce((acc, r) => acc.merge(r)).merge(r)
  );

  /**
   * Parses a contiguous range.
   */
  export const rangeContig = P.choice(rangeR1C1Contig)(rangeA1Contig);

  /**
   * Parses a discontiguous range.
   */
  export const rangeDiscontig = P.choice(rangeR1C1Discontig)(rangeA1Discontig);

  /**
   * Parses any range, A1-style or R1C1-style, contiguous or discontiguous.
   */
  export const rangeAny = P.choice(rangeDiscontig)(rangeContig);
}
