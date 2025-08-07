import { AST } from '../src/ast';
import { Primitives as PP } from '../src/primitives';

export const a1Addr = (
  row: number,
  col: number,
  rowMode: AST.AddressMode,
  colMode: AST.AddressMode,
  env = PP.EnvStub
): AST.Address => AST.Address.a1(new AST.A1Pair(row, rowMode), new AST.A1Pair(col, colMode), env);

export const relA1Addr = (row: number, col: number, env = PP.EnvStub): AST.Address =>
  AST.Address.a1(new AST.A1Pair(row, 'relative'), new AST.A1Pair(col, 'relative'), env);

export const absA1Addr = (row: number, col: number, env = PP.EnvStub): AST.Address =>
  AST.Address.a1(new AST.A1Pair(row, 'absolute'), new AST.A1Pair(col, 'absolute'), env);

export const r1c1Addr = (
  row: number,
  col: number,
  rowMode: AST.AddressMode,
  colMode: AST.AddressMode,
  env = PP.EnvStub
): AST.Address => AST.Address.r1c1(new AST.R1C1Pair(row, rowMode), new AST.R1C1Pair(col, colMode), env);

export const relR1C1Addr = (row: number, col: number, env = PP.EnvStub): AST.Address =>
  AST.Address.r1c1(new AST.R1C1Pair(row, 'relative'), new AST.R1C1Pair(col, 'relative'), env);

export const absR1C1Addr = (row: number, col: number, env = PP.EnvStub): AST.Address =>
  AST.Address.r1c1(new AST.R1C1Pair(row, 'absolute'), new AST.R1C1Pair(col, 'absolute'), env);
