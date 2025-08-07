import { Paraformula as P } from '../src/paraformula';
import { assert, expect } from 'chai';
import 'mocha';
import { AST } from '../src/ast';

// instruct mocha to look for generator tests
require('mocha-generators').install();

function assertReferenceAddress(value: AST.Expression): asserts value is AST.ReferenceAddress {
  if (value.type !== 'ReferenceAddress') {
    assert.fail('Expected reference address');
  }
}

describe('Addresses', () => {
  describe('A1 style addresses', () => {
    describe('row relative, column relative', () => {
      const input = '=A1';

      it('should parse', () => {
        const output = P.parse(input);
        assertReferenceAddress(output);

        expect(output.address.column).to.eql(1);
        expect(output.address.colMode).to.eql('relative');
        expect(output.address.row).to.eql(1);
        expect(output.address.rowMode).to.eql('relative');
      });

      it('should output', () => {
        const output = P.parse(input);
        assertReferenceAddress(output);

        expect(output.toFormula()).to.eql('A1');
      });
    });

    describe('row absolute, column absolute', () => {
      const input = '=$A$1';

      it('should parse', () => {
        const output = P.parse(input);
        assertReferenceAddress(output);

        expect(output.address.column).to.eql(1);
        expect(output.address.colMode).to.eql('absolute');
        expect(output.address.row).to.eql(1);
        expect(output.address.rowMode).to.eql('absolute');
      });

      it('should output', () => {
        const output = P.parse(input);
        assertReferenceAddress(output);

        expect(output.toFormula()).to.eql('$A$1');
      });
    });

    describe('row absolute, column relative', () => {
      const input = '=A$1';

      it('should parse', () => {
        const output = P.parse(input);
        assertReferenceAddress(output);

        expect(output.address.column).to.eql(1);
        expect(output.address.colMode).to.eql('relative');
        expect(output.address.row).to.eql(1);
        expect(output.address.rowMode).to.eql('absolute');
      });

      it('should output', () => {
        const output = P.parse(input);
        assertReferenceAddress(output);

        expect(output.toFormula()).to.eql('A$1');
      });
    });

    describe('row relative, column absolute', () => {
      const input = '=$A1';

      it('should parse', () => {
        const output = P.parse(input);
        assertReferenceAddress(output);

        expect(output.address.column).to.eql(1);
        expect(output.address.colMode).to.eql('absolute');
        expect(output.address.row).to.eql(1);
        expect(output.address.rowMode).to.eql('relative');
      });

      it('should output', () => {
        const output = P.parse(input);
        assertReferenceAddress(output);

        expect(output.toFormula()).to.eql('$A1');
      });
    });
  });

  describe('R1C1 style addresses', () => {});
});
