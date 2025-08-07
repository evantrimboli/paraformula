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

    describe('misc', () => {
      const input = '=XY100';

      it('should parse', () => {
        const output = P.parse(input);
        assertReferenceAddress(output);

        expect(output.address.column).to.eql(649);
        expect(output.address.colMode).to.eql('relative');
        expect(output.address.row).to.eql(100);
        expect(output.address.rowMode).to.eql('relative');
      });

      it('should output', () => {
        const output = P.parse(input);
        assertReferenceAddress(output);

        expect(output.toFormula()).to.eql('XY100');
      });
    });
  });

  describe('R1C1 style addresses', () => {
    interface Part {
      readonly input: string;
      readonly label: string;
      readonly expectedValue: readonly [value: number, type: AST.AddressMode];
      readonly expectedRender: string;
    }

    const rowParts: readonly Part[] = [
      {
        input: 'R',
        label: 'current row, no value',
        expectedValue: [0, 'relative'],
        expectedRender: 'R'
      },
      {
        input: 'R11',
        label: 'absolute row',
        expectedValue: [11, 'absolute'],
        expectedRender: 'R11'
      },
      {
        input: 'R[12]',
        label: 'relative row, no sign',
        expectedValue: [12, 'relative'],
        expectedRender: 'R[12]'
      },
      {
        input: 'R[+13]',
        label: 'relative row, positive sign',
        expectedValue: [13, 'relative'],
        expectedRender: 'R[13]'
      },
      {
        input: 'R[-14]',
        label: 'relative row, negative sign',
        expectedValue: [-14, 'relative'],
        expectedRender: 'R[-14]'
      }
    ];

    const colParts: readonly Part[] = [
      {
        input: 'C',
        label: 'current col, no value',
        expectedValue: [0, 'relative'],
        expectedRender: 'C'
      },
      {
        input: 'C21',
        label: 'absolute col',
        expectedValue: [21, 'absolute'],
        expectedRender: 'C21'
      },
      {
        input: 'C[22]',
        label: 'relative col, no sign',
        expectedValue: [22, 'relative'],
        expectedRender: 'C[22]'
      },
      {
        input: 'C[+23]',
        label: 'relative col, positive sign',
        expectedValue: [23, 'relative'],
        expectedRender: 'C[23]'
      },
      {
        input: 'C[-24]',
        label: 'relative col, negative sign',
        expectedValue: [-24, 'relative'],
        expectedRender: 'C[-24]'
      }
    ];

    const runIt = (rowPart: Part, colPart: Part) => {
      const input = `=${rowPart.input}${colPart.input}`;

      describe(`${rowPart.label} - ${colPart.label} - ${input}`, () => {
        it('should parse', () => {
          const output = P.parse(input);
          assertReferenceAddress(output);

          const [row, rowMode] = rowPart.expectedValue;
          const [col, colMode] = colPart.expectedValue;

          expect(output.address.column).to.eql(col);
          expect(output.address.colMode).to.eql(colMode);
          expect(output.address.row).to.eql(row);
          expect(output.address.rowMode).to.eql(rowMode);
        });

        it('should output', () => {
          const output = P.parse(input);
          assertReferenceAddress(output);

          expect(output.toFormula()).to.eql(rowPart.expectedRender + colPart.expectedRender);
        });
      });
    };

    for (const r of rowParts) {
      for (const c of colParts) {
        runIt(r, c);
      }
    }
  });
});
