export const intToColChars = (dividend: number): string => {
  let quot = Math.floor(dividend / 26);
  const rem = dividend % 26;
  if (rem === 0) {
    quot -= 1;
  }
  const ltr = rem === 0 ? 'Z' : String.fromCharCode(64 + rem);
  return quot === 0 ? ltr : intToColChars(quot) + ltr;
};

export const colCharsToInt = (col: string): number => {
  const cti = (idx: number): number => {
    // get ASCII code and then subtract 64 to get Excel column #
    const code = col.charCodeAt(idx) - 64;
    // the value depends on the position; a column is a base-26 number
    const num = Math.pow(26, col.length - idx - 1) * code;
    if (idx === 0) {
      // base case
      return num;
    } else {
      // add this letter to number and recurse
      return num + cti(idx - 1);
    }
  };
  return cti(col.length - 1);
};
