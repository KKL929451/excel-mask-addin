/* V3: reversible display masking. NEVER cryptographic encryption.
 * Cross-platform companion for vba/Module1.bas (same character order).
 * ASCII =, +, -, ~ are escaped as symbol pairs, not produced as raw ciphertext.
 * 0-9 A-Z a-z share one alphabet; other scripts remain in their own categories.
 */
(function (root, factory) {
  const codec = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = codec;
  else root.MaskCodec = codec;
})(typeof globalThis === "object" ? globalThis : this, function () {
  "use strict";
  const ALPHANUM = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
  const RESERVED = { "=": "~!", "+": "~#", "-": "~$", "~": "~%" };
  const UNRESERVED = { "~!": "=", "~#": "+", "~$": "-", "~%": "~" };
  const PUNCT = Array.from({length: 94}, (_, i) => String.fromCharCode(33 + i))
    .filter(c => !ALPHANUM.includes(c) && !"=+-~".includes(c)).join("");
  const categories = [
    [13312, 19903], [19968, 40959],
    [12353, 12438], [12449, 12538], [44032, 55203]
  ];
  const MODULUS = 1114112;
  const BASE_SHIFT = 500;
  const rotatedIndex = (index, width, shift, direction) => {
    const offset = 1 + (shift % (width - 1));
    return ((index + direction * offset) % width + width) % width;
  };
  function transformText(input, excelRow, excelCol, direction) {
    if (direction !== 1 && direction !== -1) throw new Error("Invalid direction");
    const chars = Array.from(String(input));
    const result = [];
    let charPos = 1;
    for (let i = 0; i < chars.length; i++, charPos++) {
      const c = chars[i];
      if (direction === -1 && c === "~") {
        if (i + 1 >= chars.length) throw new Error("Incomplete V3 escape");
        const token = "~" + chars[++i];
        if (!Object.prototype.hasOwnProperty.call(UNRESERVED, token)) {
          throw new Error("Invalid V3 escaped punctuation: " + token);
        }
        result.push(UNRESERVED[token]);
        continue;
      }
      const shift = (BASE_SHIFT + 7 * excelRow + 11 * excelCol + charPos) % MODULUS;
      if (direction === 1 && Object.prototype.hasOwnProperty.call(RESERVED, c)) {
        result.push(RESERVED[c]);
        continue;
      }
      const alphaIdx = ALPHANUM.indexOf(c);
      if (alphaIdx !== -1) {
        result.push(ALPHANUM[rotatedIndex(alphaIdx, ALPHANUM.length, shift, direction)]);
        continue;
      }
      const punctIdx = PUNCT.indexOf(c);
      if (punctIdx !== -1) {
        result.push(PUNCT[rotatedIndex(punctIdx, PUNCT.length, shift, direction)]);
        continue;
      }
      const cp = c.codePointAt(0);
      const range = categories.find(([low, high]) => cp >= low && cp <= high);
      if (!range) { result.push(c); continue; }
      const [low, high] = range;
      result.push(String.fromCodePoint(low + rotatedIndex(cp - low, high - low + 1, shift, direction)));
    }
    return result.join("");
  }
  return { transformText, categories, ALPHANUM, PUNCT, version: 3 };
});
