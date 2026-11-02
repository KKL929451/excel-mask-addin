/* V2: reversible readable masking, not cryptographic encryption.
   ASCII digits and both letter cases now share ONE 62-character alphabet.
   Chinese / Japanese / Korean / punctuation use the unchanged V1 rotation. */
(function (root, factory) {
  const codec = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = codec;
  else root.MaskCodec = codec;
})(typeof globalThis === "object" ? globalThis : this, function () {
  "use strict";

  // The character order MUST match the companion VBA code exactly.
  const ALPHANUM = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
  const categories = [
    [33, 47], [58, 64], [91, 96], [123, 126],
    [13312, 19903], [19968, 40959],
    [12353, 12438], [12449, 12538], [44032, 55203]
  ];
  const MODULUS = 1114112;
  const BASE_SHIFT = 500;

  function transformText(input, excelRow, excelCol, direction) {
    if (direction !== 1 && direction !== -1) throw new Error("Invalid direction");
    return Array.from(String(input)).map((char, index) => {
      const code = char.codePointAt(0);
      const shift = (BASE_SHIFT + 7 * excelRow + 11 * excelCol + index + 1) % MODULUS;
      const alphaIndex = ALPHANUM.indexOf(char);
      if (alphaIndex >= 0) {
        const offset = 1 + (shift % (ALPHANUM.length - 1));
        const nextIndex = ((alphaIndex + direction * offset) % ALPHANUM.length + ALPHANUM.length) % ALPHANUM.length;
        return ALPHANUM[nextIndex];
      }
      const range = categories.find(([low, high]) => code >= low && code <= high);
      if (!range) return char;
      const [low, high] = range;
      const width = high - low + 1;
      const offset = 1 + (shift % (width - 1));
      const nextIndex = ((code - low + direction * offset) % width + width) % width;
      return String.fromCodePoint(low + nextIndex);
    }).join("");
  }

  return { transformText, categories, ALPHANUM, version: 2 };
});
