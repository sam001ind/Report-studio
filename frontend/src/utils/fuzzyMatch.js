/**
 * Text normalization and fuzzy matching algorithms for Data Comparison Studio
 */

/**
 * Normalizes a string based on user options
 */
export function normalizeText(str, options = {}) {
  if (str === null || str === undefined) return '';
  let s = String(str);

  if (options.ignoreCase !== false) {
    s = s.toLowerCase();
  }

  if (options.stripSpaces) {
    s = s.replace(/\s+/g, ' ').trim();
  } else {
    s = s.trim();
  }

  if (options.stripPunctuation) {
    s = s.replace(/[.,/#!$%^&*;:{}=\-_`~()?"'[\]\\|<>+@]/g, '');
  }

  if (options.stripLeadingZeros) {
    s = s.replace(/^0+(?=\d)/, '');
  }

  return s;
}

/**
 * Levenshtein distance between two strings
 */
export function levenshteinDistance(a, b) {
  if (a === b) return 0;
  if (!a.length) return b.length;
  if (!b.length) return a.length;

  const v0 = new Array(b.length + 1);
  const v1 = new Array(b.length + 1);

  for (let i = 0; i <= b.length; i++) v0[i] = i;

  for (let i = 0; i < a.length; i++) {
    v1[0] = i + 1;
    for (let j = 0; j < b.length; j++) {
      const cost = a[i] === b[j] ? 0 : 1;
      v1[j + 1] = Math.min(v1[j] + 1, v0[j + 1] + 1, v0[j] + cost);
    }
    for (let j = 0; j <= b.length; j++) v0[j] = v1[j];
  }

  return v1[b.length];
}

/**
 * Calculates similarity ratio between 0.0 and 1.0 using Levenshtein distance
 */
export function stringSimilarity(str1, str2) {
  if (str1 === str2) return 1.0;
  if (!str1 || !str2) return 0.0;

  const maxLen = Math.max(str1.length, str2.length);
  if (maxLen === 0) return 1.0;

  const dist = levenshteinDistance(str1, str2);
  return Math.max(0, (maxLen - dist) / maxLen);
}

/**
 * Token Sort Similarity: handles word order variations (e.g. "John Doe" vs "Doe John")
 */
export function tokenSortSimilarity(str1, str2) {
  if (str1 === str2) return 1.0;
  if (!str1 || !str2) return 0.0;

  const tokens1 = str1.split(/\s+/).filter(Boolean).sort().join(' ');
  const tokens2 = str2.split(/\s+/).filter(Boolean).sort().join(' ');

  const ratioDirect = stringSimilarity(str1, str2);
  const ratioSorted = stringSimilarity(tokens1, tokens2);

  return Math.max(ratioDirect, ratioSorted);
}

/**
 * Composite key similarity calculator
 */
export function calculateCompositeSimilarity(rowA, rowB, keyMappings, options = {}) {
  if (!keyMappings || keyMappings.length === 0) return 0;

  let totalScore = 0;
  let weights = 0;

  for (const map of keyMappings) {
    const valA = normalizeText(rowA[map.leftCol], options);
    const valB = normalizeText(rowB[map.rightCol], options);

    if (valA === valB && valA !== '') {
      totalScore += 1.0;
    } else {
      const sim = tokenSortSimilarity(valA, valB);
      totalScore += sim;
    }
    weights += 1;
  }

  return weights > 0 ? (totalScore / weights) : 0;
}
