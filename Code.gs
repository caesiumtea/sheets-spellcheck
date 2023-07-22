/**
 * Google Sheets spell check function
 * version 1.1, 8/17/22
 * caesiumtea
 */

/**
 * Shows misspelled words in input cell or cells.
 *
 * @param {string|Array<Array<string>>} input The cell or range of cells
 *     to check spelling in.
 * @return String listing any non-dictionary words in the input.
 * @customfunction
 */
function SPCHECK(input) {
  if (Array.isArray(input)) {
    return input.map(row => row.map(cell => hasTypo_(cell) ))
  }
  else {
    return hasTypo_(input);
  }   
}

/**
 * Helper function (not accessible from Sheets):
 *   Splits the input cell into a list of words.
 *   Decides whether each word is spelled correctly by checking
 *   whether it has a page on dictionary.com.
 * Returns a string containing any unlisted words; 
 *   empty string if every word in the cell was in the dictionary.
 */
function hasTypo_(cell) {
  let typoFound = false;
  let url;
  let lookup;
  let typos = "";

  let words = cell.split(/[^a-zA-Z']+/); 
    // regex: split on anything besides letters and apostrophes

  // for word in words, if word is not in dictionary, add to typos
  for (let w of words) {
    url = Utilities.formatString("https://www.dictionary.com/browse/%s", w);
      //checks whether the word has a page on dictionary.com
    lookup = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    if (lookup.getResponseCode() == 404) {
      // this means the word doesn't exist in the dictionary
      typoFound = true;
      typos += w + " ";
    }
  }
  return typos;
}