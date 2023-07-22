[![Hippocratic License HL3-LAW-MIL-SV](https://img.shields.io/static/v1?label=Hippocratic%20License&message=HL3-LAW-MIL-SV&labelColor=5e2751&color=bc8c3d)](https://firstdonoharm.dev/version/3/0/law-mil-sv.html)

# Spell Check Function for Google Sheets

This is a custom function that allows you to passively check the spelling of individual cells in Google Sheets, using [Google Apps Script](https://developers.google.com/apps-script). To use this function in your spreadsheet, paste the script into that spreadsheet's Google Apps Script settings, then access it with the function name `SPCHECK`.

View the script's source code in the [Code.gs](Code.gs) file in this repo.  

> Warning: This is a very hacky solution, which relies on making HTTP requests to a dictionary website. Which, idk, might not be great netiquette. It is not intended for long-term use or for applying to very large spreadsheets.  

## Purpose

The spell check function built into Google Sheets can only be used interactively, stepping through the document one error at a time, and is intended for when you want to *correct* spelling errors. However, sometimes you just need to know whether a typo *exists* in a cell without wanting to change it, and sometimes you might need that information to be visible while you're typing in other cells. 

For example, this script may be useful for making a list of misspellings to add to an autocorrect list, or if you have a spreadsheet that collects text submitted by others, and you want to identify submissions with typos so that you can contact the writer for clarification. 


## Usage

The following instructions are based on the desktop browser version of Google Sheets.

### Installation

The script must be added to each spreadsheet where you want to use it. 

1. Open your spreadsheet in Google Sheets.
2. In the menu, select `Extensions > Apps Script`.
3. You should see a text editor with a file named `Code.gs`. Delete any existing text in `Code.gs`.
4. Select the entire text of [Code.gs](Code.gs), paste it into the `Code.gs` on your Google Sheet, and save. You should not need to select Deploy.

You should now have a new function named `SPCHECK` available in your spreadsheet. 

For more details about adding Apps Script or custom functions to Google Sheets, see the [Apps Script documentation](https://developers.google.com/apps-script/quickstart/custom-functions).

### Using the function in a spreadsheet

Usage is mostly the same as any other Sheets function. 

1. Select a blank cell in the spreadsheet and type`=SPCHECK`. It may appear in the list of suggestions once you start typing.
    - Currently, the function does *not* seem to appear in the list of functions under the`Insert` menu, so you'll need to type it out.
    - Be sure to include the `=` at the beginning, as you would for any other Sheets function.
2. Select a cell, or range of cells, as input to be spell-checked.
3. The output of the function will appear in the cell where you entered the function, or in the adjacent cells if you selected a range as input.

#### Output

- Each misspelling from the input will be printed in the output, separated by a space.
- A blank output cell indicates that the function did not find any misspelled words in the input cell.
    - If you prefer to see some text indicating a successful check with no errors found, see [Modifying the behavior](#modifying-the-behavior) below. 

### Troubleshooting

#### `#NAME?`
If the cell where you entered the function displays the text `#NAME?`, this means Google Sheets cannot find the definition of the function. Check that you pasted the script correctly into `Code.gs`. If the script was entered correctly, you might also try refreshing the spreadsheet after saving the script.

#### Too many requests
This function does not have its own internal dictionary, and instead checks spelling by looking up your text in an external online dictionary. If you use the spell check function on too much text too quickly, it is possible for the dictionary service to stop responding to these requests. I am not sure what the limit is. Note that one request is sent for each *word* of input, not for each cell.

#### Internet connection
This function will not work offline. An internet connection is required in order to access the dictionary.

## How it works

The script has two parts: the `SPCHECK` function, which gets called from Google Sheets, and the helper function `hasTypo_`, which the user cannot access directly. `hasTypo_` contains the main logic of the script and takes a single cell as input, while the body of `SPCHECK` handles ranges as input, and calls `hasTypo_` on each individual cell.

### Inputs

The input may either be a string that represents the content of a single cell of a spreadsheet, or an array of strings representing the contents from a range of adjacent cells.

### Outputs
Output is rendered as text in the Sheets cell where the user calls the `SPCHECK` function. There are two possible outputs:
- If the spell check finds no misspelled words, then the output is an empty string.
- Otherwise, the output is a string that contains a sequence of all the misspelled words from the input, separated by spaces.  
Note that "misspelled word" means any word which was not found in the dictionary reference.

### Steps
- Check whether the input is an array or a single cell
- If the input is an array, use a mapping function to call `hasTypo_` on each cell ad return the result as a new array of cells
- If input is not an array (which means it must be a single cell), directly call `hasTypo_` on the input

#### Tokenization
- Read the cell as a string
- Split the string into an array of words, based on the regex `/[^a-zA-Z']+/`

#### Dictionary HTTP request
- For each word, try accessing `https://www.dictionary.com/browse/` plus the word
- If accessing that URL generates response code 404, then this means the word is not in the dictionary, so add the word to a list of typos
- Return that list and print it to the output cell

## Modifying the behavior

### Changing the dictionary
You can easily alter the function to look up your words in a different online dictionary. However, not all dictionaries are compatible--it has to be one that returns a 404 HTTP response code when you try to visit the page for a non-defined word.

1. Get the URL of an example word entry from the dictionary website you would like to use. For example, `https://www.dictionary.com/browse/cat`.
2. Replace the word with `%s`, as in `https://www.dictionary.com/browse/%s`
2. On line 43 of `Code.gs`, remove just the URL `https://www.dictionary.com/browse/%s` and paste your modified URL in its place. Make sure it's still contained in quote marks and that the rest of the line is unchanged.

### Reporting correct cells
Between line 51 and line 52 (that is, right before `return typos`), add a function that checks the length of `typos`. If the length is greater than 0, it should return `typos` as is. If the length is 0 (array is empty), then it should either add an 'all clear' phrase to the list and return it, or return that phrase instead of returning the list. 

### Word delineation
The script uses a regular expression to define how the content of a cell is divided into words. Currently that expression is `/[^a-zA-Z']+/`. You may want to edit this regex to change what the script accepts as a word.

## Further development

### Known issues
- The script will stop working if you use it on too big of a spreadsheet at once (>100 lines), because Dictionary.com will start rejecting the HTTP requests if too many are submitted too fast.
- The dictionary currently used, Dictionary.com, contains some entries such as abbreviations, atomic symbols, etc that may not normally be considered words. 
- The regex splits words on numerals, so words such as "1st" will come out as "st"

### Roadmap

Possible future improvements include:
- Adding a built-in dictionary so that it doesn't rely on HTTP requests

### Contributing

Any improvements are welcome! Feel free to submit an issue with any suggestions or bugs you encounter, or to fork and submit a pull request.

I'd especially love help with setting up that built-in dictionary!

## License
The Hippocratic License is an *almost* open license that says you can do basically anything you want with this code as long as it doesn't hurt people. Check out [LICENSE.md](LICENSE.md) as well as the [Hippocratic License website](https://firstdonoharm.dev/).

## Author
Hey, I'm **caesiumtea**, AKA Vance! Feel free to contact me with any feedback.
- [Website and social links](https://caesiumtea.glitch.me/)
- [@caesiumtea_dev on Twitter](https://www.twitter.com/caesiumtea_dev)
- [@entropy@mastodon.social](https://mastodon.social/@entropy)