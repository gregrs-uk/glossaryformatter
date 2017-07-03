# Glossary Formatter

A collection of [Python](https://www.python.org) scripts designed to create a printable glossary or vocabulary list with category headers from an Excel spreadsheet containing a list of categorised terms and definitions.

## Features

* Category (and optional subcategory) headers are printed where necessary
* Terms can also be referenced e.g. to a particular piece of music or chapter of a textbook
* Scripts are included for minimal, Markdown or [LaTeX](http://www.latex-project.org) formatting and can be customised to suit your needs
* The Markdown output can be converted to an HTML web page or Word document using [Pandoc](http://pandoc.org) 
* Tested using Python 2.7 on Mac OS X but should work on Linux and Windows too

## Usage

1. Sort your spreadsheet by category, subcategory (if present) then term. (For each category, terms with a blank subcategory should come last and by default will be printed under the subcategory name 'Other' if that category has other subcategories.) Preferably use the column names found in `glossary_example.xlsx` (although these can be customised if you wish).
1. Choose one of the scripts `minimal.py`, `markdown.py` or `latex.py` depending on what formatting you'd like, or use one as a starting point to create your own script.
1. If your spreadsheet has a different number of columns from `glossary_example.xlsx`, set the `last_col` argument of `get_terms_excel` to the number of columns your spreadsheet has.
1. If your spreadsheet has different column headers from the example one or you wish to customise the formatting, supply the relevant arguments to `print_glossary`. (See the documentation for `print_glossary` in `glossaryformatter/__init__.py` for details.)
1. Run your chosen script, supplying the filename as an argument.

## Examples

The `glossary_example.xlsx` file is an example Excel file containing some categorised musical terms and their definitions.

Print the glossary from the example Excel file to stdout with minimal formatting:

```
> ./minimal.py glossary_example.xlsx

GLOSSARY
========

Articulation

staccato - Played in a detached fashion
legato - Played in a smooth fashion

Melody

Ornamentation

trill - A type of ornament where there is rapid alternation between the main note and the note above it
mordent - A type of ornament where the main note is played, followed by the note above (upper mordent) or below (lower mordent) the main note, then the main note again

Other

sequence - Repetition of a melody (or an harmonic progression) but at different pitch level(s) rather than at the same pitch
```

Print the glossary from the example Excel file to stdout with Markdown formatting:

```
> ./markdown.py glossary_example.xlsx

# Glossary

## Articulation

* **staccato** - Played in a detached fashion
* **legato** - Played in a smooth fashion

## Melody

### Ornamentation

* **trill** - A type of ornament where there is rapid alternation between the main note and the note above it
* **mordent** - A type of ornament where the main note is played, followed by the note above (upper mordent) or below (lower mordent) the main note, then the main note again

### Other

* **sequence** - Repetition of a melody (or an harmonic progression) but at different pitch level(s) rather than at the same pitch
```

Save the glossary from the example Excel file to a file with Markdown formatting:

	> ./markdown.py glossary_example.xlsx > glossary.md

Pipe the Markdown to [Pandoc](http://pandoc.org) to create an HTML file:

	> ./markdown.py glossary_example.xlsx | pandoc -o glossary.html

Pipe the Markdown to [Pandoc](http://pandoc.org) to create a Word document:

	> ./markdown.py glossary_example.xlsx | pandoc -o glossary.docx

Create a [two-column PDF](example.pdf) from the example Excel file (including references) using [LaTeX](http://www.latex-project.org):

	> ./latex.py glossary_example.xlsx > glossary.ltx
	> pdflatex glossary.ltx

## Contributing

Feel free to modify the code to suit your needs. If you make changes that might be useful to others, please submit a pull request.
