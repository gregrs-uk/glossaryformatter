# Print the glossary from the xlsx file with Markdown formatting

import glossaryformatter as gf
import sys

# check that we have a single argument i.e. the filename
if len(sys.argv) == 1:
    raise RuntimeError('Please supply the name of the glossary xlsx file')
elif len(sys.argv) > 2:
    raise RuntimeError('Please supply the name of the glossary xlsx file as ' +
                       'single argument')

gf.set_output_encoding()

print('# Glossary\n')

the_terms = gf.get_terms_excel(filename=sys.argv[1], last_col=6)
gf.print_glossary(the_terms,
                  subcat_col = 'Sub-category',
                  cat_prefix='## ',
                  cat_suffix='\n',
                  subcat_prefix='### ',
                  subcat_suffix='\n',
                  term_prefix='* **',
                  term_suffix='**')
