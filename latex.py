# Print a LaTeX source file of the glossary from the xlsx file

import glossaryformatter as gf
import sys

# check that we have a single argument i.e. the filename
if len(sys.argv) == 1:
    raise RuntimeError('Please supply the name of the glossary xlsx file')
elif len(sys.argv) > 2:
    raise RuntimeError('Please supply the name of the glossary xlsx file as ' +
                       'single argument')

gf.set_output_encoding()

# list of headers for reference columns
ref_cols = ['Piece 1', 'Piece 2']

# beginning of LaTeX file

print(r"""\documentclass[10pt,a4paper]{article}

\usepackage[margin=2cm]{geometry}
\usepackage{charter}
\usepackage{microtype}
\usepackage{multicol}
\usepackage{titlesec}
\usepackage{enumitem}

\setlength\parindent{0pt}
\setlength\parskip{6pt plus 2pt minus 1pt}
\setlength\columnsep{1cm}
\setlist[itemize]{leftmargin=*}
\providecommand\tightlist{
    \setlength{\itemsep}{6pt plus 3pt minus 3pt}
    \setlength{\parskip}{0pt}}

\titlespacing\section{0pt}{12pt}{0pt}
\titlespacing\subsection{0pt}{6pt}{3pt}
\titlespacing\subsubsection{0pt}{6pt}{3pt}
\setcounter{secnumdepth}{0}

\newcommand\term[1]{\item \textbf{#1}}
\newcommand\definition[1]{ -- #1}
\newcommand\refs[1]{\textsuperscript{ (#1)}}
\newcommand\separator{\vspace{6pt}\hrule}

\begin{document}

\section{Glossary}
""")


# reference legend

n = 1
for this_piece in ref_cols:
    print(r'\textsuperscript{' + str(n) + '} ' + this_piece)
    if n != len(ref_cols):
        print(r'\\[3pt]')
    n += 1

print(r""" 
\vspace{9pt}
\separator

\begin{multicols}{2}
""")


# the main chunk of the glossary

the_terms = gf.get_terms_excel(filename=sys.argv[1], last_col=6)
gf.print_glossary(the_terms,
                  subcat_col = 'Sub-category',
                  cat_prefix = '\subsection{',
                  cat_suffix = '}\n',
                  subcat_prefix = '\subsubsection{',
                  subcat_suffix = '}\n',
                  begin_terms = '\\begin{itemize}\n\\tightlist\n',
                  end_terms = '\n\end{itemize}\n\n\\separator\n',
                  term_prefix = '\\term{',
                  term_suffix = '}',
                  def_prefix = '\definition{',
                  def_suffix = '}',
                  ref_cols = ref_cols,
                  refs_prefix = '\\refs{',
                  refs_separator = ',',
                  refs_suffix = '}')


# end of LaTeX file

print(r"""\end{multicols}

\end{document}""")
