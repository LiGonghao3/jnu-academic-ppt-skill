# LaTeX MK build configuration for JNU Thesis
# Usage: latexmk -pdf jnuthesis

# Use XeLaTeX as the default engine
$pdf_mode = 5;
$xelatex = 'xelatex -synctex=1 -interaction=nonstopmode -file-line-error %O %S';
$lualatex = 'lualatex -synctex=1 -interaction=nonstopmode -file-line-error %O %S';

# Use biber for bibliography
$bibtex = 'biber %O %B';

# Clean up auxiliary files
$clean_ext = 'aux bbl bcf blg fdb_latexmk fls log out run.xml synctex.gz toc tdo thm lol lot lof';

# Maximum number of runs
$max_repeat = 5;

# Preview PDF after compilation
$pdf_previewer = 'open -a Skim %O %S';

# Print progress
$silent = 0;
