# Makefile notes

* Great resource: https://makefiletutorial.com/#include-makefiles

* The makefile must be named "makefile" or "Makefile" (case sensitive) for the `make` terminal command on Ubuntu linux to work
* The makefile must also be in the current working directory, use `cd` in the terminal to navigate there

* `$()` references a variable

* The flag `make -f` allows users to specify file by name

* `$^`, `$@`, `%`: these are wildcards, see web page above for disambiguation

* patsubst - potentially stands for "pattern substring"?
    * Syntax: `$(patsubst TO_REPLACE, REPLACEMENT, SEARCH_STRING)`
    * Breaks `SEARCH_STRING` into substrings on whitespace (blanks, tabs, etc.)
    * Searches for `TO_REPLACE` in every substring of `SEARCH_STRING`
    * Replaces `TO_REPLACE` with `REPLACEMENT`
    * Resource: https://www.gnu.org/software/make/manual/html_node/Text-Functions.html

* `.PHONY`: https://www.gnu.org/software/make/manual/html_node/Phony-Targets.html