# pips


**pips** is a GUI for pip and conda - Python package manager, written in PowerShell.

This script helps to keep packages updated.


1. Hit Check for Updates
2. Choose packages
3. Select *Update* action and hit Execute


## features


- Dependency-free
- Search and install from **pip**, **conda** and **github**
- Looks up for installed Python distributions
- Manage environments
- Filter and sort packages
- Documentation viewer with simple highlighting and browser-like navigation
- Completion for packages, versions, paths, git tags, PyDoc topics
- Package name typosquatting check


![](screenshot.png)


## todo


- [X] Save user env list
- [X] Add search with Github API
- [X] Add editable packages to Install dialog (git, local); pip list --editable; Appropriate checkbox is still needed in the install dialog
- [X] Mistype verification with Damerau–Levenshtein distance (fuzzy). As PowerShell v3 is JIT-compiled, no need for external DLLs
- [ ] Sort fuzzy candidates by PyPI download count
- [ ] GPG signature verification for packages (with gpg.exe)
- [ ] Add some integration with VirusTotal (sha-256 of archive + link to VT for a starter)
- [ ] Fix "isolated" checkbox behavior is somewhat uncertain
- [ ] Move known package index builder code into the main script
- [ ] Sort by Popularity, Downloads, ...
- [ ] Add package name to JobName to prevent repeating requests already being queried
- [ ] search through PyDoc browser
- [ ] Verbosity control (-v, -vv, -vvv) for pip, a combobox over the log pane
- [X] Filter modes: Whole, And, Or, Regexp (RegExp seems feasible with Linq, but a little quirky)
- [X] Fix peps url 0000 numbering
- [X] Install packages dialog like in R Studio (with package name completion)
- [X] PyDoc and PyPi <s>meta-info mirror local indexing and searching</s> already in PyDoc apropos & topics <s>(binary indeces are too big, separated repo needed or even *git lfs*)</s>
- [X] Now assuming *utf-8* for IO whilst old(?) Python versions use cp1252, which generates empty strings for Chinese, Russian, etc.
- [ ] Filter package search results for conda by relevant Architecture
- [ ] Implement some missing conda commands
- [ ] *conda* command has a nice options to work with but gives wierd output - suggests older package versions than installed. Needed intervention with channels?
- [ ] conda activate & conda env creation
- [ ] Provide conda patcher (conda from PyPI is deliberately broken by devs by means of changing entrypoints to "warning stub"). Conda is under *BSD 3-Clause License* -> sounds legit.


## trademarks


The Python logo used in this program is a trademark of [The Python Software Foundation](https://www.python.org/psf/trademarks/).
This program is written in the PowerShell programming language and has a few pieces of inline Python code, and relies on external Python executables.
