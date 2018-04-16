# pips


**pips** is a GUI for pip, github and conda - Python package browser, written in PowerShell.

This script helps to keep packages updated.


1. Hit Check for Updates
2. Choose packages
3. Select *Update* action and hit Execute


## features


- Dependency-free
- Search and install from **pip**, **conda** and **github**
- Looks up for all installed Python distributions
- Filter and sort packages
- View package dependecy tree
- Manage virtual environments
- Adds firewall rules for venv
- Manage environment variables for venv
- Documentation viewer with simple highlighting and browser-like navigation
- Completion for packages, versions, paths, git tags, PyDoc topics
- Package name typosquatting check

![](screenshot.png)


## shortcuts


| Keys                             | Action                                                                     |
| -------------------------------- | -------------------------------------------------------------------------- |
| Enter or Space                   | Toggle package selection                                                   |
| F1 .. F12                        | Choose a current action, in the same order it appears in the dropdown list |
| Shift+Enter (on package table)   | Execute action for selected packages                                       |
| Shift+Enter (install window)     | Fuzzy name search                                                          |
| Ctrl+Enter or Double click       | Open package home page in browser                                          |
| Escape                           | Clear filter, or Switch between filter and package table                   |
| Enter (on Filter)                | Focus on packages editable                                                 | 
| Shift+Mouse Hover                | Fetch package details for a tooltip                                        |
| Ctrl-C (on interpreter combobox) | Copy full python executable path                                           |
| Delete (on interpreter combobox) | Remove virtual env entry added by user, with confirmation                  |
| / (in PyDoc browser)             | Toggle search                                                              |


## todo


- [ ] Typo checking with distance of 1 while typing; with distance of 2 when requested more results by user
- [X] Save user env list
- [X] Add search with Github API
- [X] Add editable packages to Install dialog (git, local); pip list --editable; Appropriate checkbox is still needed in the install dialog
- [X] Mistype verification with Damerau–Levenshtein distance (fuzzy). As PowerShell v3 is JIT-compiled, no need for external DLLs
- [X] Dependency tree & pinning with deps
- [ ] Virtualenv creation help for user if neither virtualenv nor pipenv packages are installed
- [X] Delete user envs on Del Pressed with confirmation
- [ ] Sort fuzzy candidates by PyPI *rdeps count*, <s>*download count*</s> moved from Warehouse to [Google](https://mail.python.org/pipermail/distutils-sig/2016-May/028986.html) because of CDN
- [ ] <s>GPG signature verification for packages (with gpg.exe)</s> deprecated in Warehouse (unclear)
- [ ] <s>Add some integration with VirusTotal (sha-256 of archive + link to VT for a starter)</s>
- [ ] Fix "isolated" checkbox behavior is somewhat uncertain
- [X] <s>Move known package index builder code into the main script</s> moved to [BK-tree](https://github.com/ptytb/BK-tree)
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
- [ ] Fix freezing while running background tasks by means of using pipes and threads


## typosquatting check


**pips** has a feature of protection from [typosquatting](https://en.wikipedia.org/wiki/Typosquatting) of package names.
It assists to an unprepared user to explore and install packages instantly, without wasting
time to figure out if a package name is spelled properly, or is it popular, genuine or malicious package.

This is being achieved by using the following algorithms:

1. Search for package name candidates using [Levenshtein distance and BK-tree](https://github.com/ptytb/BK-tree)
2. Sort and filter these candidates using the index built with following parameters:
	- Number of connections with other packages using dependency graph
	- Package's first release date
	- Number of releases
	- Average time interval before the next release
	- Average number of downloads per release
3. Search through the index
   
   - Search for reverse dependencies using *Adjacency matrix*


More details about how it works [here]().


## trademarks


The Python logo used in this program is a trademark of [The Python Software Foundation](https://www.python.org/psf/trademarks/).
This program is written in the PowerShell programming language and has a few pieces of inline Python code, and relies on external Python executables.


## similar projects and related links


[pypi-cli](https://github.com/sloria/pypi-cli) a command-line interface to the Python Package Index.
[yip](https://github.com/balzss/yip) a frontend for searching PyPI, a feature rich alternative to pip search

[pipreqs](https://github.com/bndr/pipreqs) a tool to generate requirements.txt file based on imports of any project
[pigar](https://github.com/Damnever/pigar) a tool to generate requirements file for your Python project


#### pip security


[safety](https://github.com/pyupio/safety) check your installed dependencies for known security vulnerabilities
[Pytosquatting](https://www.pytosquatting.org/) fixing typosquatting+namesquatting threats in Python Package Index
[auditwheel](https://github.com/pypa/auditwheel) auditing and relabeling cross-distribution Linux wheels


# License

Copyright, 2018, Ilya Pronin.
This code is released under the MIT license.
