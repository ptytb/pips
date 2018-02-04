# pips
**pips** is a GUI for pip and conda - Python package manager, written in PowerShell.

This tiny script helps to keep packages updated.


1. Hit Check for Updates
2. Choose packages
3. Select *Update* action and hit Execute

## features


- Dependency-free
- **pip** and **conda** are supported
- Looks up for installed Python distributions
- Filter and sort of packages
- Documentation viewer with simple highlighting
- Manage user virtualenv directories


![](screenshot.png)


## todo


- [ ] Now assuming *utf-8* for IO whilst old(?) Python versions use cp1252, which generates empty strings for Chinese, Russian, etc.
- [ ] Filter package search results for conda by relevant Architecture
- [ ] Implement some missing conda commands
