# grep-docx

A command-line tool for searching through Microsoft Word (.docx) files using grep-like pattern matching. Written in Python, this tool allows you to search for text patterns across one or multiple .docx files.

## Features

- Search single files or recursively through directories
- Regular expression pattern matching
- Case-sensitive and case-insensitive search options
- Colored output highlighting matches
- Output formatting options including hanging indents
- Count-only and filename-only output modes
- Write results to file or stdout
- Debug logging support

## Requirements

- Python 3.x
- [python-docx library](https://github.com/python-openxml/python-docx)

## Installation

1. Ensure Python 3.x is installed
2. Install required dependencies:
```bash
pip install python-docx
```
3. Run the script.

## Usage

```bash
grep-docx [options] PATTERN PATH 
```

### Arguments

- `PATTERN`: Regular expression pattern to search for
- `PATH`: File or directory to search

### Options

- `-C, --color, --colour`: Color the prefix and highlight matches
- `-c, --count`: Only print a count of matching lines
- `-H, --hyperlink`: The name of each file is printed as a hyperlink that launches Word.  (Your terminal may not support this.)
- `-I, --hanging-indent`: Line output after the 1st line starts with a tab character
- `-i, --ignore-case`: Ignore case distinctions
- `-l, --files-with-matches`: Only print names of files with matches
- `-L, --files-without-matches`: Only print names of files without matches
- `-q, --quiet, --silent`: Suppress all normal output
- `-r, --recursive`: Recursively search subdirectories
- `-T, --initial-tab`: Line output starts with a tab character
- `-V, --version`: Show program version
- `--debug`: Enable debug logging

### Examples

Search a single file:
```bash
grep-docx "pattern" document.docx
```

Search recursively with case insensitive matching:
```bash
grep-docx -ri "pattern" ./documents/
```

Count matches in a directory:
```bash
grep-docx -c "pattern" ./documents/
```

List files containing matches:
```bash
grep-docx -l "pattern" ./documents/
```

## License

This project is licensed under the MIT License - see the [`LICENSE`](LICENSE) file for details.


