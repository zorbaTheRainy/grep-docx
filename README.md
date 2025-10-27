# grep-docx

A command-line tool for searching through Microsoft Word (*.docx) files using [grep](https://en.wikipedia.org/wiki/Grep)-like pattern matching. Written in Python, this tool allows you to search for text patterns across one or multiple .docx files.  Non-.docx files are simply ignored.

`grep-docx` will not work on old *.doc (no-X) binary files, used before Office 2007.  Nor, will it work on *.dotx (templates) or *.domx (macros) files.  This is a limitation of the Python python-docx library.

The basic "line" to search in the Word file is 1 paragraph.  Word files are broken into paragraphs and then searched.  Bulleted lists, etc., are composed of multiple paragraphs.  If you want complex search patterns, plan accordingly.

## Features

* Search single files or recursively through directories
* [Regular expression pattern matching](https://www.regexone.com/)
* Case-sensitive and case-insensitive search options
* Colored output highlighting matches
* Output formatting options including hanging indents
* Count-only and filename-only output modes
* Debug logging support

## Usage

``` bash
grep-docx [options] PATTERN PATH
```

### Arguments

* `PATTERN`: Regular expression pattern to search for (a Python-style regex)
* `PATH`: File or directory to search

### Options

* `-C, --color, --colour`: Color the prefix and highlight matches
* `-c, --count`: Only print a count of matching lines
* `-H, --hyperlink`: The name of each file is printed as a hyperlink that launches Word. (Your terminal may not support this.)
* `-I, --hanging-indent`: Line output after the 1st line starts with a tab character
* `-i, --ignore-case`: Ignore case distinctions
* `-l, --files-with-matches`: Only print names of files with matches
* `-L, --files-without-matches`: Only print names of files without matches
* `-q, --quiet, --silent`: Suppress all normal output
* `-r, --recursive`: Recursively search subdirectories
* `-s, --no-messages`: Suppress error messages about nonexistent or unreadable files
* `-T, --initial-tab`: Line output starts with a tab character
* `-V, --version`: Show program version
* `--debug`: Enable debug logging

### Examples

Search a single file for the text 'config':

``` bash
grep-docx config document.docx
```

Search folders/directories recursively with case insensitive matching:

``` bash
grep-docx -ri config ./documents/
```

Count the number of times in a directory the partial-word 'construc' (e.g., construct, misconstruction, constructible, etc.) is found in *.docx files:

``` bash
grep-docx -c construc ./documents/
```

List files where the words 'sake' and 'clarification' occur within the same paragraph:

``` bash
grep-docx -l '\bsake\b.*clarification\b|\bclarification\b.*sake\b' ./documents/
```

Write output to a file (via STDOUT redirection):

``` bash
grep-docx -ri config ./documents/ > output.txt
```

## Executables

Compiled executables may be found in the `executables` branch of this repository.

Executables were made for:

##### Windows

* [grep-docx.exe](https://github.com/zorbaTheRainy/grep-docx/raw/refs/tags/v1.0.0/executables/Windows/grep-docx.exe) \- Windows 64\-bit Intel executable

##### macOS

* [grep-docx](https://github.com/zorbaTheRainy/grep-docx/raw/refs/tags/v1.0.0/executables/MacOS/grep-docx) \- Apple Silicon \(M1\+\) binary

##### Linux

* [grep-docx](https://github.com/zorbaTheRainy/grep-docx/raw/refs/tags/v1.0.0/executables/Linux/x86/grep-docx) \- Linux 64\-bit Intel \(x86\) binary
* [grep-docx](https://github.com/zorbaTheRainy/grep-docx/raw/refs/tags/v1.0.0/executables/Linux/ARM/grep-docx) \- Linux 64\-bit ARM binary

### Requirements

* No Python installation required - standalone executables
* Operating system: Windows 7+, macOS 10.15+, or Linux with glibc 2.17+

### Installation

1. Download
2. For Linux or MacOS, `chmod +x grep-docx`
3. Run

## Native Script

A simple Python script, ready to run.

### Requirements

* Python 3.x
* [python-docx library](https://github.com/python-openxml/python-docx)

### Installation

1. Ensure Python 3.x is installed
2. Install required dependencies:

``` bash
pip install python-docx
```

3. Run the script.

## License

This project is licensed under the MIT License - see the [`LICENSE`](LICENSE) file for details.