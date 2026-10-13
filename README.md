# grep-docx

A command-line tool for searching through Microsoft Word files using [grep](https://en.wikipedia.org/wiki/Grep)-like pattern matching. 

**Supported formats**: `.docx`, `.dotx`, `.docm`, and `.doc` (pre-Office 2007).

**Note**: `.doc` file support requires Python 3.11+ and the [legacy-doc](https://github.com/lcorrigan704/legacy-doc) package. If not available, `.doc` files are silently skipped.

The basic "line" to search in the Word file is 1 paragraph.  Word files are broken into paragraphs and then searched.  Bulleted lists, etc., are composed of multiple paragraphs.  If you want complex search patterns, plan accordingly.

## Features

- Search single files or recursively through directories
- [Regular expression](https://www.regexone.com/) pattern matching
- Case-sensitive and case-insensitive search options
- Colored output highlighting matches
- Output formatting options including hanging indents
- Count-only and filename-only output modes
- Hyperlinked file paths (in supported terminals)
- Progress bar for large searches
- Debug logging support
- Read paths from stdin

## Installation

See [Executables](#executables) or [Native Script](#native-script) below.

## Usage

``` bash
grep-docx [options] PATTERN PATH
```

### Arguments

- `PATTERN`: Regular expression pattern to search for (a Python-style regex)
- `PATH`: File or directory to search (use - to read paths from stdin)

### Options

- `-C, --color, --colour`: Color the prefix and highlight matches
- `-c, --count`: Only print a count of matching lines
- `-H, --hyperlink`: Print filenames as clickable hyperlinks
- `-I, --hanging-indent`: Line output after the 1st line starts with a tab
- `-i, --ignore-case`: Ignore case distinctions
- `-l, --files-with-matches`: Only print names of files with matches
- `-L, --files-without-matches`: Only print names of files without matches
- `-P, --no-progress-bar`: Disable the progress bar
- `-q, --quiet, --silent`: Suppress all normal output
- `-r, --recursive`: Recursively search subdirectories
- `-s, --no-messages`: Suppress error messages
- `-T, --initial-tab`: Line output starts with a tab character
- `-V, --version`: Show program version
- `--debug`: Enable debug logging
- `--logfile FILE`: Write logs to FILE

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

Compiled executables are available in the [latest release](https://github.com/zorbaTheRainy/grep-docx/releases/latest).

##### Windows

* [grep-docx-windows-x64.zip](https://github.com/zorbaTheRainy/grep-docx/releases/latest/download/grep-docx-windows-x64.zip) \- Windows 64\-bit Intel executable

##### macOS

* [grep-docx-macos-arm64.tar.gz](https://github.com/zorbaTheRainy/grep-docx/releases/latest/download/grep-docx-macos-arm64.tar.gz) \- Apple Silicon \(M1\+\) binary

##### Linux

* [grep-docx-linux-x64.tar.gz](https://github.com/zorbaTheRainy/grep-docx/releases/latest/download/grep-docx-linux-x64.tar.gz) \- Linux 64\-bit Intel \(x86\) binary
* [grep-docx-linux-arm64.tar.gz](https://github.com/zorbaTheRainy/grep-docx/releases/latest/download/grep-docx-linux-arm64.tar.gz) \- Linux 64\-bit ARM binary

### Requirements

* No Python installation required - standalone executables
* Operating system: Windows 7+, macOS 10.15+, or Linux with glibc 2.17+

### Installation

1. Download the archive for your platform from the [latest release](https://github.com/zorbaTheRainy/grep-docx/releases/latest)
2. Extract the archive
3. For Linux or macOS, `chmod +x grep-docx`
4. macOS: as the binary is not signed with a Developer account, it will require special permissions.  macOS will walk you through this if you read the dialog box instructions.
5. Run

## Native Script

A simple Python script, ready to run.

### Requirements

* Python 3.x
* [python-docx](https://github.com/python-openxml/python-docx) library
* [tqdm](https://tqdm.github.io/) library
* (optional, Python 3.11+) [legacy-doc](https://github.com/lcorrigan704/legacy-doc) library for `.doc` support
* (optional & Windows only) [colorama](https://github.com/tartley/colorama) library


### Installation

1. Ensure Python 3.x is installed
2. Install required dependencies:

``` bash
pip install python-docx tqdm
```
or
``` bash
pip install -r requirements.txt
```

For `.doc` file support (requires Python 3.11+):
``` bash
pip install legacy-doc
```


3. Run the script.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.