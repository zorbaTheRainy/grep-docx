# grep-docx

A command-line tool for searching through Microsoft Word (*.docx) files using [grep](https://en.wikipedia.org/wiki/Grep)-like pattern matching. 

**Note**: This tool will not work on:
- Old .doc (pre-Office 2007) binary files
- .dotx template files
- .docm macro-enabled files

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

Compiled executables may be found in the `executables` branch of this repository.

Executables were made for:

##### Windows

* [grep-docx.exe](https://github.com/zorbaTheRainy/grep-docx/raw/refs/heads/executables/executables/Windows/grep-docx.exe) \- Windows 64\-bit Intel executable

##### macOS

* [grep-docx](https://github.com/zorbaTheRainy/grep-docx/raw/refs/heads/executables/executables/MacOS/grep-docx) \- Apple Silicon \(M1\+\) binary

##### Linux

* [grep-docx](https://github.com/zorbaTheRainy/grep-docx/raw/refs/heads/executables/executables/Linux/x86/grep-docx) \- Linux 64\-bit Intel \(x86\) binary
* [grep-docx](https://github.com/zorbaTheRainy/grep-docx/raw/refs/heads/executables/executables/Linux/ARM/grep-docx) \- Linux 64\-bit ARM binary

### Requirements

* No Python installation required - standalone executables
* Operating system: Windows 7+, macOS 10.15+, or Linux with glibc 2.17+

### Installation

1. Download
2. For Linux or MacOS, `chmod +x grep-docx`
3. MacOS, as the binary is not signed with a Developer account, it will require special permissions.  MacOS will walk you through this if you read the dialog box instructions.
4. Run

## Native Script

A simple Python script, ready to run.

### Requirements

* Python 3.x
* [python-docx](https://github.com/python-openxml/python-docx) library
* [tqdm](https://tqdm.github.io/) library
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


3. Run the script.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.