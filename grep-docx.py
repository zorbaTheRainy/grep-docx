#!/usr/bin/env python3

import argparse    # https://docs.python.org/3/library/argparse.html
import logging     # https://docs.python.org/3/library/logging.html
import os          # https://docs.python.org/3/library/os.html
import re          # https://docs.python.org/3/library/re.html
import shutil      # https://docs.python.org/3/library/shutil.html
import sys         # https://docs.python.org/3/library/sys.html
import textwrap    # https://docs.python.org/3/library/textwrap.html
from docx import Document # https://python-docx.readthedocs.io/     # Install via: pip install python-docx

# --------------------------------
# Global variables
VERSION = "0.8.0"
COLORS = { # ANSI color codes
    'BLACK': '30',
    'RED': '31',
    'GREEN': '32',
    'YELLOW': '33',
    'BLUE': '34',
    'MAGENTA': '35',
    'CYAN': '36',
    'WHITE': '37',
    # Bright variants
    'BRIGHT_BLACK': '90',
    'BRIGHT_RED': '91',
    'BRIGHT_GREEN': '92',
    'BRIGHT_YELLOW': '93',
    'BRIGHT_BLUE': '94',
    'BRIGHT_MAGENTA': '95',
    'BRIGHT_CYAN': '96',
    'BRIGHT_WHITE': '97',
}

def setup_logging(debug=False):
    """Configures logging settings."""
    level = logging.DEBUG if debug else logging.INFO
    logging.basicConfig(format='%(levelname)s: %(message)s', level=level)

def colorize(text, color_code):
    """Wrap text in ANSI color codes."""
    return f"\033[{color_code}m{text}\033[0m"

def highlight_matches(text, pattern, color_code, ignore_case):
    """Highlight all matches of pattern in text with the given color, respecting case sensitivity."""
    flags = re.IGNORECASE if ignore_case else 0
    regex = re.compile(pattern, flags)
    def replacer(match):
        return colorize(match.group(0), color_code)
    return regex.sub(replacer, text)

def search_file(file_path, regex, config):
    """
    Searches a single .docx file for matches to the regex pattern.

    Args:
        file_path (str): Path to the .docx file.
        regex (re.Pattern): Compiled regex pattern.
        config (dict): Configuration dictionary with options.

    Returns:
        list: List of matching lines with hanging indent formatting and color.
        bool: True if any match found, False otherwise.
    """
    matches = []
    matched = False
    try:
        doc = Document(file_path)
        for i, para in enumerate(doc.paragraphs):
            if regex.search(para.text):
                matched = True

                if config["quiet"] or config["list_unmatched_files"]:
                    return matches, matched  # exit ASAP if quiet mode is enabled or all we care about is if there was a match

                prefix = f"{file_path} [Paragraph {i+1}]: "
                if config.get("color"):
                    prefix_colored = colorize(prefix, COLORS['GREEN'])  # Color for prefix
                    para_text = highlight_matches(
                        para.text, regex.pattern, COLORS['RED'], config["ignore_case"] # Color for match
                    )  
                else:
                    prefix_colored = prefix
                    para_text = para.text

                if config.get("initial_tab"):
                    prefix_colored = "\t" + prefix_colored

                if config.get("hanging_indent"):
                    term_width = shutil.get_terminal_size((80, 20)).columns
                    wrap_width = term_width - len(prefix)
                    indent = "\t"
                    wrapped_text = textwrap.fill(
                        para_text,
                        width=wrap_width,
                        initial_indent='',
                        subsequent_indent=indent
                    )
                    formatted_line = prefix_colored + wrapped_text
                    matches.append(formatted_line)
                else:
                    formatted_line = prefix_colored + para_text
                    matches.append(formatted_line)

                logging.debug(f"Match found in {file_path} at paragraph {i+1}")
    except Exception as e:
        logging.error(f"Error reading {file_path}: {e}")
    # logging.debug(f"matched: {matched}")
    # logging.debug(f"matches: {matches}")
    return matches, matched

def grep_docx(config):
    """
    Searches for a pattern in .docx files based on configuration.

    Args:
        config (dict): Configuration dictionary with options.
    """
    flags = re.IGNORECASE if config["ignore_case"] else 0
    regex = re.compile(config["pattern"], flags)
    matches = []
    match_count = 0
    matched_files = {}
    unmatched_files = set()
    path = config["path"]
    if os.path.isfile(path):
        if path.endswith(".docx"):
            full_path = path # rename var to match the directory case, for ease of maintenance
            logging.debug(f"Searching file: {full_path}")
            file_matches, matched = search_file(full_path, regex, config)
            if matched:
                if config["quiet"]:
                    # exit with status 0 (success)
                    sys.exit(0)
                match_count += len(file_matches)
                matched_files[full_path] = len(file_matches)
                matches.extend(file_matches)
            else:
                unmatched_files.add(full_path)
    elif os.path.isdir(path):
        for root, _, files in os.walk(path):
            for file in files:
                if file.endswith(".docx"):
                    full_path = os.path.join(root, file)
                    logging.debug(f"Searching file: {full_path}")
                    file_matches, matched = search_file(full_path, regex, config)
                    if matched:
                        if config["quiet"]:
                            # exit with status 0 (success)
                            sys.exit(0)
                        match_count += len(file_matches)
                        matched_files[full_path] = len(file_matches)
                        matches.extend(file_matches)
                    else:
                        unmatched_files.add(full_path)
            if not config["recursive"]:
                break
    else:
        logging.error(f"Invalid path: {path}")
        return

    if config["quiet"]:
        # exit with status 1 (failure)
        sys.exit(1)
    # elif config["output"]:
    #     with open(config["output"], 'w', encoding='utf-8') as f:
    #         for line in matches:
    #             f.write(line + '\n')
    elif config["list_unmatched_files"]:
        for f in unmatched_files:
            print(f)
    elif config["list_matched_files"]:
        if config["count"]:
            for f, cnt in matched_files.items():
                print(f"{f}: {cnt}")
            print(match_count)
        else:
            for f, cnt in matched_files.items():
                print(f"{f}")
    elif config["count"]:
        print(match_count)
    else:
        for line in matches:
            print(line)

def main():
    """Parses command-line arguments and initiates the grep-like search."""
    parser = argparse.ArgumentParser(description="Search for PATTERN in .docx files like grep.")
    parser.add_argument("pattern", help="Regex pattern to search for")
    parser.add_argument("path", help="File or directory to search")
    parser.add_argument("-i", "--ignore-case", action="store_true", help="Ignore case distinctions")
    parser.add_argument("-r", "--recursive", action="store_true", help="Recursively search subdirectories")
    # parser.add_argument("-o", "--output", help="Write results to a file instead of stdout")
    parser.add_argument("-c", "--count", action="store_true", help="Only print a count of matching lines")
    parser.add_argument("-l", "--files-with-matches", action="store_true", help="Only print names of files with matches")
    parser.add_argument("-L", "--files-without-matches", action="store_true", help="Only print names of files without matches")
    parser.add_argument("-q", "--quiet", "--silent", action="store_true", help="Suppress all normal output")
    parser.add_argument("-T", "--initial-tab", action="store_true", help="Line output starts with a tab character")
    parser.add_argument("-I", "--hanging-indent", action="store_true", help="Line output after the 1st line starts with a tab character")
    parser.add_argument("-C", "--color", "--colour", action="store_true", help="Color the prefix and highlight matches")
    parser.add_argument("-V", "--version", action="version", version=f"%(prog)s {VERSION}")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    args = parser.parse_args()
    setup_logging(args.debug)
    config = {
        "pattern": args.pattern,
        "path": args.path,
        "ignore_case": args.ignore_case,
        "recursive": args.recursive,
        # "output": args.output,
        "count": args.count,
        "list_matched_files": args.files_with_matches,
        "list_unmatched_files": args.files_without_matches,
        "quiet": args.quiet,
        "hanging_indent": args.hanging_indent,
        "initial_tab": args.initial_tab,
        "color": args.color
    }
    grep_docx(config)

if __name__ == "__main__":
    main()