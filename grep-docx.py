#!/usr/bin/env python3

import argparse    # https://docs.python.org/3/library/argparse.html
import logging     # https://docs.python.org/3/library/logging.html
import os          # https://docs.python.org/3/library/os.html
import re          # https://docs.python.org/3/library/re.html
import shutil      # https://docs.python.org/3/library/shutil.html
import sys         # https://docs.python.org/3/library/sys.html
import textwrap    # https://docs.python.org/3/library/textwrap.html

# need to install with pip
from docx import Document # https://python-docx.readthedocs.io/     # Install via: pip install python-docx

# --------------------------------
# Global variables
VERSION = "0.9.0"
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

# -----------------------------------------------------------------------
def parse_args():
    """Parses command-line arguments."""
    parser = argparse.ArgumentParser(description="Search for PATTERN in .docx files like grep.")
    parser.add_argument("pattern", help="Regex pattern to search for")
    parser.add_argument("path", help="File or directory to search")
    parser.add_argument("-C", "--color", "--colour", action="store_true", help="Color the prefix and highlight matches")
    parser.add_argument("-c", "--count", action="store_true", help="Only print a count of matching lines")
    parser.add_argument("-H", "--hyperlink", action="store_true", help="The name of each file is printed as a hyperlink that launches Word.  (Your terminal may not support this.)")
    parser.add_argument("-I", "--hanging-indent", action="store_true", help="Line output after the 1st line starts with a tab character")
    parser.add_argument("-i", "--ignore-case", action="store_true", help="Ignore case distinctions")
    parser.add_argument("-l", "--files-with-matches", action="store_true", help="Only print names of files with matches")
    parser.add_argument("-L", "--files-without-matches", action="store_true", help="Only print names of files without matches")
    parser.add_argument("-q", "--quiet", "--silent", action="store_true", help="Suppress all normal output")
    parser.add_argument("-r", "--recursive", action="store_true", help="Recursively search subdirectories")
    parser.add_argument("-s", "--no-messages", action="store_true", help="Suppress error messages about nonexistent or unreadable files")
    parser.add_argument("-T", "--initial-tab", action="store_true", help="Line output starts with a tab character")
    parser.add_argument("-V", "--version", action="version", version=f"%(prog)s {VERSION}")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")

    args = parser.parse_args()
    return args

# -----------------------------------------------------------------------
def main():
    """ Searches for a pattern in .docx files based on configuration. """

    # Parse command line arguments
    args = parse_args()
    
    # turn on logs if requested
    setup_logging(args.debug)

    # prepare to search
    flags = re.IGNORECASE if args.ignore_case else 0
    regex = re.compile(args.pattern, flags)

    results = {
        "matches": [],
        "match_count": 0,
        "matched_files": {},
        "unmatched_files": set(),
    }
    
    # perform the search
    path = args.path
    if os.path.isfile(path):
        if path.endswith(".docx"):
            results = process_file(path, regex, args, results)

    elif os.path.isdir(path):
        for root, _, files in os.walk(path):
            for file in files:
                if file.endswith(".docx"):
                    results = process_file(os.path.join(root, file), regex, args, results)
            if not args.recursive:
                break

    else:
        if not (args.quiet or args.no_messages):
            logging.error(f"Invalid path: {path}")
        sys.exit(1) # exit with status 1 (failure)

    #  finally, print results
    print_results(results, args)

# -----------------------------------------------------------------------
def setup_logging(debug=False):
    """Configure the root logger.

    Args:
        debug (bool): If True sets logging level to DEBUG, otherwise INFO.

    This sets a simple log format of "LEVEL: message".
    """
    level = logging.DEBUG if debug else logging.INFO
    logging.basicConfig(format='%(levelname)s: %(message)s', level=level)

def process_file(file_path, regex, args, results):
    """Search a single file and update the aggregated results.

    Args:
        file_path (str): Path to a .docx file to search.
        regex (re.Pattern): Compiled regular expression to search with.
        args (argparse.Namespace): Parsed CLI arguments.
        results (dict): Aggregated results dict modified in-place. Expected keys:
            - "matches": list of formatted match lines
            - "match_count": int total matches found
            - "matched_files": dict mapping file_path -> match count
            - "unmatched_files": set of file paths without matches

    Returns:
        dict: The updated results dict (same object passed in).
    """
    logging.debug(f"Searching file: {file_path}")

    # perform the actual file search
    file_matches, matched = search_file(file_path, regex, args)

    # change results based on match
    if matched:
        if args.quiet:
            sys.exit(0) # exit with status 0 (success)
        results["match_count"] += len(file_matches)
        results["matched_files"][file_path] = len(file_matches)
        results["matches"].extend(file_matches)
    else:
        results["unmatched_files"].add(file_path)

    return results

def search_file(file_path, regex, args):
    """
    Searches a single .docx file for matches to the regex pattern.

    Args:
        file_path (str): Path to the .docx file.
        regex (re.Pattern): Compiled regex pattern.
        args (Namespace): Parsed command-line arguments.

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

                if args.quiet or args.files_without_matches or (args.files_with_matches and not args.count):
                    return matches, matched  # exit ASAP if quiet mode is enabled or all we care about is if there was a match

                if args.hyperlink:
                    hyper_path = make_hyperlink(file_path)
                    prefix = f"{hyper_path} [Paragraph {i+1}]: "
                else:
                    prefix = f"{file_path} [Paragraph {i+1}]: "

                if args.color:
                    prefix_colored = colorize(prefix, COLORS['GREEN'])  # Color for prefix
                    para_text = highlight_matches(para.text, regex.pattern, COLORS['RED'], args.ignore_case) # Color for match
                else:
                    prefix_colored = prefix
                    para_text = para.text

                if args.initial_tab:
                    prefix_colored = "\t" + prefix_colored

                if args.hanging_indent:
                    term_width = shutil.get_terminal_size((80, 20)).columns
                    wrap_width = term_width - 8
                    indent = "\t"
                    wrapped_text = textwrap.fill(
                        para_text,
                        width=wrap_width,
                        initial_indent=indent,
                        subsequent_indent=indent,
                        break_on_hyphens=False
                    )
                    formatted_line = prefix_colored + os.linesep + wrapped_text
                    matches.append(formatted_line)
                else:
                    formatted_line = prefix_colored + para_text
                    matches.append(formatted_line)

                logging.debug(f"Match found in {file_path} at paragraph {i+1}")
    except Exception as e:
        if not (args.quiet or args.no_messages):
            logging.error(f"Error reading {file_path}: {e}")

    return matches, matched

def make_hyperlink(path, label=None):
    """
    Wrap `path` in an OSC 8 hyperlink sequence.
    - path: filesystem path to your .docx
    - label: visible text; defaults to path
    """
    # Ensure absolute URI
    uri = f"file://{os.path.abspath(path)}"
    label = label or path
    # OSC 8 sequence: \033]8;;URI\033\\TEXT\033]8;;\033\\
    return (
        f"\033]8;;{uri}\033\\"
        f"{label}"
        f"\033]8;;\033\\"
    )

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

def print_results(results, args):
    """Print aggregated search results according to CLI options.

    Args:
        results (dict): Aggregated results produced by main/process_file:
            - "matches": list of formatted match lines
            - "match_count": int total matches
            - "matched_files": dict mapping path -> match count
            - "unmatched_files": set of file paths
        args (argparse.Namespace): Parsed CLI arguments controlling output modes.

    Behavior:
        - Exits or prints depending on args.quiet, args.count, args.files_with_matches,
          args.files_without_matches, args.hyperlink, etc.
    """
    # unpack results to renamed variables
    match_count     = results["match_count"]
    matched_files   = results["matched_files"]
    matches         = results["matches"]
    unmatched_files = results["unmatched_files"]
    
    # actually print
    if args.quiet:
        # if we got here and quiet-mode is on, we found nothing; so, it is a failure
        sys.exit(1)  # exit with status 1 (failure)
    elif args.files_without_matches:
        if args.hyperlink:
            for f in unmatched_files:
                print(make_hyperlink(f))
        else:
            for f in unmatched_files:
                print(f)
    elif args.files_with_matches:
        for f, cnt in matched_files.items():
            if args.hyperlink:
                f = make_hyperlink(f)
            if args.count:
                print(f"{f}: {cnt}")
            else:
                print(f"{f}")
        if args.count:
            print(match_count)
    elif args.count:
        print(match_count)
    else:
        for line in matches:
            print(line)
    
    return


# -----------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------
if __name__ == "__main__":
    main()