#!/usr/bin/env python3

import argparse    # https://docs.python.org/3/library/argparse.html
import glob        # https://docs.python.org/3/library/glob.html
import logging     # https://docs.python.org/3/library/logging.html
import os          # https://docs.python.org/3/library/os.html
import re          # https://docs.python.org/3/library/re.html
import shutil      # https://docs.python.org/3/library/shutil.html
import sys         # https://docs.python.org/3/library/sys.html
import textwrap    # https://docs.python.org/3/library/textwrap.html
import unicodedata # https://docs.python.org/3/library/unicodedata.html
from pathlib import Path # https://docs.python.org/3/library/pathlib.html

# need to install with pip
from docx import Document # https://python-docx.readthedocs.io/     # Install via: pip install python-docx
from tqdm import tqdm     # https://tqdm.github.io/                 # Install via: pip install tqdm
try:
    # only needed for Windows.  Put in requirements.txt  --> colorama; sys_platform == "win32"
    from colorama import just_fix_windows_console # https://github.com/tartley/colorama # Install via:  pip install colorama
    HAVE_COLORAMA = True
except ImportError:
    HAVE_COLORAMA = False


# --------------------------------
# Global variables
VERSION = "1.0.9"
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
    parser.add_argument("paths", nargs="+", help="One or more files or directories to search; use - to read paths from stdin (one per line)")
    parser.add_argument("-C", "--color", "--colour", action="store_true", help="Color the prefix and highlight matches")
    parser.add_argument("-c", "--count", action="store_true", help="Only print a count of matching lines")
    parser.add_argument("-H", "--hyperlink", action="store_true", help="The name of each file is printed as a hyperlink that launches Word.  (Your terminal may not support this.)")
    parser.add_argument("-I", "--hanging-indent", action="store_true", help="Line output after the 1st line starts with a tab character")
    parser.add_argument("-i", "--ignore-case", action="store_true", help="Ignore case distinctions")
    parser.add_argument("-l", "--files-with-matches", action="store_true", help="Only print names of files with matches")
    parser.add_argument("-L", "--files-without-matches", action="store_true", help="Only print names of files without matches")
    parser.add_argument("-P", "--no-progress-bar", action="store_true", help="Do not display the progress bar")
    parser.add_argument("-q", "--quiet", "--silent", action="store_true", help="Suppress all normal output")
    parser.add_argument("-r", "--recursive", action="store_true", help="Recursively search subdirectories")
    parser.add_argument("-s", "--no-messages", action="store_true", help="Suppress error messages about nonexistent or unreadable files")
    parser.add_argument("-T", "--initial-tab", action="store_true", help="Line output starts with a tab character")
    parser.add_argument("-V", "--version", action="version", version=f"%(prog)s {VERSION}")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    parser.add_argument("--logfile", metavar="FILE", help="Write log output to FILE. Use - for stdin-sourced paths (see paths argument)")

    args = parser.parse_args()
    return args

# -----------------------------------------------------------------------
def main():
    """ Searches for a pattern in .docx files based on configuration. """

    # Parse command line arguments
    args = parse_args()
    
    # turn on logs if requested; also configure file logging if requested
    setup_logging(args.debug, args.quiet, logfile=args.logfile)

    # normal Windows consoles don't natively display ANSI colors; fix that   ; OS detection is in the subroutine already.
    if HAVE_COLORAMA:
        just_fix_windows_console()
    # if hyperlinks are requested, check if supported; disable if not
    if args.hyperlink:
        if not supports_hyperlink():
            suggest_terminals_if_no_hyperlink()
            # logging.warning("Terminal does not appear to support hyperlinks; disabling --hyperlink.")
            args.hyperlink = False

    # prepare to search
    flags = re.IGNORECASE if args.ignore_case else 0
    regex = re.compile(args.pattern, flags)

    results = {
        "matches": [],
        "match_count": 0,
        "matched_files": {},
        "unmatched_files": set(),
    }
    
    # '-' in paths means "read a list of paths from stdin, one per line".
    if "-" in args.paths:
        args.paths.remove("-")
        for line in sys.stdin:
            s = line.strip()
            if s:
                args.paths.append(s)
    
    # obtain list of all the files to search from one or more input paths
    file_list = []
    for path in args.paths:
        for globed_path in glob.glob(path):
            if not os.path.exists(globed_path):
                if not (args.quiet or args.no_messages):
                    logging.error(f"Invalid path: {globed_path}")
                continue
            file_list.extend(get_file_list(globed_path, args.recursive))

    # If no files found across all provided paths, exit
    if not file_list:
        if not (args.quiet or args.no_messages):
            logging.error("No .docx files found to search.")
        sys.exit(2) # exit with status 2 (failure with an error)

    # setup the progress bar (disabled based on args.quiet and other flags)
    disable_flag = args.quiet or args.no_progress_bar
    try:
        iterator = tqdm( file_list, \
                    total=len(file_list), \
                    desc="Searching", \
                    leave=False, \
                    unit="file", \
                    disable=disable_flag)
        for file in iterator:
            results = process_file(file, regex, args, results)
    finally:
        if iterator is not None:
            iterator.close()

    #  finally, print results
    print_results(results, args)

# -----------------------------------------------------------------------
def setup_logging(debug=False, quiet=False, logfile=None):
    """Configure the root logger.

    Args:
        debug (bool): If True sets logging level to DEBUG, otherwise INFO.
        logfile (str|None): If provided, path to a file to append/write logs to.
                            If None, logs remain on stderr.
    """
    level = logging.DEBUG if debug else logging.INFO
    handlers = []

    # Console handler (stderr)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
    handlers.append(console_handler)

    # If logfile is provided, add a FileHandler. If logfile is '-', treat that as literal filename '-'
    if logfile:
        try:
            file_handler = logging.FileHandler(logfile, mode="a", encoding="utf-8")
            file_handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s: %(message)s', "%Y-%m-%d %H:%M:%S"))
            handlers.append(file_handler)
        except Exception as e:
            # If we cannot open the logfile, still proceed but warn on stderr via basicConfig fallback
            logging.basicConfig(format='%(levelname)s: %(message)s', level=level)
            if not quiet:
                logging.warning(f"Could not open logfile '{logfile}' for writing: {e}")
            return

    # Apply handlers to root logger
    root = logging.getLogger()
    root.setLevel(level)
    # Remove existing handlers to avoid duplicate messages when re-running in same process
    for h in list(root.handlers):
        root.removeHandler(h)
    for h in handlers:
        root.addHandler(h)

def get_file_list(path, recursive):
    """
    Return a list of .docx files for the given path.
    """
    file_list = []
    if os.path.isfile(path):
        base = os.path.basename(path)
        if path.lower().endswith(".docx") and (not base.lower().startswith("~$")):
            file_list.append(path)
    elif os.path.isdir(path):
        for root, _, files in os.walk(path):
            for file in files:
                if file.lower().endswith(".docx") and (not file.lower().startswith("~$")):
                    file_list.append(os.path.join(root, file))
            if not recursive:
                break
    return file_list

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
            # Normalize paragraph text before searching
            para_text = unicodedata.normalize("NFC", para.text)
            if regex.search(para_text):
                matched = True

                # Exit early if caller only needs existence or quiet behavior
                if args.quiet or args.files_without_matches or (args.files_with_matches and not args.count):
                    return matches, matched  # exit ASAP if quiet mode is enabled or all we care about is if there was a match

                # Build the formatted line using the helper that defers color/hyperlink
                formatted_line = format_matched_paragraph(file_path, para_text, regex, args, i)
                matches.append(formatted_line)

                logging.debug(f"Match found in {file_path} at paragraph {i+1}")
    except Exception as e:
        if not (args.quiet or args.no_messages):
            logging.error(f"Error reading {file_path}: {e}")

    return matches, matched

def format_matched_paragraph(file_path, para_text, regex, args, para_index):
    """
    Build a single formatted match line for display.

    - Do wrapping and indentation on plain text (no color or hyperlink codes).
    - After wrapping/indenting is done, apply hyperlinking to the prefix (if requested)
      and apply color/highlighting to the prefix and matched text (if requested).
    - Returns a single formatted string ready to print.
    """

    # --------------------------------------------------
    # Perform all indent actions base don the text withOUT color/hyperlink codes (codes effect sizing calculations)

    # Prefix (throwaway for prefix width calculations)
    # plain_prefix = f"{file_path} [Paragraph {para_index+1}]: "
    # if args.initial_tab:
    #     plain_prefix = "\t" + plain_prefix

    # Hanging indent: wrap the plain body using the visible prefix length
    if args.hanging_indent:
        term_width = shutil.get_terminal_size((80, 20)).columns
        # visible_prefix_len = len(plain_prefix)
        margin = 8  # extra margin to avoid edge overflow (use 2 tabs)
        # wrap_width = max(20, term_width - visible_prefix_len - margin) # now that we put a newline after the filename, the prefix width is not longer needed
        wrap_width = max(20, term_width - margin)
        indent = "\t"
        para_text = textwrap.fill(
            para_text,
            width=wrap_width,
            initial_indent=indent,
            subsequent_indent=indent,
            break_on_hyphens=False
        )

    # --------------------------------------------------
    # Now that the sizing is done, apply the color/hyperlink codes

    # Build the display prefix (may include hyperlink)
    if args.hyperlink:
        display_prefix = f"{make_hyperlink(file_path)} [Paragraph {para_index+1}]: "
    else:
        display_prefix = f"{file_path} [Paragraph {para_index+1}]: "
    # Initial tab indent (the real one this time)
    if args.initial_tab:
        display_prefix = "\t" + display_prefix

    # Apply color/highlighting after wrapping/indenting
    if args.color:
        display_prefix = colorize(display_prefix, COLORS['GREEN'])
        display_body = highlight_matches(para_text, regex.pattern, COLORS['RED'], args.ignore_case)
    else:
        display_body = para_text

    # Combine prefix and body using the same layout as the plain version
    if args.hanging_indent:
        formatted = display_prefix + os.linesep + display_body
    else:
        formatted = display_prefix + display_body

    return formatted

def make_hyperlink(path, label=None):
    """
    Wrap `path` in an OSC 8 hyperlink sequence.
    - path: filesystem path to your .docx
    - label: visible text; defaults to path
    """
    # Ensure absolute URI
    # uri = f"file://{os.path.abspath(path)}"
    uri = Path(path).absolute().as_uri()
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
        # if we got here and quiet-mode is on, we found nothing; so, it is a failure/no match
        sys.exit(1)  # exit with status 1 (failure / no match)
    elif args.files_without_matches:
        if args.hyperlink:
            for f in sorted(unmatched_files):
                print(make_hyperlink(f))
        else:
            for f in sorted(unmatched_files):
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

def supports_hyperlink():
    """
    Detect if the current terminal likely supports OSC 8 hyperlinks.
    Heuristic priority (high -> low):
      1. Terminal-specific env markers (DOMTERM, WT_SESSION, KONSOLE_VERSION)
      2. VTE version check (VTE_VERSION >= 5000)
      3. TERM_PROGRAM known terminals
      4. TERM known terminal names
      5. COLORTERM indicating modern terminal (truecolor/24bit or known name)

    Returns True when it is very likely that OSC 8 will render links.
    Returns False when non-interactive or when no reliable indicator exists.
    """
    # --- If stdout is not a TTY, avoid claiming support (non-interactive).
    # This mirrors common patterns: hyperlinks are only useful for interactive terminals.
    if not sys.stdout.isatty():
        return False

    # --- DOMTERM
    # DOMTERM is set by DomTerm, a terminal with HTML/OSC capabilities.
    # Example: DOMTERM=1
    if "DOMTERM" in os.environ:
        return True

    # --- Windows Terminal (WT_SESSION)
    # WT_SESSION is set by Windows Terminal to a GUID-like string when running inside it.
    # Example: WT_SESSION=\\?\pipe\WindowsTerminal…
    if "WT_SESSION" in os.environ:
        return True

    # --- Konsole (KONSOLE_VERSION)
    # KDE Konsole sets KONSOLE_VERSION; its presence implies Konsole which supports OSC 8.
    # Example: KONSOLE_VERSION=245.7
    if "KONSOLE_VERSION" in os.environ:
        return True

    # --- VTE-based terminals (VTE_VERSION)
    # GNOME Terminal, Tilix, Guake, and other VTE-based emulators set VTE_VERSION.
    # Historically, VTE >= 0.50 (reported as 5000) reliably supports OSC 8.
    vte = os.environ.get("VTE_VERSION")
    if vte:
        # Common form is an integer string; try that first.
        try:
            parsed =  int(vte)
        except ValueError:
            # Fallback: keep only digits (rare cases), then parse
            digits = "".join(ch for ch in vte if ch.isdigit())
            if digits:
                try:
                    parsed =  int(digits)
                except ValueError:
                    parsed =  None
            return False
        if parsed is not None and parsed >= 5000:
            return True

    # --- TERM_PROGRAM
    # Many frontends set TERM_PROGRAM to identify themselves:
    #   - "iTerm.app" => iTerm2 on macOS
    #   - "WezTerm" => WezTerm
    #   - "Hyper" => Hyper
    #   - "vscode", "vscode-insiders" => VS Code integrated terminal
    #   - "terminology" => Enlightenment Terminology
    term_program = os.environ.get("TERM_PROGRAM", "").strip()
    if term_program:
        # Normalized matching for common known-positive terminals.
        if term_program in {
            "iTerm.app",
            "WezTerm",
            "Hyper",
            "terminology",
            "vscode",
            "vscode-insiders",
        }:
            return True

    # --- TERM values that identify specific emulators
    # Some terminals don't set TERM_PROGRAM but use distinctive TERM values:
    #   - "xterm-kitty" or "kitty" => kitty terminal
    #   - "alacritty", "alacritty-direct" => Alacritty
    #   - "konsole" => Konsole (some distros)
    term = os.environ.get("TERM", "").strip()
    if term:
        if term in {"xterm-kitty", "kitty", "alacritty", "alacritty-direct", "konsole"}:
            return True

    # --- COLORTERM hints
    # COLORTERM is often set to "truecolor", "24bit", or a terminal name.
    # Many modern terminals set COLORTERM; truecolor/24bit suggests modern feature set.
    colorterm = os.environ.get("COLORTERM", "").lower().strip()
    if colorterm:
        if "truecolor" in colorterm or "24bit" in colorterm:
            return True
        # Some terminals set their name in COLORTERM, e.g., "xfce4-terminal"
        if colorterm == "xfce4-terminal":
            return True


    # Default: do not claim hyperlink support.
    return False

def suggest_terminals_if_no_hyperlink():
    """
    Prints suggestions  if hyperlinks are not supported
    """

    logging.warning("Your terminal does not appear to support OSC 8 hyperlinks.")
    logging.warning("Try Windows Terminal — https://github.com/microsoft/terminal")
    logging.warning("Try iTerm2 (macOS)   — https://iterm2.com")

    return
    
# -----------------------------------------------------------------------
# MAIN
# -----------------------------------------------------------------------
if __name__ == "__main__":
    main()