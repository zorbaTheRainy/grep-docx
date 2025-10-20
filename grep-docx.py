#!/usr/bin/env python3

import os
import re
import argparse
import logging
import shutil
import textwrap
from docx import Document

VERSION = "1.0"

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
                if not (config["quiet"] or config["count"] or config["list_files"]):
                    prefix = f"{file_path} [Paragraph {i+1}]: "
                    if config.get("color"):
                        prefix_colored = colorize(prefix, "36")  # Cyan for prefix
                        para_text = highlight_matches(
                            para.text, regex.pattern, "31", config["ignore_case"]
                        )  # Red for match
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
                    else:
                        formatted_line = prefix_colored + para_text

                    matches.append(formatted_line)
                    logging.debug(f"Match found in {file_path} at paragraph {i+1}")
    except Exception as e:
        logging.error(f"Error reading {file_path}: {e}")
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
    matched_files = set()
    path = config["path"]
    if os.path.isfile(path):
        if path.endswith(".docx"):
            logging.debug(f"Searching file: {path}")
            file_matches, matched = search_file(path, regex, config)
            match_count += len(file_matches)
            if matched:
                matched_files.add(path)
            matches.extend(file_matches)
    elif os.path.isdir(path):
        for root, _, files in os.walk(path):
            for file in files:
                if file.endswith(".docx"):
                    full_path = os.path.join(root, file)
                    logging.debug(f"Searching file: {full_path}")
                    file_matches, matched = search_file(full_path, regex, config)
                    match_count += len(file_matches)
                    if matched:
                        matched_files.add(full_path)
                    matches.extend(file_matches)
            if not config["recursive"]:
                break
    else:
        logging.error(f"Invalid path: {path}")
        return
    if config["count"]:
        print(match_count)
    elif config["list_files"]:
        for f in matched_files:
            print(f)
    elif not config["quiet"]:
        if config["output"]:
            with open(config["output"], 'w', encoding='utf-8') as f:
                for line in matches:
                    f.write(line + '\n')
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
    parser.add_argument("-o", "--output", help="Write results to a file instead of stdout")
    parser.add_argument("-c", "--count", action="store_true", help="Only print a count of matching lines")
    parser.add_argument("-l", "--files-with-matches", action="store_true", help="Only print names of files with matches")
    parser.add_argument("-q", "--quiet", "--silent", action="store_true", help="Suppress all normal output")
    parser.add_argument("-T", "--initial-tab", action="store_true", help="Line output starts with a tab character")
    parser.add_argument("-I", "--hanging_indent", action="store_true", help="Line output after the 1st line starts with a tab character")
    parser.add_argument("-C", "--color", action="store_true", help="Color the prefix and highlight matches")
    parser.add_argument("-V", "--version", action="version", version=f"%(prog)s {VERSION}")
    parser.add_argument("--debug", action="store_true", help="Enable debug logging")
    args = parser.parse_args()
    setup_logging(args.debug)
    config = {
        "pattern": args.pattern,
        "path": args.path,
        "ignore_case": args.ignore_case,
        "recursive": args.recursive,
        "output": args.output,
        "count": args.count,
        "list_files": args.files_with_matches,
        "quiet": args.quiet,
        "hanging_indent": args.hanging_indent,
        "initial_tab": args.initial_tab,
        "color": args.color
    }
    grep_docx(config)

if __name__ == "__main__":
    main()