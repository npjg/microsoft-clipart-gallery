#!/usr/local/bin/python3

import argparse
import os 
import json

def main():
    # DEFINE THE COMMAND-LINE ARGUMENT PARSERS.
    argument_parser = argparse.ArgumentParser(
        formatter_class = argparse.RawTextHelpFormatter,
        description = "Extracts information from Microsoft Office clipart catalogs")
    subparsers = argument_parser.add_subparsers(
        required = True,
        title = 'Easter eggs',
        description = 'Choose one of the following supported Microsoft Office clipart catalogs. You must have access to the required original file(s) from the original Microsoft products.')

    # Windows 3.1 Credits.
    office97_clipart_argument_parser = subparsers.add_parser(
        name = 'office97',
        description = "Extracts information from the Office 97 clipart catalog (Clip Art 3.0).")
    input_argument_help = "The filepath to the desired Office 97 CAG file(s)."
    office97_clipart_argument_parser.add_argument('input', nargs = '+', help = input_argument_help)
    export_argument_help = "Specify the directory location for exporting information."
    office97_clipart_argument_parser.add_argument('export', help = export_argument_help)
    def parse_office97_clipart_arguments(command_line_args):
        from MicrosoftClipartCatalog.Office97 import MicrosoftClipArt30Catalog
        for input_filepath in command_line_args.input:
            clipart_catalog = MicrosoftClipArt30Catalog(input_filepath)
            clipart_catalog.export(command_line_args.export)
    office97_clipart_argument_parser.set_defaults(func = parse_office97_clipart_arguments)

    # EXTRACT THE ASSETS.
    command_line_args = argument_parser.parse_args()
    command_line_args.func(command_line_args)

if __name__ == "__main__":
    main()