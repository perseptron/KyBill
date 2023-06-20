import sys
from argparse import ArgumentParser

import gui
from processor import process_file


def main():
    if len(sys.argv) <= 1:
        gui.show_gui()
        exit(0)

    args = ArgumentParser()
    args.add_argument("source", help="Filepath of source Invoice")
    args.add_argument("destination", nargs="?", help="Filepath of destination report (default same as source but .xlsx)")
    args.add_argument('-d', '--detailed', action='store_true', help="Enabled detailed expense")
    args = args.parse_args()
    source = args.source
    destination = args.destination
    detailed = args.detailed
    process_file(source, destination, detailed)


if __name__ == "__main__":
    main()
