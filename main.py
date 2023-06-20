import sys
from argparse import ArgumentParser

import gui
import processor


def main():
    if len(sys.argv) <= 1:
        gui.show_gui()
        sys.exit(0)

    args = ArgumentParser()
    args.add_argument("source", help="Filepath of source Invoice")
    args.add_argument("destination", nargs="?",
                      help="Filepath of destination report (default same as source but .xlsx)")
    args.add_argument('-d', '--detailed', action='store_true', help="Enabled detailed expense")
    args = args.parse_args()
    source = args.source
    destination = args.destination if args.destination else ""

    detailed = args.detailed
    processor.process_file(src_xml=source, detailed=detailed, callback=handle_ready, dst_xls=destination)


def handle_ready():
    print("Ready")


if __name__ == "__main__":
    main()
