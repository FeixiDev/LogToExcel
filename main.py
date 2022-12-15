# -*- coding:utf-8 -*-
import argparse
import sys
import contrl


class InputParser(object):

    def __init__(self):
        self.parser = argparse.ArgumentParser(description="Json to Excel")
        self.setup_parser()

    def setup_parser(self):

        self.parser.add_argument("-p", help="Path of log file", dest="path", type=str)
        self.parser.add_argument("-s", help="Save path of excel", dest='save', type=str)
        self.parser.set_defaults(func=self.main_usage)

    def main_usage(self, args):
        if args.path:
            contrl.fetch_file('.json', files_path=args.path)
        elif args.save:
            contrl.fetch_file('.json', excel_path=args.save)
        elif args.save and args.path:
            contrl.fetch_file('.json', files_path=args.path, excel_path=args.save)
        else:
            contrl.fetch_file('.json')

    def parse(self):
        args = self.parser.parse_args()
        args.func(args)


def main():
    try:
        run_program = InputParser()
        run_program.parse()
    except KeyboardInterrupt:
        sys.stderr.write("\nClient exiting (received SIGINT)\n")


if __name__ == '__main__':
    main()
