import argparse
import os
import re  # noqa - For `exec`.
import sys
from pathlib import Path  # noqa - For `exec`.


def main():
    parser = argparse.ArgumentParser(prog="prs", description="Add some Python to your shell life.")
    parser.add_argument("script", metavar="CODE", type=str, help="The code to run.")

    args = parser.parse_args()

    if os.isatty(0):
        # Nothing was piped in, maybe the user just wants us to run with no input.
        i = []
    else:
        # Read stdin.
        i = sys.stdin.readlines()

        # Strip final newline.
        i = [x[:-1] for x in i]

    scope = globals().copy()
    scope["i"] = i

    exec(args.script, scope)

    if "o" not in scope:
        print("prs expects its output in a variable called `o`.", file=sys.stderr)
        sys.exit(1)

    result = scope["o"]

    if isinstance(result, list) or isinstance(result, tuple):
        for l in result:
            print(l)
    else:
        print(str(result))


if __name__ == "__main__":
    main()
