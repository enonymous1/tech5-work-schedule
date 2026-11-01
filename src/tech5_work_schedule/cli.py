"""Command-line interface for tech5-work-schedule."""

from __future__ import annotations

import argparse
import sys
from collections import Counter

from . import __version__
from .config import load_config
from .errors import ConfigError
from .io import generate
from .verify import verify_files


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="tech5-work-schedule",
        description="Generate TECH-5 work-schedule workbooks from JSON configs.",
    )
    parser.add_argument(
        "--version", action="version", version=f"%(prog)s {__version__}"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    gen = sub.add_parser("generate", help="Generate a TECH-5 workbook from a config.")
    gen.add_argument("--config", required=True, help="Path to a JSON config file.")
    gen.add_argument(
        "--output",
        help="Output .xlsx path. Defaults to output/TECH5_<project>_<timestamp>.xlsx.",
    )

    val = sub.add_parser(
        "validate", help="Validate a config without generating anything."
    )
    val.add_argument("--config", required=True, help="Path to a JSON config file.")

    ver = sub.add_parser(
        "verify",
        help="Check a generated workbook's content and key styling against its config.",
    )
    ver.add_argument("--config", required=True, help="Path to a JSON config file.")
    ver.add_argument(
        "--workbook", required=True, help="Path to a generated .xlsx file."
    )
    return parser


def _cmd_generate(args: argparse.Namespace) -> int:
    """Generate a workbook, reporting config and I/O errors cleanly.

    Config problems and filesystem problems (permission denied, output
    path is a directory, disk full) are user errors, not bugs, so they
    get a one-line message and exit code 2 instead of a traceback.
    """
    try:
        output = generate(args.config, args.output)
    except ConfigError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    except OSError as exc:
        print(f"error: cannot write workbook: {exc}", file=sys.stderr)
        return 2
    print(f"Wrote: {output}")
    return 0


def _cmd_validate(args: argparse.Namespace) -> int:
    try:
        config = load_config(args.config)
    except ConfigError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    counts = Counter(item.type for item in config.items)
    print(f"Project: {config.meta.project_name}")
    print(f"Duration: {config.meta.total_duration} {config.meta.time_unit_label}")
    if config.meta.rfp_reference:
        print(f"RFP Reference: {config.meta.rfp_reference}")
    print(f"Items: {len(config.items)}")
    for item_type in sorted(counts):
        print(f"  {item_type}: {counts[item_type]}")
    print("Config is valid.")
    return 0


def _cmd_verify(args: argparse.Namespace) -> int:
    mismatches = verify_files(args.workbook, args.config)
    if not mismatches:
        print(f"OK: {args.workbook} matches {args.config}")
        return 0
    for mismatch in mismatches:
        print(f"mismatch: {mismatch}", file=sys.stderr)
    return 1


def main(argv: list[str] | None = None) -> int:
    """CLI entry point. Returns an exit code; never raises for bad input.

    Exit codes: 0 success, 1 verify mismatch, 2 usage/config/I-O error.
    Note argparse itself exits (via SystemExit) on usage errors and on
    ``--version`` — callers embedding ``main()`` as a library function
    must expect SystemExit, not just an int return.
    """
    args = _build_parser().parse_args(argv)
    if args.command == "generate":
        return _cmd_generate(args)
    if args.command == "validate":
        return _cmd_validate(args)
    if args.command == "verify":
        return _cmd_verify(args)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
