"""PyInstaller entry point for the command-line build. See gui_entry.py."""

import multiprocessing
import sys

if __name__ == "__main__":
    multiprocessing.freeze_support()

    from excel_mass_replacer.cli import main

    sys.exit(main())
