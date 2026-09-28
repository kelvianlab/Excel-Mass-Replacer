"""PyInstaller entry point for the desktop app.

A frozen app re-launches its own executable to create worker processes.
freeze_support() must run before anything else, or every worker would open a
second window instead of doing its share of the work.
"""

import multiprocessing


def main() -> None:
    from excel_mass_replacer.gui import App

    App().mainloop()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()
