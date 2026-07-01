# -*- coding: utf-8 -*-
"""`python3 -m md2pptx` で CLI を起動するためのエントリポイント．"""
import sys

from .cli import main

if __name__ == "__main__":
    sys.exit(main())
