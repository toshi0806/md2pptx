# -*- coding: utf-8 -*-
"""`python3 -m md2pptx` で CLI を起動するためのエントリポイント．

このモジュールは `-m` 実行時にのみ `__main__` として読み込まれるため，
`if __name__ == "__main__":` のガードは不要（常に main を呼ぶ）．
"""
import sys

from .cli import main

sys.exit(main())
