"""
バックアップモード別のファイル複製処理。

モード：
  常に作成           … 毎回タイムスタンプ付きで作成
  日付フォルダにまとめる … 同日のバックアップを 1 フォルダにまとめる
  直前1件のみ保持    … 既存バックアップを削除してから新規作成
  実行時に確認       … main.py 側でダイアログを出した後、このモジュールは「常に作成」と同等
  作成しない         … 何もしない
"""

import glob
import os
import shutil
from datetime import datetime
from pathlib import Path

from logger import Logger


def backup_file(dst_path: str, mode: str, logger: Logger) -> bool:
    """
    戻り値：成功（またはスキップ）= True、失敗 = False
    """
    if mode == '作成しない':
        logger.log("バックアップ：スキップ（設定：作成しない）")
        return True

    src = Path(dst_path)
    if not src.exists():
        logger.log(f"バックアップ失敗：元ファイルが存在しません [{dst_path}]")
        return False

    folder = src.parent
    stem   = src.stem
    suffix = src.suffix
    ts     = datetime.now().strftime('%Y%m%d_%H%M%S')

    if mode in ('常に作成', '実行時に確認'):
        backup_path = folder / f"{stem}_バックアップ_{ts}{suffix}"
        return _do_copy(src, backup_path, logger)

    elif mode == '日付フォルダにまとめる':
        date_str   = datetime.now().strftime('%Y%m%d')
        bk_folder  = folder / f"バックアップ_{date_str}"
        bk_folder.mkdir(exist_ok=True)
        backup_path = bk_folder / f"{stem}_バックアップ_{ts}{suffix}"
        return _do_copy(src, backup_path, logger)

    elif mode == '直前1件のみ保持':
        pattern   = str(folder / f"{stem}_バックアップ_*{suffix}")
        old_files = glob.glob(pattern)
        for f in old_files:
            try:
                os.remove(f)
                logger.log(f"古いバックアップ削除：{f}")
            except OSError as e:
                logger.log(f"古いバックアップ削除失敗：{f} ({e})")
        backup_path = folder / f"{stem}_バックアップ_{ts}{suffix}"
        return _do_copy(src, backup_path, logger)

    logger.log(f"バックアップ：未知のモード [{mode}]、スキップします")
    return True


def _do_copy(src: Path, dst: Path, logger: Logger) -> bool:
    try:
        shutil.copy2(str(src), str(dst))
        logger.log(f"バックアップ作成：{dst}")
        return True
    except Exception as e:
        logger.log(f"バックアップ失敗：{e}")
        return False
