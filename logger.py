from datetime import datetime
from pathlib import Path


class Logger:
    def __init__(self):
        self._lines: list[str] = []

    def log(self, msg: str) -> None:
        ts = datetime.now().strftime('%Y/%m/%d %H:%M:%S')
        self._lines.append(f"{ts} {msg}")

    def flush(self, dst_path: str | None = None) -> None:
        if not self._lines:
            return
        text = '\n'.join(self._lines) + '\n\n'
        if dst_path:
            log_path = Path(dst_path).parent / 'tensya.log'
        else:
            log_path = Path(__file__).parent / 'tensya.log'
        with open(log_path, 'a', encoding='utf-8') as f:
            f.write(text)
        self._lines.clear()
