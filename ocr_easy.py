#!/usr/bin/env python3

import argparse
import os
import re
import sys
from typing import List

try:
    from easyocr import Reader  # type: ignore
except Exception as exc:  # pragma: no cover
    Reader = None
    _IMPORT_ERROR = exc
else:
    _IMPORT_ERROR = None

_READERS = {}


def _clean_text(value: str) -> str:
    if not value:
        return ''
    return re.sub(r'[^A-Za-z0-9]', '', value)


def _get_reader(lang: str):
    if Reader is None:
        raise RuntimeError(f'EasyOCR is not available: {_IMPORT_ERROR}')
    key = tuple(sorted(lang.split('+'))) if lang else ('en',)
    if key not in _READERS:
        languages = [part or 'en' for part in key]
        _READERS[key] = Reader(languages, gpu=False, verbose=False)
    return _READERS[key]


def recognize_image(image_path: str, lang: str) -> str:
    if not image_path or not os.path.exists(image_path):
        return ''
    try:
        reader = _get_reader(lang or 'en')
    except Exception:
        return ''
    try:
        results = reader.readtext(image_path, detail=0, paragraph=False)
    except Exception:
        return ''
    combined = ''.join(result for result in results if isinstance(result, str))
    return _clean_text(combined)


def run_server(lang: str) -> int:
    if Reader is None:
        print('__ERROR__', flush=True)
        return 1
    try:
        _get_reader(lang or 'en')
    except Exception:
        print('__ERROR__', flush=True)
        return 1
    print('__READY__', flush=True)
    for line in sys.stdin:
        path = line.strip()
        if path == '__EXIT__':
            break
        if not path:
            print('', flush=True)
            continue
        text = recognize_image(path, lang)
        print(text, flush=True)
    return 0


def run_once(image_path: str, lang: str) -> int:
    text = recognize_image(image_path, lang)
    sys.stdout.write((text or '') + '\n')
    sys.stdout.flush()
    return 0 if text else 1


def main() -> int:
    parser = argparse.ArgumentParser(description='EasyOCR helper for captcha recognition')
    parser.add_argument('image', nargs='?', help='Path to captcha image')
    parser.add_argument('--lang', default='en', help='Language code(s), default=en')
    parser.add_argument('--server', action='store_true', help='Run in stdin/stdout server mode')
    args = parser.parse_args()

    if args.server:
        return run_server(args.lang)

    if not args.image:
        parser.print_usage(sys.stderr)
        return 2

    return run_once(args.image, args.lang)


if __name__ == '__main__':
    sys.exit(main())
