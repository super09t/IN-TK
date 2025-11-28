#!/usr/bin/env python3

import argparse
import os
import re
import sys
from typing import List

try:
    from paddleocr import PaddleOCR  # type: ignore
except Exception as exc:  # pragma: no cover - import failure handled at runtime
    PaddleOCR = None
    _IMPORT_ERROR = exc
else:
    _IMPORT_ERROR = None

_OCR_CACHE = {}


def _clean_text(value: str) -> str:
    if not value:
        return ''
    return re.sub(r'[^A-Za-z0-9]', '', value)


def _extract_text(result) -> str:
    if not result:
        return ''
    texts: List[str] = []
    for item in result:
        if isinstance(item, (list, tuple)) and len(item) >= 2:
            maybe = item[1]
            if isinstance(maybe, (list, tuple)) and maybe:
                candidate = maybe[0]
                if isinstance(candidate, str):
                    texts.append(candidate)
                    continue
        if isinstance(item, (list, tuple)):
            for sub in item:
                if isinstance(sub, (list, tuple)) and len(sub) >= 2:
                    maybe = sub[1]
                    if isinstance(maybe, (list, tuple)) and maybe:
                        candidate = maybe[0]
                        if isinstance(candidate, str):
                            texts.append(candidate)
    return ''.join(texts)


def _get_ocr(lang: str):
    if PaddleOCR is None:
        raise RuntimeError(f'PaddleOCR is not available: {_IMPORT_ERROR}')
    key = lang or 'en'
    cached = _OCR_CACHE.get(key)
    if cached is None:
        cached = PaddleOCR(lang=key)
        _OCR_CACHE[key] = cached
    return cached


def recognize_image(image_path: str, lang: str) -> str:
    if not image_path or not os.path.exists(image_path):
        return ''
    try:
        ocr = _get_ocr(lang)
    except Exception:
        return ''
    try:
        result = ocr.ocr(image_path, cls=False)
    except Exception:
        return ''
    return _clean_text(_extract_text(result))


def run_server(lang: str) -> int:
    if PaddleOCR is None:
        print('__ERROR__', flush=True)
        return 1
    try:
        _get_ocr(lang)
    except Exception as exc:
        import traceback
        traceback.print_exc()
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
    if text is None:
        text = ''
    sys.stdout.write((text or '') + '\n')
    sys.stdout.flush()
    return 0 if text else 1


def main() -> int:
    parser = argparse.ArgumentParser(description='PaddleOCR helper for captcha recognition')
    parser.add_argument('image', nargs='?', help='Path to captcha image')
    parser.add_argument('--lang', default='en', help='Recognition language, default=en')
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
