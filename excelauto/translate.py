"""Korean words in item names -> German/English.

Base newPrice item names must not contain Korean. Order of work for one name:
1. option labels like "크기: " / "맛: " are removed (same idea as "Geschmack: " in 망고변환)
2. each remaining Korean phrase is looked up in the glossary (ExcelAuto_단어장.xlsx, editable in Excel)
3. phrases not in the glossary are translated with DeepL (if an API key is set) and saved to the glossary
Only the Korean option/name phrases are sent to DeepL, never recipient data.
"""
import json
import os
import re
import urllib.error
import urllib.request
from typing import Dict, Iterable, List

import openpyxl

HANGUL = '가-힣ㄱ-ㆎ'
HANGUL_RE = re.compile(f'[{HANGUL}]')
LABEL_RE = re.compile(rf'[{HANGUL}][{HANGUL} ]*:\s*')                  # 크기: / 색상 선택:
PHRASE_RE = re.compile(rf'[{HANGUL}]+(?:[ ]+[{HANGUL}]+)*')            # 리스트레토 / 무향 캡슐
GLOSSARY_NAME = 'ExcelAuto_단어장.xlsx'
CRED_TARGET = 'ExcelAuto/DeepL'


def has_hangul(s: str) -> bool:
    return bool(HANGUL_RE.search(s or ''))


def strip_labels(s: str) -> str:
    return LABEL_RE.sub('', s)


def phrases(s: str) -> List[str]:
    return PHRASE_RE.findall(s or '')


class Glossary:
    """Two columns: 한글 | 변환. Row 1 is the header. Later rows win."""

    def __init__(self, path: str):
        self.path = path
        self.words: Dict[str, str] = {}
        if os.path.exists(path):
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            for r in wb.worksheets[0].iter_rows(min_row=2, max_col=2, values_only=True):
                k = str(r[0] or '').strip()
                v = '' if len(r) < 2 or r[1] is None else str(r[1]).strip()
                if k and v:
                    self.words[k] = v
            wb.close()

    def get(self, k):
        return self.words.get(k)

    def add(self, new: Dict[str, str]):
        new = {k.strip(): v.strip() for k, v in new.items() if k.strip() and v.strip()}
        if not new:
            return
        self.words.update(new)
        self.save()

    def save(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '단어장'
        ws.append(['한글', '변환 (독일어/영어)', ])
        for k in sorted(self.words):
            ws.append([k, self.words[k]])
        ws.column_dimensions['A'].width = 30
        ws.column_dimensions['B'].width = 40
        tmp = self.path + '.tmp.xlsx'
        wb.save(tmp)
        os.replace(tmp, self.path)


def apply(s: str, glossary: Glossary) -> str:
    """Labels removed, known phrases replaced. Unknown Korean stays (caller checks has_hangul)."""
    s = strip_labels(s)
    return PHRASE_RE.sub(lambda m: glossary.get(m.group(0)) or m.group(0), s)


# ---------------------------------------------------------------- DeepL
def deepl_url(key: str) -> str:
    free = key.strip().endswith(':fx')
    return 'https://api-free.deepl.com/v2/translate' if free else 'https://api.deepl.com/v2/translate'


def deepl_translate(texts: List[str], key: str, target: str = 'DE', timeout=20) -> List[str]:
    body = json.dumps({'text': texts, 'source_lang': 'KO', 'target_lang': target}).encode('utf-8')
    req = urllib.request.Request(deepl_url(key), data=body, method='POST',
                                 headers={'Authorization': 'DeepL-Auth-Key ' + key.strip(),
                                          'Content-Type': 'application/json'})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            data = json.loads(r.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        msg = {403: 'API 키가 맞지 않습니다', 456: '이번 달 무료 사용량을 다 썼습니다', 429: '요청이 너무 많습니다. 잠시 후 다시'}
        raise RuntimeError(f'DeepL 오류 {e.code}: {msg.get(e.code, e.reason)}') from None
    except (urllib.error.URLError, TimeoutError) as e:
        raise RuntimeError(f'DeepL에 연결하지 못했습니다: {e}') from None
    out = [t['text'] for t in data.get('translations', [])]
    if len(out) != len(texts):
        raise RuntimeError('DeepL 응답 형식이 예상과 다릅니다.')
    return out


def learn(names: Iterable[str], glossary: Glossary, translator) -> Dict[str, str]:
    """Translates the Korean phrases of `names` that the glossary lacks; saves and returns them.

    translator: fn(list of str) -> list of str, or None (then nothing is translated).
    """
    missing = sorted({p for n in names for p in phrases(strip_labels(n)) if not glossary.get(p)})
    if not missing or translator is None:
        return {}
    got = {}
    for i in range(0, len(missing), 50):
        part = missing[i:i + 50]
        for k, v in zip(part, translator(part)):
            v = v.strip()
            if v and not has_hangul(v):
                got[k] = v
    glossary.add(got)
    return got


# ---------------------------------------------------------------- API key
# Saved next to the program as ExcelAuto_deepl.key, encrypted with Windows DPAPI (CryptProtectData):
# only the same Windows user on the same PC can read it. (v0.4.1 used the Windows 자격 증명 관리자 through
# pywin32, which did not keep the key in the built program.)
KEY_NAME = 'ExcelAuto_deepl.key'
_memory_key = {}


def _dpapi(data: bytes, protect: bool) -> bytes:
    import ctypes
    from ctypes import wintypes

    class BLOB(ctypes.Structure):
        _fields_ = [('cbData', wintypes.DWORD), ('pbData', ctypes.POINTER(ctypes.c_char))]
    buf = ctypes.create_string_buffer(data, len(data))
    src = BLOB(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_char)))
    out = BLOB()
    crypt32, kernel32 = ctypes.windll.crypt32, ctypes.windll.kernel32
    fn = crypt32.CryptProtectData if protect else crypt32.CryptUnprotectData
    if not fn(ctypes.byref(src), None, None, None, None, 0x1, ctypes.byref(out)):    # 0x1: no UI
        raise ctypes.WinError()
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(out.pbData, ctypes.c_void_p))


def load_key(folder: str) -> str:
    """'' when no key is saved or it cannot be read."""
    path = os.path.join(folder, KEY_NAME)
    if os.name != 'nt':
        return _memory_key.get(path, '')
    try:
        with open(path, 'rb') as f:
            return _dpapi(f.read(), False).decode('utf-8')
    except Exception:          # no file, or saved by another Windows user
        return ''


def save_key(folder: str, key: str):
    """Saves (or with an empty key deletes) the key, then reads it back; raises when that fails."""
    key = key.strip()
    path = os.path.join(folder, KEY_NAME)
    if os.name != 'nt':
        _memory_key[path] = key
    elif not key:
        if os.path.exists(path):
            os.remove(path)
    else:
        data = _dpapi(key.encode('utf-8'), True)
        tmp = path + '.tmp'
        with open(tmp, 'wb') as f:
            f.write(data)
        os.replace(tmp, path)
    if load_key(folder) != key:
        raise RuntimeError(f'키를 저장한 뒤 다시 읽지 못했습니다: {path}')


def mask(key: str) -> str:
    return ('•' * 8 + key[-4:]) if len(key) > 4 else ('•' * len(key))


def test_key(key: str, target: str = 'DE') -> str:
    """One short request; returns the translation of 시험 or raises RuntimeError."""
    return deepl_translate(['시험'], key, target, timeout=15)[0]
