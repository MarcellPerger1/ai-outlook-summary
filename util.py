import datetime
import functools
import pickle
import threading
import time
from datetime import timedelta
from typing import ParamSpec, TypeVar, Hashable, Callable

import pdfkit
import requests

# I can't believe ParamSpec was added as early as 3.10 - it feels like such a 3.13 thing
P = ParamSpec('P')
R = TypeVar('R')
H = TypeVar('H', bound=Hashable)


def _default_cache_key(*args, **kwargs):
    return args, frozenset(kwargs)


def diskcache(filename: str = None, key_fn: Callable[P, H] = _default_cache_key,
              lifetime: datetime.timedelta | float | None = datetime.timedelta(minutes=10)):
    # NOTE: not thread-safe in the slightest! Also, the performance is A LOT
    # worse than functools.cache but this one should be used for very expensive
    # operations (e.g. requesting data from the web)
    lifetime_sec = (
        lifetime.total_seconds() if isinstance(lifetime, datetime.timedelta)
        else datetime.timedelta(days=1e15) if lifetime is None else lifetime)

    def decor(fn: Callable[P, R]) -> Callable[P, R]:
        try:
            if filename is not None:
                with open(filename, 'rb') as fr:
                    cache: dict[H, tuple[float, R]] = pickle.load(fr)
            else:
                cache = {}
        except FileNotFoundError:  # create cache file after first entry added
            cache = {}
        cache_lock = threading.RLock()

        @functools.wraps(fn)
        def new_fn(*args, **kwargs):
            with cache_lock:
                key = key_fn(*args, **kwargs)
                try:  # Could also use contextlib.suppress here but this is clearer
                    birth, value = cache[key]
                    if birth + lifetime_sec > time.time():
                        print(f'(Cache hit for {filename})')
                        return value
                    print(f'(Cache expired for {filename})')
                    del cache[key]  # don't leak even if error below
                except KeyError:
                    print(f'(Cache miss for {filename})')
            value = fn(*args, **kwargs)
            birth = time.time()
            with cache_lock:
                cache[key] = birth, value
                if filename:
                    with open(filename, 'wb') as fw:
                        # Pycharm still doesn't understand Protocol after 6 years!
                        # noinspection PyTypeChecker
                        pickle.dump(cache, fw)
            return value
        return new_fn
    return decor


def fetch(url, conf: dict[str, ...]) -> requests.Response:
    return requests.request(
        conf['method'], url, params=conf.get('params'),
        headers=conf.get('headers'), data=conf.get('body'))


def writefile(fnm: str, data: str):
    with open(fnm, 'w', encoding='utf8') as f:
        f.write(data)


@diskcache('.app_cache/topdf_cache.pkl', lifetime=timedelta(days=30))
def html_to_pdf(html: str) -> bytes:
    return pdfkit.from_string(html, configuration=pdfkit.configuration(
        wkhtmltopdf='C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe'))
