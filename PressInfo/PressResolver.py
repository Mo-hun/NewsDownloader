# PressInfo/PressResolver.py

from urllib.parse import urlparse
from .PressRegistry import PRESS_MAP


def normalize_domain(url):

    if not url:
        return ""

    domain = urlparse(url).netloc.lower()

    if domain.startswith("www."):
        domain = domain[4:]

    return domain


def extract_press_info(originallink):

    domain = normalize_domain(originallink)

    if not domain:
        return "", "", ""

    # 1️⃣ 정확 매핑
    if domain in PRESS_MAP:
        press_name, category = PRESS_MAP[domain]
        return press_name, category, domain

    # 2️⃣ 서브도메인 처리
    parts = domain.split(".")
    if len(parts) >= 2:
        base_domain = ".".join(parts[-2:])
        if base_domain in PRESS_MAP:
            press_name, category = PRESS_MAP[base_domain]
            return press_name, category, base_domain

    # 3️⃣ fallback
    press_name = parts[0].upper()
    category = "기타"

    return press_name, category, domain