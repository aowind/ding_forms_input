"""Cookie 解析模块 - 将原始 Cookie 字符串解析为 Playwright cookie 列表"""


def parse_cookie_string(raw: str, domain: str = ".dingtalk.com") -> list[dict]:
    """解析原始 Cookie 字符串为 cookie 对象列表。

    Args:
        raw: 原始 Cookie 字符串，格式 ``name1=value1; name2=value2; ...``
        domain: 默认域名，默认 ``.dingtalk.com``

    Returns:
        Playwright add_cookies 所需的 cookie 字典列表
    """
    cookies = []
    for part in raw.split(";"):
        part = part.strip()
        if not part or "=" not in part:
            continue
        name, _, value = part.partition("=")
        name = name.strip()
        value = value.strip()
        cookies.append({
            "name": name,
            "value": value,
            "domain": domain,
            "path": "/",
        })
    return cookies
