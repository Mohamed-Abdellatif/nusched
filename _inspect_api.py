"""Temporary script to inspect the API response structure."""
import requests
import json

API_URL = "https://register.nu.edu.eg/PowerCampusSelfService/Schedule/Student"
HEADERS = {
    "accept": "application/json",
    "accept-language": "en-US,en-GB;q=0.9,en-ZA;q=0.8,en;q=0.7,ar;q=0.6",
    "cache-control": "max-age=0",
    "content-type": "application/json",
    "priority": "u=1, i",
    "sec-ch-ua": "\"Not(A:Brand\";v=\"8\", \"Chromium\";v=\"144\", \"Google Chrome\";v=\"144\"",
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": "\"Windows\"",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "cookie": (
        "ASP.NET_SessionId=islkk4aslqpogm0ierrjgphz; "
        ".AspNet.Cookies=rpP4AHwFLLNWwmGW6IFd0jYONSmI0j1btMSKy5Yz7Vm6iEtTxbt0WLR2hlkpvKN6xR7vEYOlJTNYmEvlsRbBnuX"
        "_ncGMnd0Rtv8-SwzqvjqcJAMlGUHl7ccyElpZe4pe6RwO-gcEfz_d7VUDZyOUCFOyEVLwATca3PrSYEN9uXKhBOwHKp9cXvWSbVOYywLs"
        "-nDh1Oyk0XRl70nO0xrhg2sxobRtxhadLvMOk6vL5L80ndaiqs8GSD0AGBWWUNdK4-H4Ht_SmEloP-47bLYH_6ofWxQJNeJO5dc-O-d5x"
        "_deyYg5zoHJkoQBeE3uNnGCzdPLg_q8IFn8vqUaQ0-m2EmJQrCOpvJJSKqv9rkR1I-S3esZzUBAZKvDXJi_3LpN4nR2bZQvrZthK97sRMv4"
        "--7VDgSmfgJEmvG1RAlCvyQMTf9r5VEqSK8qqYonUgitH2V66-tse5VL55HjfNvTfT7zAepkboHqmRq9pk7Jk-TQc4XfMHR4Oy4EsAelkX"
        "2XiIVPDxAQ7gNuEb6_kzvxLrVZVRJUIeVlVvIaR-EX9ylK7dHorv-iwZhMZ9o0MZspVg_NwI_C0y1OWbjCOuyGsA"
    ),
    "Referer": "https://register.nu.edu.eg/PowerCampusSelfService/Registration/Schedule",
}
BODY = {
    "personId": 15734,
    "yearTermSession": {"year": "2026", "term": "SPRG", "session": ""},
}

resp = requests.post(API_URL, headers=HEADERS, json=BODY, timeout=30)
print("Status:", resp.status_code)
print("Content-Type:", resp.headers.get("content-type"))
print()

data = resp.json()
print("Type:", type(data).__name__)

if isinstance(data, dict):
    print("Top-level keys:", list(data.keys()))
    for k, v in data.items():
        print(f"  [{k}]: type={type(v).__name__}", end="")
        if isinstance(v, list):
            print(f", len={len(v)}")
            if v:
                first = v[0]
                print(f"    first item type: {type(first).__name__}")
                if isinstance(first, dict):
                    print(f"    first item keys: {list(first.keys())}")
                    # Show nested structure
                    for fk, fv in first.items():
                        vtype = type(fv).__name__
                        if isinstance(fv, list) and fv:
                            print(f"      [{fk}]: list, len={len(fv)}, first={type(fv[0]).__name__}")
                            if isinstance(fv[0], dict):
                                print(f"        keys: {list(fv[0].keys())}")
                        elif isinstance(fv, dict):
                            print(f"      [{fk}]: dict, keys={list(fv.keys())}")
                        else:
                            print(f"      [{fk}]: {vtype} = {repr(fv)[:120]}")
        elif isinstance(v, dict):
            print(f", keys={list(v.keys())}")
        else:
            print(f", value={repr(v)[:200]}")
elif isinstance(data, list):
    print("List length:", len(data))
    if data:
        first = data[0]
        print("First item type:", type(first).__name__)
        if isinstance(first, dict):
            print("First item keys:", list(first.keys()))
            for fk, fv in first.items():
                vtype = type(fv).__name__
                if isinstance(fv, list) and fv:
                    print(f"  [{fk}]: list, len={len(fv)}, first={type(fv[0]).__name__}")
                    if isinstance(fv[0], dict):
                        print(f"    keys: {list(fv[0].keys())}")
                elif isinstance(fv, dict):
                    print(f"  [{fk}]: dict, keys={list(fv.keys())}")
                else:
                    print(f"  [{fk}]: {vtype} = {repr(fv)[:120]}")

print()
print("=== RAW JSON (first 5000 chars) ===")
raw = json.dumps(data, indent=2, ensure_ascii=False)
print(raw[:5000])
if len(raw) > 5000:
    print(f"\n... (truncated, total {len(raw)} chars)")
