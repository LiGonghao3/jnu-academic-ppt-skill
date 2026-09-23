# -*- coding: utf-8 -*-
"""调 ChatGPT（OpenAI Images API）生成示意图，存成 PNG 供 deck.json 引用。

只用来画**概念示意图**：技术路线、流程框图、类比插画、封面装饰。
**绝对不能用来画实验结果图、数据图表、性能曲线** —— 那等于伪造实验证据。
真实结果一律用原图或自己跑出来的图，详见 references/figures.md。

配置（任选其一）：
  1. 环境变量  OPENAI_API_KEY=sk-...
  2. 配置文件  <skill根目录>/.openai.json  ->  {"api_key": "sk-...", "model": "gpt-image-2.5-flare"}
     （已在 .gitignore 里，不会被提交）

用法：
  python gen_figure.py "一张解释自注意力如何连接序列中任意两个位置的示意图" \\
      --out 素材/self_attn.png --theme jnu-teal --size 1536x1024

  python gen_figure.py --check          # 只检查接口是否配好
"""
from __future__ import annotations

import argparse
import base64
import json
import os
import re
import sys
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
API_URL = "https://api.openai.com/v1/images/generations"
DEFAULT_MODEL = "gpt-image-2.5-flare"
SIZES = {"1536x1024", "1024x1536", "1024x1024", "auto"}   # 官方推荐尺寸


def size_error(size):
    """按 GPT Image 自定义尺寸规则校验；合法返回 None。
    规则：宽高均为 16 的倍数，长边 ≤3840，宽高比 1:3–3:1，总像素 655,360–8,294,400。"""
    if size in SIZES:
        return None
    m = re.fullmatch(r"(\d+)x(\d+)", size)
    if not m:
        return "应为 auto 或 宽x高，例如 1536x1024"
    w, h = int(m.group(1)), int(m.group(2))
    if w % 16 or h % 16:
        return "宽和高都必须是 16 的倍数"
    if max(w, h) > 3840:
        return "长边不能超过 3840"
    if max(w, h) > 3 * min(w, h):
        return "宽高比须在 1:3 到 3:1 之间"
    if not 655_360 <= w * h <= 8_294_400:
        return "总像素须在 655,360 到 8,294,400 之间"
    return None

# 学术示意图的通用约束。生成模型很爱加装饰、乱写英文标签、上渐变，这里全部掐死。
STYLE_PREAMBLE = (
    "A clean, flat, minimal technical diagram for an academic slide. "
    "Vector-style shapes, thin uniform strokes, generous whitespace, no 3D, "
    "no gradients, no drop shadows, no photographic texture, no decorative icons, "
    "no background scenery. White background. "
    "Do NOT invent any numbers, axes, data points, charts, or plotted curves. "
    "Keep any text labels to a minimum and spell them exactly as given. "
)


def load_config():
    cfg_path = ROOT / ".openai.json"
    cfg = {}
    if cfg_path.exists():
        try:
            cfg = json.loads(cfg_path.read_text(encoding="utf-8"))
        except Exception as e:
            print(f"! 读不了 {cfg_path}：{e}", file=sys.stderr)
    key = os.environ.get("OPENAI_API_KEY") or cfg.get("api_key")
    model = os.environ.get("OPENAI_IMAGE_MODEL") or cfg.get("model") or DEFAULT_MODEL
    base = os.environ.get("OPENAI_BASE_URL") or cfg.get("base_url")  # 兼容代理/中转
    url = (base.rstrip("/") + "/images/generations") if base else API_URL
    return key, model, url


def theme_palette_hint(theme_id):
    """把主题的主色塞进 prompt，让示意图和 PPT 配色是一套。"""
    p = ROOT / "themes" / theme_id / "theme.json"
    if not p.exists():
        return ""
    try:
        pal = json.loads(p.read_text(encoding="utf-8"))["palette"]
    except Exception:
        return ""
    return (f"Use this restricted palette only: primary {pal['primary']}, "
            f"secondary {pal['secondary']}, light fill {pal['surface']}, "
            f"outline {pal['line']}, text {pal['ink']} on white. ")


def generate(prompt, out, theme=None, size="1536x1024", quality="high",
             output_format="png", timeout=180):
    key, model, url = load_config()
    if not key:
        raise SystemExit(
            "没配 API key。二选一：\n"
            "  set OPENAI_API_KEY=sk-...\n"
            f"  或写 {ROOT / '.openai.json'}  ->  {{\"api_key\": \"sk-...\"}}")

    full = STYLE_PREAMBLE + (theme_palette_hint(theme) if theme else "") + prompt
    body = {"model": model, "prompt": full, "n": 1, "size": size}
    if model.startswith("gpt-image-"):
        body["quality"] = quality
        body["output_format"] = output_format

    req = urllib.request.Request(
        url, data=json.dumps(body).encode("utf-8"),
        headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"})

    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            data = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", "ignore")[:600]
        raise SystemExit(f"接口报错 HTTP {e.code}：{detail}")
    except urllib.error.URLError as e:
        raise SystemExit(f"连不上 {url}：{e.reason}（国内直连 OpenAI 通常要代理，"
                         f"可在 .openai.json 里设 base_url 走中转）")

    item = data["data"][0]
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    if item.get("b64_json"):
        out.write_bytes(base64.b64decode(item["b64_json"]))
    elif item.get("url"):
        with urllib.request.urlopen(item["url"], timeout=timeout) as r:
            out.write_bytes(r.read())
    else:
        raise SystemExit(f"返回里既没有 b64_json 也没有 url：{str(data)[:300]}")

    # 记一份出处，交付时要能说清哪张图是 AI 画的
    sidecar = out.with_suffix(out.suffix + ".json")
    sidecar.write_text(json.dumps(
        {"generated_by": model, "prompt": prompt, "theme": theme, "size": size,
         "quality": quality, "output_format": output_format,
         "note": "AI 生成的示意图，非实验结果。引用时必须在图注里注明。"},
        ensure_ascii=False, indent=2), encoding="utf-8")
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("prompt", nargs="?", help="画什么。用中文或英文描述结构，别描述数据")
    ap.add_argument("--out", help="输出 PNG 路径")
    ap.add_argument("--theme", help="按这套主题的配色画，和 PPT 保持一致")
    ap.add_argument("--size", default="1536x1024",
                    help="推荐 " + " | ".join(sorted(SIZES)) + "；也可自定义宽x高")
    ap.add_argument("--quality", default="high", choices=["low", "medium", "high", "auto"])
    ap.add_argument("--format", dest="output_format", default="png",
                    choices=["png", "jpeg", "webp"], help="GPT Image 输出格式")
    ap.add_argument("--check", action="store_true", help="只检查接口配置")
    a = ap.parse_args()

    if a.check:
        key, model, url = load_config()
        print(f"endpoint : {url}")
        print(f"model    : {model}")
        print(f"api_key  : {'已配置（' + key[:7] + '…）' if key else '未配置'}")
        sys.exit(0 if key else 1)

    if not (a.prompt and a.out):
        ap.error("要同时给 prompt 和 --out（或者只用 --check）")
    err = size_error(a.size)
    if err:
        ap.error(f"size 不合法：{err}")

    p = generate(a.prompt, a.out, a.theme, a.size, a.quality, a.output_format)
    print(f"已生成：{p}（出处记录在 {p.name}.json）")
    print("提醒：这是 AI 示意图，deck.json 的 caption 里要注明「示意图（AI 生成）」。")


if __name__ == "__main__":
    main()
