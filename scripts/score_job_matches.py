#!/usr/bin/env python3
"""岗位匹配初筛：关键词评分 + 招聘骗局风险检测。

评分结果用于 Triage（初筛分层），不代表真实录取概率。
低分岗位进入观察池，不进入投递列表。

用法：
  python score_job_matches.py --profile profile.yaml --jd jobs.jsonl --output score.csv
"""

from __future__ import annotations

import argparse
import csv
import json
import pathlib
import re
import sys


# ── 招聘骗局风险词库 ──────────────────────────────────────────
RISK_TERMS = [
    "培训贷", "贷款培训", "先交费", "先培训后上岗",
    "包就业", "保offer", "保录取", "内推包过",
    "无经验高薪", "月入过万无经验",
    "刷流水", "刷单", "刷信用",
    "纯销售", "地推", "电话销售",
    "交钱", "押金", "服装费", "材料费",
]

# ── 岗位匹配权重 ──────────────────────────────────────────────
# 总分 100
WEIGHTS = {
    "experience_fit": 40,   # 经历匹配度
    "hard_requirements": 20, # 硬性要求
    "user_interest": 15,    # 用户兴趣
    "practical": 15,        # 现实约束
    "risk_screen": 10,      # 风险筛查
}


def read_file(path: str | None) -> str:
    if not path:
        return ""
    p = pathlib.Path(path)
    return p.read_text(encoding="utf-8") if p.exists() else ""


def tokens(text: str) -> list[str]:
    """提取中英文关键词，过滤单字和过短词。"""
    words = re.findall(r"[A-Za-z0-9+#.]+|[\u4e00-\u9fff]{2,}", text.lower())
    return [w for w in words if len(w.strip()) >= 2]


def simple_yaml_values(text: str) -> list[str]:
    """从 YAML 文本中提取所有值列表。"""
    values: list[str] = []
    for line in text.splitlines():
        stripped = line.strip().strip("-").strip().strip('"').strip("'")
        if not stripped or stripped.endswith(":"):
            continue
        if ":" in stripped:
            stripped = stripped.split(":", 1)[1].strip().strip('"').strip("'")
        if stripped:
            values.append(stripped)
    return values


def keyword_score(source: str, target: str, max_points: int) -> tuple[int, list[str]]:
    """计算关键词匹配得分。"""
    source_terms = set(tokens(source))
    target_terms = set(tokens(target))
    if not target_terms:
        return 0, []
    overlap = sorted(source_terms & target_terms)
    # 取 target 词汇数与 5 之间的较大值作为分母，避免小文本过度得分
    score = min(max_points, round(max_points * len(overlap) / max(5, len(target_terms))))
    return score, overlap[:15]


def risk_score(jd_text: str) -> tuple[int, list[str]]:
    """招聘骗局风险检测。"""
    hits = [term for term in RISK_TERMS if term in jd_text]
    return max(0, 10 - len(hits) * 3), hits


def priority_band(score: int) -> str:
    """根据总分确优先级档位。"""
    if score >= 80:
        return "A_强匹配-优先投递"
    if score >= 70:
        return "B_良好匹配-选择性投递"
    if score >= 60:
        return "C_观察池-补证据后投递"
    return "D_不推荐-暂不投递"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--profile", required=True, help="用户画像 YAML 文件")
    parser.add_argument("--jd", required=True, help="JD 列表 JSONL 文件")
    parser.add_argument("--output", required=True, help="输出 CSV 路径")
    args = parser.parse_args()

    # 读取用户画像
    profile_text = read_file(args.profile)
    profile_values = simple_yaml_values(profile_text)
    if not profile_values:
        print(f"警告：{args.profile} 未读取到有效内容", file=sys.stderr)

    # 读取 JD 列表
    jd_path = pathlib.Path(args.jd)
    if not jd_path.exists():
        print(f"错误：JD 文件不存在：{jd_path}", file=sys.stderr)
        return 2

    rows: list[dict[str, str | int]] = []
    for line in jd_path.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue

        job = json.loads(line)
        job_text = " ".join(
            str(job.get(k, "")) for k in [
                "title", "company", "city", "jd_text",
                "hard_requirements", "nice_to_have"
            ]
        )

        # 各维度评分
        exp_fit, exp_terms = keyword_score(profile_text, job_text, WEIGHTS["experience_fit"])
        hard, hard_terms = keyword_score(
            profile_text,
            str(job.get("hard_requirements") or job.get("jd_text", "")),
            WEIGHTS["hard_requirements"]
        )
        interest, int_terms = keyword_score(
            " ".join(profile_values),
            job_text,
            WEIGHTS["user_interest"]
        )
        # 现实约束评分（含城市等维度）
        practical = min(WEIGHTS["practical"],
                        keyword_score(" ".join(profile_values), job_text, WEIGHTS["practical"])[0])
        if practical < 5:
            practical = 5  # 保底分

        risk, risk_hits = risk_score(job_text)

        total = exp_fit + hard + interest + practical + risk

        rows.append({
            "job_id": job.get("job_id", ""),
            "title": job.get("title", ""),
            "company": job.get("company", ""),
            "platform": job.get("platform", ""),
            "city": job.get("city", ""),
            "match_score": total,
            "experience_fit": exp_fit,
            "hard_requirements": hard,
            "interest_fit": interest,
            "practical_fit": practical,
            "risk_score": risk,
            "priority_band": priority_band(total),
            "matching_keywords": "、".join(exp_terms[:10]),
            "risk_flags": "、".join(risk_hits),
            "next_action": (
                "用户阅读JD并手动确认是否进入简历定制"
                if total >= 60
                else "暂不建议投递；补证据或放入观察池"
            ),
        })

    # 写入 CSV
    output_path = pathlib.Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    fieldnames = [
        "job_id", "title", "company", "platform", "city",
        "match_score", "experience_fit", "hard_requirements",
        "interest_fit", "practical_fit", "risk_score",
        "priority_band", "matching_keywords", "risk_flags", "next_action",
    ]
    with output_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    # 输出统计
    bands: dict[str, int] = {}
    for r in rows:
        band = r["priority_band"]
        bands[band] = bands.get(band, 0) + 1

    print(f"评分完成：{len(rows)} 个岗位")
    for band, count in sorted(bands.items()):
        print(f"  {band}: {count} 个")
    print(f"输出文件：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())