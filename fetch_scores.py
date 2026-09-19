import csv
import json
import os
import sys
import time
from pathlib import Path

import requests


# ============================================================
# 要抓取的 5 个接口
# ============================================================

ROAD_SHOW_IDS = [
    150,
    151,
    152,
    135,
    136,
]

BASE_URL = (
    "https://itcjspapif.woczx.com/"
    "contest-lives/road-show/{road_show_id}?sort_type=2"
)


# ============================================================
# 模拟浏览器请求头
# ============================================================

HEADERS = {
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "zh-CN,zh;q=0.9",
    "Content-Type": "application/json",
    "Origin": "https://jspitcm.woczx.com",
    "Referer": "https://jspitcm.woczx.com/",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-site",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/152.0.0.0 Safari/537.36"
    ),
    "X-Client-Type": "mobile",
    "sec-ch-ua": (
        '"Chromium";v="152", '
        '"Not?A_Brand";v="24", '
        '"Google Chrome";v="152"'
    ),
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
}


# ============================================================
# 如果以后网站要求 Cookie / Authorization
# 可以放进 GitHub Secrets，不需要写进代码
# ============================================================

cookie = os.getenv("WOCZX_COOKIE", "").strip()
authorization = os.getenv("WOCZX_AUTHORIZATION", "").strip()

if cookie:
    HEADERS["Cookie"] = cookie

if authorization:
    HEADERS["Authorization"] = authorization


def fetch_one(road_show_id):
    url = BASE_URL.format(road_show_id=road_show_id)

    print(f"\n正在读取：{url}")

    response = requests.get(
        url,
        headers=HEADERS,
        timeout=30,
    )

    print(f"HTTP 状态码：{response.status_code}")

    if response.status_code == 401:
        raise RuntimeError(
            f"接口 {road_show_id} 返回 401 Unauthorized。\n"
            "网站可能增加了身份验证，请检查 Cookie 或 Authorization。"
        )

    response.raise_for_status()

    try:
        return response.json()

    except Exception:
        print("服务器返回内容：")
        print(response.text[:1000])
        raise


def parse_data(road_show_id, data):

    contest_road_show = data.get("contestRoadShow", {}) or {}

    contest = contest_road_show.get("contest", {}) or {}
    judge_group = contest_road_show.get("judgeGroup", {}) or {}

    contest_name = contest.get("name", "")
    group_name = judge_group.get("name", "")

    projects = data.get("apply_project", []) or []

    rows = []

    for item in projects:

        project = item.get("project", {}) or {}
        category = item.get("applyCategory", {}) or {}

        row = {
            "road_show_id": road_show_id,
            "比赛名称": contest_name,
            "小组": group_name,
            "排名": item.get("rank", ""),
            "组内排名": item.get("group_rank", ""),
            "路演顺序": item.get("road_show_sort", ""),
            "项目名称": project.get("name", ""),
            "项目类别": category.get("name", ""),
            "得分": item.get("grade", ""),
            "project_id": item.get("project_id", ""),
            "school_id": project.get("school_id", ""),
        }

        rows.append(row)

    return rows


def rank_number(value):
    try:
        return int(value)
    except Exception:
        return 999999


def main():

    all_rows = []

    successful = 0

    for road_show_id in ROAD_SHOW_IDS:

        try:
            data = fetch_one(road_show_id)

            rows = parse_data(
                road_show_id,
                data,
            )

            print(
                f"读取成功：road-show/{road_show_id}，"
                f"共 {len(rows)} 个项目"
            )

            all_rows.extend(rows)
            successful += 1

        except Exception as e:
            print(
                f"\n❌ road-show/{road_show_id} 读取失败：{e}",
                file=sys.stderr,
            )

        # 避免连续过快访问服务器
        time.sleep(1)

    if successful == 0:
        print("\n5 个接口全部读取失败。")
        sys.exit(1)

    # --------------------------------------------------------
    # 排序：
    # 第一层按照 road_show_id
    # 第二层按照排名
    # --------------------------------------------------------

    all_rows.sort(
        key=lambda x: (
            ROAD_SHOW_IDS.index(x["road_show_id"]),
            rank_number(x["排名"]),
        )
    )

    output_dir = Path("data")
    output_dir.mkdir(exist_ok=True)

    # ========================================================
    # JSON
    # ========================================================

    json_path = output_dir / "scores.json"

    with open(
        json_path,
        "w",
        encoding="utf-8",
    ) as f:

        json.dump(
            all_rows,
            f,
            ensure_ascii=False,
            indent=2,
        )

    # ========================================================
    # CSV
    # utf-8-sig 可以直接用 Excel 打开，不会中文乱码
    # ========================================================

    csv_path = output_dir / "scores.csv"

    fieldnames = [
        "road_show_id",
        "比赛名称",
        "小组",
        "排名",
        "组内排名",
        "路演顺序",
        "项目名称",
        "项目类别",
        "得分",
        "project_id",
        "school_id",
    ]

    with open(
        csv_path,
        "w",
        newline="",
        encoding="utf-8-sig",
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=fieldnames,
        )

        writer.writeheader()

        writer.writerows(all_rows)

    print("\n==============================")
    print("全部完成")
    print("==============================")
    print(f"成功接口：{successful}/{len(ROAD_SHOW_IDS)}")
    print(f"项目总数：{len(all_rows)}")
    print(f"CSV：{csv_path}")
    print(f"JSON：{json_path}")


if __name__ == "__main__":
    main()