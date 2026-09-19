import csv
import html
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from openpyxl import load_workbook


# ============================================================
# 基础配置
# ============================================================

EXCEL_FILE = Path("参考逻辑.xlsx")

DATA_DIR = Path("data")
JSON_FILE = DATA_DIR / "scores.json"
CSV_FILE = DATA_DIR / "scores.csv"
README_FILE = Path("README.md")


# 按 Excel “展示”页从左到右的顺序
GROUPS = [
    {
        "name": "公益1组",
        "url": "https://itcjspapif.woczx.com/contest-lives/road-show/150?sort_type=2",
    },
    {
        "name": "公益2组",
        "url": "https://itcjspapif.woczx.com/contest-lives/road-show/151?sort_type=2",
    },
    {
        "name": "创意1组",
        "url": "https://itcjspapif.woczx.com/contest-lives/road-show/152?sort_type=2",
    },
    {
        "name": "创意2组",
        "url": "https://itcjspapif.woczx.com/contest-lives/road-show/135?sort_type=2",
    },
    {
        "name": "创意3组",
        "url": "https://itcjspapif.woczx.com/contest-lives/road-show/136?sort_type=2",
    },
]


# ============================================================
# 模拟网页请求
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
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
}


# 如果网站以后需要 Cookie / Authorization，
# 可以放到 GitHub Secrets，不需要写进公开代码
cookie = os.getenv("WOCZX_COOKIE", "").strip()
authorization = os.getenv("WOCZX_AUTHORIZATION", "").strip()

if cookie:
    HEADERS["Cookie"] = cookie

if authorization:
    HEADERS["Authorization"] = authorization


# ============================================================
# 工具函数
# ============================================================

def clean_text(value):
    """只去除首尾空白，不修改项目名称中的标点等内容。"""
    if value is None:
        return ""
    return str(value).strip()


def rank_number(value):
    """用于排名排序。"""
    try:
        return int(value)
    except (TypeError, ValueError):
        return 999999


def format_rank(rank):
    """将 1 转换成 Excel 里的“第1名”形式。"""
    if rank in (None, ""):
        return ""
    return f"第{rank}名"


def escape(value):
    """用于生成安全的 HTML。"""
    if value is None:
        return ""
    return html.escape(str(value))


# ============================================================
# 第一步：读取 Excel 中的“项目名称 → 学校名称”
# ============================================================

def load_school_mapping():
    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            f"找不到 {EXCEL_FILE}。\n"
            "请确保“参考逻辑.xlsx”已经上传到 GitHub 仓库根目录。"
        )

    print(f"正在读取 Excel：{EXCEL_FILE}")

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    if "路演答辩顺序" not in workbook.sheetnames:
        raise RuntimeError(
            'Excel 中找不到工作表“路演答辩顺序”。'
        )

    sheet = workbook["路演答辩顺序"]

    mapping = {}
    duplicates = []

    # Excel 中：
    # A列 = 项目名称
    # B列 = 学校名称
    #
    # 第1行为来源说明
    # 第2行为标题
    # 第3行开始是真正的数据
    for row in sheet.iter_rows(
        min_row=3,
        values_only=True,
    ):
        project_name = clean_text(row[0])
        school_name = clean_text(row[1])

        if not project_name:
            continue

        if (
            project_name in mapping
            and mapping[project_name] != school_name
        ):
            duplicates.append(project_name)
            continue

        mapping[project_name] = school_name

    workbook.close()

    print(f"已读取 {len(mapping)} 条项目-学校对应关系")

    if duplicates:
        print(
            "⚠️ 以下项目名称在 Excel 中存在重复且学校不同：",
            duplicates,
        )

    return mapping


# ============================================================
# 第二步：读取比赛 API
# ============================================================

def fetch_group(session, group):
    group_name = group["name"]
    url = group["url"]

    print()
    print("=" * 60)
    print(f"读取：{group_name}")
    print(url)

    response = session.get(
        url,
        timeout=30,
    )

    print(f"HTTP：{response.status_code}")

    if response.status_code == 401:
        raise RuntimeError(
            f"{group_name} 返回 401 Unauthorized。"
            "如网站增加登录验证，请配置 GitHub Secrets。"
        )

    response.raise_for_status()

    data = response.json()

    # 核对接口本身报告的小组名称
    contest_road_show = data.get("contestRoadShow", {}) or {}
    judge_group = contest_road_show.get("judgeGroup", {}) or {}
    api_group_name = clean_text(judge_group.get("name"))

    if api_group_name:
        print(f"API 小组名称：{api_group_name}")

        # 比如 Excel 是“创意3组”，API 可能是“本科生创意3组”
        if group_name not in api_group_name:
            print(
                f"⚠️ 注意：配置名称“{group_name}”"
                f"与 API 返回名称“{api_group_name}”不完全对应。"
            )

    return data


# ============================================================
# 第三步：解析项目
# ============================================================

def parse_group(group_name, data, school_mapping):
    projects = data.get("apply_project", []) or []

    result = []

    unmatched = []

    for item in projects:
        project = item.get("project", {}) or {}

        project_name = clean_text(project.get("name"))

        # Excel VLOOKUP(FALSE) 的核心逻辑：
        # 根据项目名称精确查找学校
        school = school_mapping.get(project_name, "")

        if not school:
            unmatched.append(project_name)
            school = "⚠️ 未匹配"

        # 优先用网页最终 rank；
        # 如果没有则使用 group_rank
        rank = item.get("rank")

        if rank in (None, ""):
            rank = item.get("group_rank")

        score = item.get("grade", "")

        result.append(
            {
                "group": group_name,
                "school": school,
                "project": project_name,
                "score": score,
                "rank": rank,
                "rank_text": format_rank(rank),
                "road_show_sort": item.get(
                    "road_show_sort",
                    "",
                ),
                "project_id": item.get(
                    "project_id",
                    "",
                ),
            }
        )

    # 与 Excel“按得分高低顺序”后的显示一致：
    # 最终以排名顺序排列
    result.sort(
        key=lambda x: rank_number(x["rank"])
    )

    if unmatched:
        print()
        print(f"⚠️ {group_name} 有项目无法在 Excel 中找到学校：")

        for project_name in unmatched:
            print(f"   - {project_name}")

    print(f"{group_name}：读取 {len(result)} 个项目")

    return result


# ============================================================
# 第四步：生成 CSV
# ============================================================

def write_csv(groups_data):
    DATA_DIR.mkdir(exist_ok=True)

    fieldnames = [
        "组别",
        "学校",
        "项目名称",
        "得分",
        "名次",
        "路演顺序",
        "project_id",
    ]

    with CSV_FILE.open(
        "w",
        newline="",
        encoding="utf-8-sig",
    ) as f:

        writer = csv.DictWriter(
            f,
            fieldnames=fieldnames,
        )

        writer.writeheader()

        for group_name, rows in groups_data.items():

            for row in rows:
                writer.writerow(
                    {
                        "组别": group_name,
                        "学校": row["school"],
                        "项目名称": row["project"],
                        "得分": row["score"],
                        "名次": row["rank_text"],
                        "路演顺序": row["road_show_sort"],
                        "project_id": row["project_id"],
                    }
                )


# ============================================================
# 第五步：生成 GitHub 首页榜单
# ============================================================

def generate_scoreboard_html(groups_data, updated_at):
    group_names = [
        group["name"]
        for group in GROUPS
    ]

    # 找到项目最多的小组
    max_rows = max(
        len(groups_data.get(name, []))
        for name in group_names
    )

    lines = []

    lines.append(
        '<div align="center">'
    )

    lines.append(
        "<h2>“建行杯”江苏大学生创新大赛（2026）实时成绩</h2>"
    )

    lines.append(
        f"<p>最近一次成绩变化：{escape(updated_at)}</p>"
    )

    lines.append(
        "</div>"
    )

    lines.append("")
    lines.append("<table>")

    # --------------------------------------------------------
    # 第一层表头：五个小组
    # --------------------------------------------------------

    lines.append("<thead>")
    lines.append("<tr>")

    for group_name in group_names:
        lines.append(
            f'<th colspan="4" align="center">'
            f"<strong>{escape(group_name)}</strong>"
            f"</th>"
        )

    lines.append("</tr>")

    # --------------------------------------------------------
    # 第二层表头
    # --------------------------------------------------------

    lines.append("<tr>")

    for _ in group_names:
        lines.append('<th align="center">学校</th>')
        lines.append('<th align="center">项目名称</th>')
        lines.append('<th align="center">得分</th>')
        lines.append('<th align="center">名次</th>')

    lines.append("</tr>")
    lines.append("</thead>")

    # --------------------------------------------------------
    # 数据
    # --------------------------------------------------------

    lines.append("<tbody>")

    for index in range(max_rows):

        lines.append("<tr>")

        for group_name in group_names:

            rows = groups_data.get(group_name, [])

            if index < len(rows):
                row = rows[index]

                lines.append(
                    f"<td>{escape(row['school'])}</td>"
                )

                lines.append(
                    f"<td>{escape(row['project'])}</td>"
                )

                lines.append(
                    f'<td align="center">'
                    f"{escape(row['score'])}"
                    f"</td>"
                )

                lines.append(
                    f'<td align="center">'
                    f"<strong>{escape(row['rank_text'])}</strong>"
                    f"</td>"
                )

            else:
                # 保持五组横向对齐
                lines.extend(
                    [
                        "<td></td>",
                        "<td></td>",
                        "<td></td>",
                        "<td></td>",
                    ]
                )

        lines.append("</tr>")

    lines.append("</tbody>")
    lines.append("</table>")

    lines.append("")
    lines.append(
        "> 数据由 GitHub Actions 自动获取；"
        "学校名称根据仓库根目录 `参考逻辑.xlsx` "
        "中的“路演答辩顺序”工作表匹配。"
    )

    return "\n".join(lines)


# ============================================================
# 第六步：只更新 README 中的成绩区域
# ============================================================

README_START = "<!-- SCOREBOARD_START -->"
README_END = "<!-- SCOREBOARD_END -->"


def update_readme(scoreboard_html):
    block = (
        f"{README_START}\n"
        f"{scoreboard_html}\n"
        f"{README_END}"
    )

    if README_FILE.exists():
        existing = README_FILE.read_text(
            encoding="utf-8"
        )
    else:
        existing = (
            "# 江苏大学生创新大赛实时成绩\n"
        )

    if (
        README_START in existing
        and README_END in existing
    ):

        pattern = re.compile(
            re.escape(README_START)
            + r".*?"
            + re.escape(README_END),
            re.S,
        )

        new_content = pattern.sub(
            block,
            existing,
        )

    else:
        new_content = (
            existing.rstrip()
            + "\n\n"
            + block
            + "\n"
        )

    README_FILE.write_text(
        new_content,
        encoding="utf-8",
    )


# ============================================================
# 第七步：判断榜单有没有变化
# ============================================================

def comparable_data(groups_data):
    """
    去除时间等信息，只比较实际榜单。
    """

    return {
        group_name: [
            {
                "school": row["school"],
                "project": row["project"],
                "score": row["score"],
                "rank": row["rank"],
                "road_show_sort": row["road_show_sort"],
                "project_id": row["project_id"],
            }
            for row in rows
        ]
        for group_name, rows in groups_data.items()
    }


def load_previous_data():
    if not JSON_FILE.exists():
        return None

    try:
        with JSON_FILE.open(
            "r",
            encoding="utf-8",
        ) as f:
            data = json.load(f)

        return data.get("groups")

    except Exception:
        return None


# ============================================================
# 主程序
# ============================================================

def main():
    print("=" * 60)
    print("江苏大学生创新大赛成绩自动更新")
    print("=" * 60)

    school_mapping = load_school_mapping()

    session = requests.Session()
    session.headers.update(HEADERS)

    groups_data = {}

    failed_groups = []

    for group in GROUPS:

        try:
            raw = fetch_group(
                session,
                group,
            )

            rows = parse_group(
                group["name"],
                raw,
                school_mapping,
            )

            groups_data[group["name"]] = rows

        except Exception as exc:

            failed_groups.append(
                group["name"]
            )

            print(
                f"❌ {group['name']} 读取失败：{exc}",
                file=sys.stderr,
            )

        # 避免过快访问
        time.sleep(1)

    # --------------------------------------------------------
    # 为防止某一组接口失败导致首页榜单被清空，
    # 只要存在读取失败，本轮就直接退出，不覆盖旧榜单。
    # --------------------------------------------------------

    if failed_groups:
        print()
        print(
            "❌ 以下小组读取失败："
            + "、".join(failed_groups)
        )

        print(
            "为保护已有榜单，本次不更新 README。"
        )

        sys.exit(1)

    current_data = comparable_data(
        groups_data
    )

    previous_data = load_previous_data()

    # --------------------------------------------------------
    # 如果实际成绩完全没有变化，就什么都不写
    # 避免 GitHub 每10分钟制造一个无意义 commit
    # --------------------------------------------------------

    if previous_data == current_data:
        print()
        print("✅ 榜单与上一次完全一致，无需更新。")
        return

    # --------------------------------------------------------
    # 有变化才更新时间、README、CSV、JSON
    # --------------------------------------------------------

    now = datetime.now(
        ZoneInfo("Asia/Shanghai")
    ).strftime("%Y-%m-%d %H:%M:%S")

    DATA_DIR.mkdir(exist_ok=True)

    json_payload = {
        "updated_at": now,
        "groups": current_data,
    }

    with JSON_FILE.open(
        "w",
        encoding="utf-8",
    ) as f:

        json.dump(
            json_payload,
            f,
            ensure_ascii=False,
            indent=2,
        )

    write_csv(groups_data)

    scoreboard = generate_scoreboard_html(
        groups_data,
        now,
    )

    update_readme(scoreboard)

    print()
    print("=" * 60)
    print("✅ 榜单发生变化")
    print("=" * 60)
    print(f"更新时间：{now}")
    print(f"JSON：{JSON_FILE}")
    print(f"CSV：{CSV_FILE}")
    print(f"首页：{README_FILE}")


if __name__ == "__main__":
    main()