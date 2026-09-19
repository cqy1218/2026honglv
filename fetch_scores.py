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


# ============================================================
# 五个比赛组配置
#
# name:
#     GitHub 首页显示名称
#
# road_show_id:
#     比赛 API 对应 ID
#
# excel_group:
#     “参考逻辑.xlsx” → “路演答辩顺序”
#     中“路演项目组”列对应的准确名称
#
# url:
#     sort_type=2 = 按得分高低顺序
# ============================================================

GROUPS = [
    {
        "name": "公益1组",
        "road_show_id": 150,
        "excel_group": "红旅赛道公益组1",
        "url": (
            "https://itcjspapif.woczx.com/"
            "contest-lives/road-show/150?sort_type=2"
        ),
    },
    {
        "name": "公益2组",
        "road_show_id": 151,
        "excel_group": "红旅赛道公益组2",
        "url": (
            "https://itcjspapif.woczx.com/"
            "contest-lives/road-show/151?sort_type=2"
        ),
    },
    {
        "name": "创意1组",
        "road_show_id": 152,
        "excel_group": "红旅赛道创意组1",
        "url": (
            "https://itcjspapif.woczx.com/"
            "contest-lives/road-show/152?sort_type=2"
        ),
    },
    {
        "name": "创意2组",
        "road_show_id": 135,
        "excel_group": "高教主赛道本科生创意组2",
        "url": (
            "https://itcjspapif.woczx.com/"
            "contest-lives/road-show/135?sort_type=2"
        ),
    },
    {
        "name": "创意3组",
        "road_show_id": 136,
        "excel_group": "高教主赛道本科生创意组3",
        "url": (
            "https://itcjspapif.woczx.com/"
            "contest-lives/road-show/136?sort_type=2"
        ),
    },
]


# ============================================================
# HTTP 请求头
#
# 模拟比赛网页本身的请求。
# 如果以后网站增加 Cookie / Authorization，
# 可以通过 GitHub Secrets 注入。
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


# ============================================================
# 可选 GitHub Secrets
# ============================================================

cookie = os.getenv(
    "WOCZX_COOKIE",
    "",
).strip()

authorization = os.getenv(
    "WOCZX_AUTHORIZATION",
    "",
).strip()

if cookie:
    HEADERS["Cookie"] = cookie

if authorization:
    HEADERS["Authorization"] = authorization


# ============================================================
# README 自动更新标记
# ============================================================

README_START = "<!-- SCOREBOARD_START -->"
README_END = "<!-- SCOREBOARD_END -->"


# ============================================================
# 通用函数
# ============================================================

def clean_text(value):
    """
    将 Excel / API 中的值转换为字符串，
    只删除首尾空白，不改变项目名称内容。
    """

    if value is None:
        return ""

    return str(value).strip()


def rank_number(value):
    """
    将排名转换成整数，用于排序。
    """

    try:
        return int(value)

    except (TypeError, ValueError):
        return 999999


def format_rank(rank):
    """
    1 → 第1名
    2 → 第2名
    """

    if rank in (
        None,
        "",
    ):
        return ""

    return f"第{rank}名"


def escape(value):
    """
    GitHub README HTML 转义。
    """

    if value is None:
        return ""

    return html.escape(
        str(value)
    )


# ============================================================
# 读取 Excel
#
# 核心逻辑：
#
# 不再建立
#
#     项目名称 → 学校
#
# 而是建立
#
#     路演项目组 + 答辩顺序
#                 ↓
#          项目名称 + 学校
#
# 这样不会因为项目名称标点变化造成错配。
# ============================================================

def load_source_mapping():

    if not EXCEL_FILE.exists():

        raise FileNotFoundError(
            f"找不到文件：{EXCEL_FILE}\n"
            "请确保“参考逻辑.xlsx”位于 GitHub 仓库根目录。"
        )

    print("=" * 70)
    print(f"读取 Excel：{EXCEL_FILE}")
    print("=" * 70)

    workbook = load_workbook(
        EXCEL_FILE,
        read_only=True,
        data_only=True,
    )

    sheet_name = "路演答辩顺序"

    if sheet_name not in workbook.sheetnames:

        workbook.close()

        raise RuntimeError(
            f'Excel 中找不到工作表“{sheet_name}”。'
        )

    sheet = workbook[
        sheet_name
    ]

    mapping = {}

    duplicate_keys = []

    # --------------------------------------------------------
    # 当前 Excel 列结构：
    #
    # A：项目名称
    # B：学校名称
    # C：参赛赛道
    # D：参赛组别
    # E：路演项目组
    # F：答辩顺序
    #
    # 第1行为说明
    # 第2行为表头
    # 第3行开始为数据
    # --------------------------------------------------------

    for row in sheet.iter_rows(
        min_row=3,
        values_only=True,
    ):

        if not row:
            continue

        project_name = clean_text(
            row[0]
        )

        school_name = clean_text(
            row[1]
        )

        road_group = clean_text(
            row[4]
        )

        road_order_raw = row[5]

        # 缺少关键字段则跳过
        if not road_group:

            continue

        if road_order_raw in (
            None,
            "",
        ):

            continue

        try:

            road_order = int(
                road_order_raw
            )

        except (
            TypeError,
            ValueError,
        ):

            print(
                "⚠️ 无法识别答辩顺序：",
                road_group,
                road_order_raw,
                project_name,
            )

            continue

        key = (
            road_group,
            road_order,
        )

        if key in mapping:

            old = mapping[key]

            if (
                old["project"]
                != project_name
                or old["school"]
                != school_name
            ):

                duplicate_keys.append(
                    {
                        "key": key,
                        "old": old,
                        "new": {
                            "project":
                                project_name,
                            "school":
                                school_name,
                        },
                    }
                )

            continue

        mapping[key] = {
            "project":
                project_name,

            "school":
                school_name,
        }

    workbook.close()

    print(
        f"Excel 匹配关系读取完成：{len(mapping)} 条"
    )

    if duplicate_keys:

        print()
        print(
            "⚠️ Excel 中发现重复的"
            "“路演项目组 + 答辩顺序”："
        )

        for item in duplicate_keys:

            print(
                f"   {item['key']}"
            )

            print(
                f"      原：{item['old']}"
            )

            print(
                f"      新：{item['new']}"
            )

    print()

    return mapping


# ============================================================
# 读取一个比赛组 API
# ============================================================

def fetch_group(
    session,
    group,
):

    group_name = group[
        "name"
    ]

    url = group[
        "url"
    ]

    print("=" * 70)
    print(
        f"开始读取：{group_name}"
    )
    print(
        f"API：{url}"
    )

    response = session.get(
        url,
        timeout=30,
    )

    print(
        f"HTTP 状态码：{response.status_code}"
    )

    if response.status_code == 401:

        raise RuntimeError(
            f"{group_name} 返回 401 Unauthorized。\n"
            "如果网站增加了身份验证，请检查 "
            "WOCZX_COOKIE / WOCZX_AUTHORIZATION。"
        )

    response.raise_for_status()

    try:

        data = response.json()

    except Exception:

        print(
            "服务器返回内容不是合法 JSON："
        )

        print(
            response.text[:1000]
        )

        raise

    # --------------------------------------------------------
    # 输出 API 自己报告的小组名称，方便人工检查
    # --------------------------------------------------------

    contest_road_show = (
        data.get(
            "contestRoadShow",
            {},
        )
        or {}
    )

    judge_group = (
        contest_road_show.get(
            "judgeGroup",
            {},
        )
        or {}
    )

    api_group_name = clean_text(
        judge_group.get(
            "name"
        )
    )

    if api_group_name:

        print(
            f"API 返回小组：{api_group_name}"
        )

    return data


# ============================================================
# 解析 API 数据
#
# 最关键的匹配方式：
#
# Excel：
#
#     路演项目组
#           +
#       答辩顺序
#
# API：
#
#     excel_group
#           +
#     road_show_sort
#
# 两者组合得到唯一匹配。
# ============================================================

def parse_group(
    group,
    data,
    source_mapping,
):

    group_name = group[
        "name"
    ]

    excel_group = group[
        "excel_group"
    ]

    projects = (
        data.get(
            "apply_project",
            [],
        )
        or []
    )

    result = []

    unmatched = []

    name_mismatches = []

    duplicated_orders = []

    used_orders = set()

    for item in projects:

        project = (
            item.get(
                "project",
                {},
            )
            or {}
        )

        # ----------------------------------------------------
        # API 项目名称
        # ----------------------------------------------------

        api_project_name = clean_text(
            project.get(
                "name"
            )
        )

        # ----------------------------------------------------
        # API 路演顺序
        # ----------------------------------------------------

        road_show_sort_raw = item.get(
            "road_show_sort",
            "",
        )

        try:

            road_show_sort = int(
                road_show_sort_raw
            )

        except (
            TypeError,
            ValueError,
        ):

            road_show_sort = None

        # ----------------------------------------------------
        # 防止 API 出现重复答辩顺序
        # ----------------------------------------------------

        if road_show_sort is not None:

            if road_show_sort in used_orders:

                duplicated_orders.append(
                    road_show_sort
                )

            used_orders.add(
                road_show_sort
            )

        # ----------------------------------------------------
        # 核心匹配
        # ----------------------------------------------------

        source = None

        if road_show_sort is not None:

            source = source_mapping.get(
                (
                    excel_group,
                    road_show_sort,
                )
            )

        # ----------------------------------------------------
        # 匹配成功
        # ----------------------------------------------------

        if source:

            school = source[
                "school"
            ]

            excel_project_name = source[
                "project"
            ]

            # ----------------------------------------------
            # 项目名称不再作为匹配条件。
            #
            # 这里只检查 API 与 Excel 名称是否一致，
            # 方便发现源数据变动。
            # ----------------------------------------------

            if (
                api_project_name
                and excel_project_name
                and api_project_name
                != excel_project_name
            ):

                name_mismatches.append(
                    {
                        "order":
                            road_show_sort,

                        "api":
                            api_project_name,

                        "excel":
                            excel_project_name,

                        "school":
                            school,
                    }
                )

        # ----------------------------------------------------
        # 匹配失败
        # ----------------------------------------------------

        else:

            school = "⚠️ 未匹配"

            excel_project_name = ""

            unmatched.append(
                {
                    "order":
                        road_show_sort_raw,

                    "project":
                        api_project_name,
                }
            )

        # ----------------------------------------------------
        # 排名
        # ----------------------------------------------------

        rank = item.get(
            "rank"
        )

        if rank in (
            None,
            "",
        ):

            rank = item.get(
                "group_rank"
            )

        # ----------------------------------------------------
        # 得分
        # ----------------------------------------------------

        score = item.get(
            "grade",
            "",
        )

        # ----------------------------------------------------
        # 加入结果
        # ----------------------------------------------------

        result.append(
            {
                "group":
                    group_name,

                "excel_group":
                    excel_group,

                "school":
                    school,

                # 页面始终展示 API 最新名称
                "project":
                    api_project_name,

                "excel_project":
                    excel_project_name,

                "score":
                    score,

                "rank":
                    rank,

                "rank_text":
                    format_rank(
                        rank
                    ),

                "road_show_sort":
                    road_show_sort_raw,

                "project_id":
                    item.get(
                        "project_id",
                        "",
                    ),

                "school_id":
                    project.get(
                        "school_id",
                        "",
                    ),
            }
        )

    # --------------------------------------------------------
    # 按最终排名排序
    # --------------------------------------------------------

    result.sort(
        key=lambda x:
        rank_number(
            x["rank"]
        )
    )

    # --------------------------------------------------------
    # 输出校验结果
    # --------------------------------------------------------

    print(
        f"{group_name}：读取到 {len(result)} 个项目"
    )

    print(
        f"Excel 对应组：{excel_group}"
    )

    if duplicated_orders:

        print()
        print(
            f"❌ {group_name} API 中存在重复答辩顺序："
        )

        for order in duplicated_orders:

            print(
                f"   答辩顺序 {order}"
            )

    if unmatched:

        print()
        print(
            f"❌ {group_name} 存在无法匹配学校的项目："
        )

        for item in unmatched:

            print(
                f"   答辩顺序：{item['order']}"
            )

            print(
                f"   项目名称：{item['project']}"
            )

    if name_mismatches:

        print()
        print(
            f"⚠️ {group_name} "
            "存在 API 名称与 Excel 名称不一致："
        )

        for item in name_mismatches:

            print(
                f"   答辩顺序 {item['order']}"
            )

            print(
                f"      API   ：{item['api']}"
            )

            print(
                f"      Excel ：{item['excel']}"
            )

            print(
                f"      学校  ：{item['school']}"
            )

    if (
        not unmatched
        and not duplicated_orders
    ):

        print(
            f"✅ {group_name} 学校匹配正常"
        )

    print()

    return result


# ============================================================
# 用于比较本次榜单和上一次榜单
#
# 不加入 updated_at，
# 防止每次运行都因为时间不同产生 commit。
# ============================================================

def comparable_data(
    groups_data,
):

    return {
        group_name: [
            {
                "school":
                    row["school"],

                "project":
                    row["project"],

                "score":
                    row["score"],

                "rank":
                    row["rank"],

                "road_show_sort":
                    row["road_show_sort"],

                "project_id":
                    row["project_id"],
            }

            for row in rows
        ]

        for (
            group_name,
            rows,
        ) in groups_data.items()
    }


# ============================================================
# 读取上一次保存的数据
# ============================================================

def load_previous_data():

    if not JSON_FILE.exists():

        return None

    try:

        with JSON_FILE.open(
            "r",
            encoding="utf-8",
        ) as f:

            data = json.load(
                f
            )

        if isinstance(
            data,
            dict,
        ):

            return data.get(
                "groups"
            )

        return None

    except Exception as exc:

        print(
            f"⚠️ 无法读取旧 scores.json：{exc}"
        )

        return None


# ============================================================
# 写 CSV
# ============================================================

def write_csv(
    groups_data,
):

    DATA_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    fieldnames = [
        "组别",
        "Excel路演项目组",
        "学校",
        "项目名称",
        "Excel项目名称",
        "得分",
        "名次",
        "路演顺序",
        "project_id",
        "school_id",
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

        for (
            group_name,
            rows,
        ) in groups_data.items():

            for row in rows:

                writer.writerow(
                    {
                        "组别":
                            group_name,

                        "Excel路演项目组":
                            row[
                                "excel_group"
                            ],

                        "学校":
                            row[
                                "school"
                            ],

                        "项目名称":
                            row[
                                "project"
                            ],

                        "Excel项目名称":
                            row[
                                "excel_project"
                            ],

                        "得分":
                            row[
                                "score"
                            ],

                        "名次":
                            row[
                                "rank_text"
                            ],

                        "路演顺序":
                            row[
                                "road_show_sort"
                            ],

                        "project_id":
                            row[
                                "project_id"
                            ],

                        "school_id":
                            row[
                                "school_id"
                            ],
                    }
                )


# ============================================================
# 生成 README 首页榜单 HTML
# ============================================================

def generate_scoreboard_html(
    groups_data,
    updated_at,
):

    group_names = [
        group["name"]
        for group in GROUPS
    ]

    max_rows = max(
        len(
            groups_data.get(
                group_name,
                [],
            )
        )
        for group_name in group_names
    )

    lines = []

    lines.append(
        '<div align="center">'
    )

    lines.append(
        "<h2>“建行杯”江苏大学生创新大赛（2026）实时成绩</h2>"
    )

    lines.append(
        f"<p>最近一次数据变化：{escape(updated_at)}</p>"
    )

    lines.append(
        "</div>"
    )

    lines.append("")

    lines.append(
        "<table>"
    )

    # --------------------------------------------------------
    # 第一层表头
    # --------------------------------------------------------

    lines.append(
        "<thead>"
    )

    lines.append(
        "<tr>"
    )

    for group_name in group_names:

        lines.append(
            '<th colspan="4" align="center">'
            f"<strong>{escape(group_name)}</strong>"
            "</th>"
        )

    lines.append(
        "</tr>"
    )

    # --------------------------------------------------------
    # 第二层表头
    # --------------------------------------------------------

    lines.append(
        "<tr>"
    )

    for _ in group_names:

        lines.append(
            '<th align="center">学校</th>'
        )

        lines.append(
            '<th align="center">项目名称</th>'
        )

        lines.append(
            '<th align="center">得分</th>'
        )

        lines.append(
            '<th align="center">名次</th>'
        )

    lines.append(
        "</tr>"
    )

    lines.append(
        "</thead>"
    )

    # --------------------------------------------------------
    # 表格正文
    # --------------------------------------------------------

    lines.append(
        "<tbody>"
    )

    for index in range(
        max_rows
    ):

        lines.append(
            "<tr>"
        )

        for group_name in group_names:

            rows = groups_data.get(
                group_name,
                [],
            )

            if index < len(
                rows
            ):

                row = rows[
                    index
                ]

                lines.append(
                    f"<td>{escape(row['school'])}</td>"
                )

                lines.append(
                    f"<td>{escape(row['project'])}</td>"
                )

                lines.append(
                    '<td align="center">'
                    f"{escape(row['score'])}"
                    "</td>"
                )

                lines.append(
                    '<td align="center">'
                    f"<strong>{escape(row['rank_text'])}</strong>"
                    "</td>"
                )

            else:

                lines.extend(
                    [
                        "<td></td>",
                        "<td></td>",
                        "<td></td>",
                        "<td></td>",
                    ]
                )

        lines.append(
            "</tr>"
        )

    lines.append(
        "</tbody>"
    )

    lines.append(
        "</table>"
    )

    lines.append("")

    lines.append(
        "> 成绩由 GitHub Actions 自动读取比赛接口；"
        "学校信息按照仓库根目录 `参考逻辑.xlsx` "
        "中的“路演项目组 + 答辩顺序”进行匹配。"
    )

    return "\n".join(
        lines
    )


# ============================================================
# 更新 README
#
# 只修改：
#
# <!-- SCOREBOARD_START -->
#
# ...
#
# <!-- SCOREBOARD_END -->
#
# 之间的内容。
# ============================================================

def update_readme(
    scoreboard_html,
):

    block = (
        f"{README_START}\n"
        f"{scoreboard_html}\n"
        f"{README_END}"
    )

    if README_FILE.exists():

        existing = README_FILE.read_text(
            encoding="utf-8",
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
            re.escape(
                README_START
            )
            + r".*?"
            + re.escape(
                README_END
            ),
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
# 主程序
# ============================================================

def main():

    print()
    print("=" * 70)
    print("江苏大学生创新大赛实时成绩自动更新")
    print("=" * 70)
    print()

    # --------------------------------------------------------
    # 1. 读取 Excel
    # --------------------------------------------------------

    source_mapping = (
        load_source_mapping()
    )

    # --------------------------------------------------------
    # 2. 创建 HTTP Session
    # --------------------------------------------------------

    session = requests.Session()

    session.headers.update(
        HEADERS
    )

    # --------------------------------------------------------
    # 3. 读取五个 API
    # --------------------------------------------------------

    groups_data = {}

    failed_groups = []

    for group in GROUPS:

        try:

            raw_data = fetch_group(
                session,
                group,
            )

            rows = parse_group(
                group,
                raw_data,
                source_mapping,
            )

            groups_data[
                group["name"]
            ] = rows

        except Exception as exc:

            failed_groups.append(
                group["name"]
            )

            print(
                f"❌ {group['name']} 读取失败：{exc}",
                file=sys.stderr,
            )

        # 避免请求速度过快
        time.sleep(
            1
        )

    # --------------------------------------------------------
    # 4. 任何一个组失败，都不覆盖旧榜单
    # --------------------------------------------------------

    if failed_groups:

        print()
        print("=" * 70)
        print(
            "❌ 本次更新失败"
        )
        print("=" * 70)

        print(
            "以下小组读取失败："
            + "、".join(
                failed_groups
            )
        )

        print(
            "为保护已有成绩，本次不会覆盖 README、CSV、JSON。"
        )

        sys.exit(
            1
        )

    # --------------------------------------------------------
    # 5. 检查项目数量
    # --------------------------------------------------------

    total_projects = sum(
        len(rows)
        for rows in groups_data.values()
    )

    print("=" * 70)
    print(
        f"五个小组全部读取成功，共 {total_projects} 个项目"
    )
    print("=" * 70)
    print()

    for group in GROUPS:

        name = group[
            "name"
        ]

        print(
            f"{name}："
            f"{len(groups_data.get(name, []))} 个项目"
        )

    print()

    # --------------------------------------------------------
    # 6. 比较是否发生变化
    # --------------------------------------------------------

    current_data = comparable_data(
        groups_data
    )

    previous_data = (
        load_previous_data()
    )

    if previous_data == current_data:

        print("=" * 70)
        print(
            "✅ 当前榜单与上一次完全一致"
        )
        print(
            "无需修改 README / CSV / JSON"
        )
        print("=" * 70)

        return

    # --------------------------------------------------------
    # 7. 数据发生变化
    # --------------------------------------------------------

    now = datetime.now(
        ZoneInfo(
            "Asia/Shanghai"
        )
    ).strftime(
        "%Y-%m-%d %H:%M:%S"
    )

    print("=" * 70)
    print(
        "🔄 检测到榜单变化，开始更新文件"
    )
    print(
        f"更新时间：{now}"
    )
    print("=" * 70)
    print()

    DATA_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    # --------------------------------------------------------
    # 8. 写 JSON
    # --------------------------------------------------------

    json_payload = {
        "updated_at":
            now,

        "groups":
            current_data,
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

    print(
        f"✅ 已生成：{JSON_FILE}"
    )

    # --------------------------------------------------------
    # 9. 写 CSV
    # --------------------------------------------------------

    write_csv(
        groups_data
    )

    print(
        f"✅ 已生成：{CSV_FILE}"
    )

    # --------------------------------------------------------
    # 10. 更新 README
    # --------------------------------------------------------

    scoreboard_html = (
        generate_scoreboard_html(
            groups_data,
            now,
        )
    )

    update_readme(
        scoreboard_html
    )

    print(
        f"✅ 已更新：{README_FILE}"
    )

    print()
    print("=" * 70)
    print(
        "✅ 本次成绩更新完成"
    )
    print("=" * 70)


# ============================================================
# 程序入口
# ============================================================

if __name__ == "__main__":

    main()