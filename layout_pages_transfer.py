import argparse
import asyncio
import base64
import csv
import io
import json
import pathlib

import httpx
import pypdfium2 as pdfium


BASE_URL = "http://172.168.48.51:4002"
CSV_PATH = pathlib.Path("风水书籍转换/1-DD-RAG应用 - B-步骤2-PDF章节分类.csv")
PDF_DIR = pathlib.Path(r"D:\pythonprojects\风水图片rag测试\测试书籍")
OUTPUT_DIR = pathlib.Path("风水书籍转换")


EXCLUDE_STRUCTURES = []
EXCLUDE_SCENES = ["商业建筑", "其他建筑"]
EXCLUDE_BRANCHES = ["周边环境", "公共空间"]
EXCLUDE_PAIRS = []  # 例如 [("住宅", "通用"), ("商业建筑", "周边环境")]
DRY_RUN = False


def load_tasks():
    tasks = []
    occupied_pages = {}

    with CSV_PATH.open("r", encoding="utf-8-sig", newline="") as f:
        for row_number, row in enumerate(csv.DictReader(f), start=2):
            chapter = row["开头章节标题"].strip()
            if not chapter or row["属性"].strip() == "说明":
                continue

            filename = row["文件名称"].strip()
            pdf_path = PDF_DIR / f"{filename}.pdf"
            if not pdf_path.is_file():
                raise FileNotFoundError(f"CSV 第 {row_number} 行找不到 PDF：{pdf_path}")

            try:
                start_page = int(row["开头页码"].strip())
            except ValueError as exc:
                raise ValueError(f"CSV 第 {row_number} 行开头页码必须是数字") from exc

            with pdfium.PdfDocument(str(pdf_path)) as pdf:
                page_count = len(pdf)

            end_value = row["结尾页码"].strip()
            try:
                end_page = page_count + 1 if end_value == "-" else int(end_value)
            except ValueError as exc:
                raise ValueError(f"CSV 第 {row_number} 行结尾页码必须是数字或 -") from exc

            if not 1 <= start_page < end_page <= page_count + 1:
                raise ValueError(
                    f"CSV 第 {row_number} 行页码范围 [{start_page}, {end_page}) "
                    f"超出 {pdf_path.name} 的 1-{page_count} 页"
                )

            pages = set(range(start_page, end_page))
            overlap = occupied_pages.setdefault(pdf_path, set()) & pages
            if overlap:
                raise ValueError(
                    f"CSV 第 {row_number} 行与已有范围重叠：{pdf_path.name} 第 {min(overlap)} 页"
                )
            occupied_pages[pdf_path].update(pages)

            metadata = {}
            for key in ("文件结构分类", "建筑应用场景", "建筑应用场景子分支"):
                value = row[key].strip()
                if not value:
                    raise ValueError(f"CSV 第 {row_number} 行缺少 metadata：{key}")
                metadata[key] = value

            tasks.append(
                {
                    "row_number": row_number,
                    "chapter": chapter,
                    "pdf_path": pdf_path,
                    "start_page": start_page,
                    "end_page": end_page,
                    "metadata": metadata,
                }
            )

    return tasks


def parse_pairs(values):
    pairs = set()
    for value in values:
        if isinstance(value, str):
            scene, separator, branch = value.partition("-")
            if not separator or not scene.strip() or not branch.strip():
                raise ValueError(f"交集筛选必须使用 应用场景-子分支 格式：{value}")
        else:
            try:
                scene, branch = value
            except (TypeError, ValueError) as exc:
                raise ValueError(f"交集筛选必须包含两个值：{value}") from exc
        pairs.add((scene.strip(), branch.strip()))
    return pairs


def resolve_config(args):
    pair_values = EXCLUDE_PAIRS if args.exclude_pair is None else args.exclude_pair
    return {
        "structures": set(
            EXCLUDE_STRUCTURES
            if args.exclude_structure is None
            else args.exclude_structure
        ),
        "scenes": set(EXCLUDE_SCENES if args.exclude_scene is None else args.exclude_scene),
        "branches": set(
            EXCLUDE_BRANCHES if args.exclude_branch is None else args.exclude_branch
        ),
        "pairs": parse_pairs(pair_values),
        "dry_run": DRY_RUN if args.dry_run is None else args.dry_run,
    }


def exclusion_reason(metadata, config):
    reasons = []
    if metadata["文件结构分类"] in config["structures"]:
        reasons.append(f"文件结构分类={metadata['文件结构分类']}")
    if metadata["建筑应用场景"] in config["scenes"]:
        reasons.append(f"应用场景={metadata['建筑应用场景']}")
    if metadata["建筑应用场景子分支"] in config["branches"]:
        reasons.append(f"子分支={metadata['建筑应用场景子分支']}")
    pair = (metadata["建筑应用场景"], metadata["建筑应用场景子分支"])
    if pair in config["pairs"]:
        reasons.append(f"交集={pair[0]}-{pair[1]}")
    return "；".join(reasons)


def find_unknown_filters(tasks, config):
    available = {
        "structures": {task["metadata"]["文件结构分类"] for task in tasks},
        "scenes": {task["metadata"]["建筑应用场景"] for task in tasks},
        "branches": {task["metadata"]["建筑应用场景子分支"] for task in tasks},
        "pairs": {
            (
                task["metadata"]["建筑应用场景"],
                task["metadata"]["建筑应用场景子分支"],
            )
            for task in tasks
        },
    }
    unknown = {
        "文件结构分类": sorted(config["structures"] - available["structures"]),
        "建筑应用场景": sorted(config["scenes"] - available["scenes"]),
        "建筑应用场景子分支": sorted(config["branches"] - available["branches"]),
        "应用场景-子分支组合": [
            f"{scene}-{branch}"
            for scene, branch in sorted(config["pairs"] - available["pairs"])
        ],
    }
    return {name: values for name, values in unknown.items() if values}


def print_unknown_filters(unknown):
    print("警告：当前 CSV 未出现以下筛选条件：")
    for name, values in unknown.items():
        print(f"  {name}：{', '.join(values)}")
    print("这些条件本次不会排除任何章节，但以后 CSV 出现对应值时会生效。")


def confirm_unknown_filters(unknown):
    print_unknown_filters(unknown)
    while True:
        try:
            answer = input("是否继续转换？[y/n]: ").strip().lower()
        except EOFError:
            answer = "n"
        if answer in {"y", "yes"}:
            return True
        if answer in {"", "n", "no"}:
            return False
        print("请输入 y 或 n。")


def filter_tasks(tasks, config):
    included = []
    excluded = []
    for task in tasks:
        reason = exclusion_reason(task["metadata"], config)
        (excluded if reason else included).append((task, reason))
    return included, excluded


def print_summary(included, excluded, config):
    pairs = [f"{scene}-{branch}" for scene, branch in sorted(config["pairs"])]
    print(
        f"反向筛选：文件结构分类={sorted(config['structures']) or '无'}，"
        f"应用场景={sorted(config['scenes']) or '无'}，"
        f"子分支={sorted(config['branches']) or '无'}，"
        f"交集={pairs or '无'}"
    )
    for task, reason in included:
        print(
            f"保留：{task['pdf_path'].stem} [{task['start_page']}, {task['end_page']}) "
            f"{task['chapter']}"
        )
    for task, reason in excluded:
        print(
            f"排除：{task['pdf_path'].stem} [{task['start_page']}, {task['end_page']}) "
            f"{task['chapter']}（{reason}）"
        )

    included_pages = sum(task["end_page"] - task["start_page"] for task, _ in included)
    excluded_pages = sum(task["end_page"] - task["start_page"] for task, _ in excluded)
    print(
        f"合计：保留 {len(included)} 个区间/{included_pages} 页，"
        f"排除 {len(excluded)} 个区间/{excluded_pages} 页"
    )

    old_pages = sum(
        (OUTPUT_DIR / task["pdf_path"].stem / f"page_{page}.md").is_file()
        for task, _ in excluded
        for page in range(task["start_page"], task["end_page"])
    )
    if old_pages:
        print(f"警告：被排除范围中已有 {old_pages} 个旧 Markdown，本次不会删除")


def extract_pdf_range(pdf_path, start_page, end_page):
    output = io.BytesIO()
    with pdfium.PdfDocument(str(pdf_path)) as source, pdfium.PdfDocument.new() as selected:
        selected.import_pages(source, pages=list(range(start_page - 1, end_page - 1)))
        selected.save(output)
    return output.getvalue()


def _split_data_uri(image_b64):
    if image_b64.lstrip().startswith("data:image/") and "," in image_b64:
        return image_b64.split(",", 1)[1]
    return image_b64


def _image_suffix(image_bytes):
    if image_bytes.startswith(b"\xff\xd8\xff"):
        return ".jpg"
    if image_bytes.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png"
    if image_bytes.startswith((b"GIF87a", b"GIF89a")):
        return ".gif"
    if image_bytes.startswith(b"RIFF") and image_bytes[8:12] == b"WEBP":
        return ".webp"
    raise ValueError("无法识别布局图片格式")


def save_layout_outputs(layout_results, pdf_path, start_page):
    layout_dir = OUTPUT_DIR / f"{pdf_path.stem}-layout"
    layout_dir.mkdir(parents=True, exist_ok=True)

    for page_number, result in enumerate(layout_results, start=start_page):
        (layout_dir / f"page_{page_number}.json").write_text(
            json.dumps(result["prunedResult"], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        output_images = result.get("outputImages") or {}
        if len(output_images) != 1:
            raise RuntimeError(
                f"{pdf_path.name} 第 {page_number} 页应返回 1 张布局图，"
                f"实际返回 {len(output_images)} 张"
            )
        image_b64 = next(iter(output_images.values()))
        try:
            image_bytes = base64.b64decode(_split_data_uri(image_b64), validate=True)
            suffix = _image_suffix(image_bytes)
        except (ValueError, TypeError) as exc:
            raise RuntimeError(f"{pdf_path.name} 第 {page_number} 页布局图解码失败") from exc

        image_path = layout_dir / f"page_{page_number}{suffix}"
        image_path.write_bytes(image_bytes)
        print(f"已保存：{image_path}")


def markdown_with_metadata(result, pdf_name, metadata):
    markdown = result["markdown"]
    text = markdown["text"]
    for image_path, image_b64 in (markdown.get("images") or {}).items():
        image_uri = (
            image_b64
            if image_b64.lstrip().startswith("data:image/")
            else f"data:image/png;base64,{image_b64}"
        )
        text = text.replace(f"]({image_path})", f"]({image_uri})")
        text = text.replace(f'src="{image_path}"', f'src="{image_uri}"')

    values = {"文件名": pdf_name, **metadata}
    front_matter = "\n".join(
        f"{key}: {json.dumps(value, ensure_ascii=False)}" for key, value in values.items()
    )
    return f"---\n{front_matter}\n---\n\n{text}"


async def transfer_task(task, client):
    pdf_path = task["pdf_path"]
    start_page = task["start_page"]
    end_page = task["end_page"]
    expected_pages = end_page - start_page
    pdf_b64 = base64.b64encode(
        extract_pdf_range(pdf_path, start_page, end_page)
    ).decode("ascii")

    response = await client.post(
        f"{BASE_URL}/layout-parsing",
        json={
            "file": pdf_b64,
            "fileType": 0,
            "visualize": True,
            "maxNewTokens": 512,
            "layoutThreshold": 0.7,
            "useDocOrientationClassify": False,
            "useDocUnwarping": False,
            "prettifyMarkdown": False,
        },
    )
    if response.status_code != 200:
        raise RuntimeError(
            f"{pdf_path.name} [{start_page}, {end_page}) layout-parsing 失败："
            f"{response.status_code} {response.text}"
        )

    layout_results = response.json()["result"]["layoutParsingResults"]
    if len(layout_results) != expected_pages:
        raise RuntimeError(
            f"{pdf_path.name} [{start_page}, {end_page}) 应返回 {expected_pages} 页，"
            f"layout-parsing 实际返回 {len(layout_results)} 页"
        )

    save_layout_outputs(layout_results, pdf_path, start_page)

    pages = [
        {
            "prunedResult": result["prunedResult"],
            "markdownImages": result["markdown"].get("images"),
        }
        for result in layout_results
    ]
    response = await client.post(
        f"{BASE_URL}/restructure-pages",
        json={"pages": pages, "concatenatePages": False},
    )
    if response.status_code != 200:
        raise RuntimeError(
            f"{pdf_path.name} [{start_page}, {end_page}) restructure-pages 失败："
            f"{response.status_code} {response.text}"
        )

    results = response.json()["result"]["layoutParsingResults"]
    # 暂不回退为逐页调用 /restructure-pages；批量返回异常时直接报错。
    if len(results) != expected_pages:
        raise RuntimeError(
            f"{pdf_path.name} [{start_page}, {end_page}) 使用 concatenatePages=False "
            f"应返回 {expected_pages} 页，实际返回 {len(results)} 页"
        )

    output_dir = OUTPUT_DIR / pdf_path.stem
    output_dir.mkdir(parents=True, exist_ok=True)
    for page_number, result in enumerate(results, start=start_page):
        output_path = output_dir / f"page_{page_number}.md"
        output_path.write_text(
            markdown_with_metadata(result, pdf_path.name, task["metadata"]),
            encoding="utf-8",
        )
        print(f"已保存：{output_path}")


async def main(config):
    tasks = load_tasks()
    unknown = find_unknown_filters(tasks, config)
    if unknown:
        if config["dry_run"]:
            print_unknown_filters(unknown)
        elif not confirm_unknown_filters(unknown):
            print("已取消转换")
            return
    included, excluded = filter_tasks(tasks, config)
    print_summary(included, excluded, config)
    if config["dry_run"]:
        return

    async with httpx.AsyncClient(timeout=None) as client:
        for task, _ in included:
            await transfer_task(task, client)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="按 CSV 页码范围逐页转换 PDF")
    parser.add_argument("--exclude-structure", nargs="*", default=None)
    parser.add_argument("--exclude-scene", nargs="*", default=None)
    parser.add_argument("--exclude-branch", nargs="*", default=None)
    parser.add_argument("--exclude-pair", nargs="*", default=None, metavar="场景-子分支")
    parser.add_argument(
        "--dry-run",
        action=argparse.BooleanOptionalAction,
        default=None,
        help="只校验 CSV 和 PDF，不调用接口",
    )
    asyncio.run(main(resolve_config(parser.parse_args())))
