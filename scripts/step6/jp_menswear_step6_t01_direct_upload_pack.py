from __future__ import annotations

import json
import shutil
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parent
STEP6_DIR = ROOT / "step6_output"
EXPORT_DATE = "2026-03-23"

CLONE_DIR = STEP6_DIR / f"T01_Oakln复刻版_{EXPORT_DATE}"
PACK_DIR = STEP6_DIR / f"T01_直接上架素材包_{EXPORT_DATE}"
PACK_DIR.mkdir(parents=True, exist_ok=True)

MAIN_DIR = PACK_DIR / "01_商品主图_按上传顺序"
DETAIL_DIR = PACK_DIR / "02_详情页三联卖点图"
COLOR_DIR = PACK_DIR / "03_色卡与补充图"
REC_DIR = PACK_DIR / "04_推荐商品图"
for folder in [MAIN_DIR, DETAIL_DIR, COLOR_DIR, REC_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

XLSX_PATH = STEP6_DIR / f"T01_直接上架素材包_{EXPORT_DATE}.xlsx"
MD_PATH = STEP6_DIR / f"T01_直接上架素材包说明_{EXPORT_DATE}.md"
JSON_PATH = STEP6_DIR / f"T01_直接上架素材包_{EXPORT_DATE}.json"

ASSETS = CLONE_DIR / "assets"


def cp(src_name: str, target_dir: Path, target_name: str) -> str:
    src = ASSETS / src_name
    dst = target_dir / target_name
    shutil.copy2(src, dst)
    return str(dst)


main_rows = [
    {
        "上传顺序": 1,
        "用途": "商品主图1",
        "说明": "纯白T正面平铺，作为首图。",
        "文件": cp("hero_white.jpg", MAIN_DIR, "01_商品主图1_白T正面平铺.jpg"),
        "可直接用": "是",
    },
    {
        "上传顺序": 2,
        "用途": "商品主图2",
        "说明": "同款上身裁切图，补足穿着感。",
        "文件": cp("model_crop.jpg", MAIN_DIR, "02_商品主图2_上身展示.jpg"),
        "可直接用": "是",
    },
    {
        "上传顺序": 3,
        "用途": "商品主图3",
        "说明": "挂拍陈列图，增强店铺陈列感。",
        "文件": cp("hero_hanger_clean.jpg", MAIN_DIR, "03_商品主图3_挂拍陈列.jpg"),
        "可直接用": "是",
    },
    {
        "上传顺序": 4,
        "用途": "商品主图4",
        "说明": "领口细节图。",
        "文件": cp("detail_collar.jpg", MAIN_DIR, "04_商品主图4_领口细节.jpg"),
        "可直接用": "是",
    },
    {
        "上传顺序": 5,
        "用途": "商品主图5",
        "说明": "面料折叠图，表现厚实感。",
        "文件": cp("detail_fold.jpg", MAIN_DIR, "05_商品主图5_折叠细节.jpg"),
        "可直接用": "是",
    },
    {
        "上传顺序": 6,
        "用途": "商品主图6",
        "说明": "排挂近景图。",
        "文件": cp("detail_rack.jpg", MAIN_DIR, "06_商品主图6_排挂近景.jpg"),
        "可直接用": "是",
    },
    {
        "上传顺序": 7,
        "用途": "商品主图7",
        "说明": "多色挂拍，适合颜色补充展示。",
        "文件": cp("color_multi_clean.jpg", MAIN_DIR, "07_商品主图7_多色挂拍.jpg"),
        "可直接用": "是",
    },
    {
        "上传顺序": 8,
        "用途": "商品主图8",
        "说明": "灰色平铺，展示拓展色。",
        "文件": cp("color_gray.jpg", MAIN_DIR, "08_商品主图8_灰色平铺.jpg"),
        "可直接用": "是",
    },
]

detail_rows = [
    {
        "模块": "三联卖点1",
        "建议标题": "轮廓干净，基础好搭",
        "建议正文": "重磅棉感让白T更容易撑住轮廓，适合四月单穿与轻外套内搭。",
        "文件": cp("model_crop.jpg", DETAIL_DIR, "01_三联卖点1_上身展示.jpg"),
        "可直接用": "是",
    },
    {
        "模块": "三联卖点2",
        "建议标题": "圆领罗纹，细节更稳",
        "建议正文": "领口与罗纹细节清晰，日常搭配不容易显得单薄。",
        "文件": cp("detail_collar.jpg", DETAIL_DIR, "02_三联卖点2_领口细节.jpg"),
        "可直接用": "是",
    },
    {
        "模块": "三联卖点3",
        "建议标题": "300g重磅，降低透感",
        "建议正文": "通过面料折叠与近景展示，突出厚实感与更安心的穿着体验。",
        "文件": cp("detail_fold.jpg", DETAIL_DIR, "03_三联卖点3_折叠细节.jpg"),
        "可直接用": "是",
    },
]

color_rows = [
    {"颜色": "白色", "文件": cp("hero_white.jpg", COLOR_DIR, "01_色卡_白色.jpg")},
    {"颜色": "黑色", "文件": cp("detail_fold_alt.jpg", COLOR_DIR, "02_色卡_黑色.jpg")},
    {"颜色": "灰色", "文件": cp("color_gray.jpg", COLOR_DIR, "03_色卡_灰色.jpg")},
    {"颜色": "军绿", "文件": cp("color_olive.jpg", COLOR_DIR, "04_色卡_军绿.jpg")},
    {"颜色": "米白", "文件": cp("color_beige.jpg", COLOR_DIR, "05_色卡_米白.jpg")},
    {"颜色": "多色挂拍", "文件": cp("color_multi_clean.jpg", COLOR_DIR, "06_色卡_多色挂拍.jpg")},
]

recommend_rows = [
    {"位置": 1, "文件": cp("hero_white.jpg", REC_DIR, "01_推荐商品_白色.jpg"), "标题": "300g重磅コットン クルーネックTシャツ", "颜色": "ホワイト", "价格": "¥6,980"},
    {"位置": 2, "文件": cp("detail_fold_alt.jpg", REC_DIR, "02_推荐商品_黑色.jpg"), "标题": "300g重磅コットン クルーネックTシャツ", "颜色": "ブラック", "价格": "¥6,980"},
    {"位置": 3, "文件": cp("color_gray.jpg", REC_DIR, "03_推荐商品_灰色.jpg"), "标题": "300g重磅コットン クルーネックTシャツ", "颜色": "グレー", "价格": "¥6,980"},
]

template_rows = [
    {"页面区域": "左侧商品图瀑布流", "所需素材": "8张商品图", "当前覆盖": "已覆盖", "对应目录": str(MAIN_DIR)},
    {"页面区域": "中段三联卖点区", "所需素材": "3张图+3段文案", "当前覆盖": "已覆盖", "对应目录": str(DETAIL_DIR)},
    {"页面区域": "颜色/拓展图", "所需素材": "颜色平铺/多色挂拍", "当前覆盖": "已覆盖", "对应目录": str(COLOR_DIR)},
    {"页面区域": "推荐商品区", "所需素材": "3张推荐图+标题价格", "当前覆盖": "已覆盖", "对应目录": str(REC_DIR)},
    {"页面区域": "真正缺口", "所需素材": "高质量男模同款上身正拍", "当前覆盖": "未覆盖", "对应目录": "现有仅有同款自拍裁切图"},
]

copy_rows = [
    {"字段": "日文标题", "内容": "300g重磅コットン クルーネックTシャツ メンズ 半袖 厚手 透けにくい ベーシック"},
    {"字段": "价格", "内容": "¥6,980"},
    {"字段": "副标题", "内容": "肉感のある300gコットンで、白でも透け感を抑えやすいベーシックTシャツ。"},
    {"字段": "卖点1标题", "内容": "輪郭がきれいに見える、ベーシックな主力白T"},
    {"字段": "卖点1正文", "内容": "重みのあるコットンTで、一枚着でもシルエットが崩れにくい。4月の主力カットソーとして提案しやすい。"},
    {"字段": "卖点2标题", "内容": "クルーネックの細部を整え、首回りがすっきり見える"},
    {"字段": "卖点2正文", "内容": "リブと縫製の見え方を重視し、ベーシックながら安っぽく見えにくいディテール感。"},
    {"字段": "卖点3标题", "内容": "重みのある綿感で、白でも透けにくい見え方へ"},
    {"字段": "卖点3正文", "内容": "300gの厚手コットンをベースに、春の単品着用でも安心感が出やすい構成。"},
    {"字段": "商品说明", "内容": "肉感のある300gコットンを使用したクルーネックTシャツ。ゆったりとしたベーシックなシルエットで、春の立ち上がりから初夏まで長く活躍します。"},
]

with pd.ExcelWriter(XLSX_PATH, engine="openpyxl") as writer:
    pd.DataFrame(template_rows).to_excel(writer, sheet_name="模板槽位覆盖", index=False)
    pd.DataFrame(main_rows).to_excel(writer, sheet_name="主图上传顺序", index=False)
    pd.DataFrame(detail_rows).to_excel(writer, sheet_name="三联卖点图", index=False)
    pd.DataFrame(color_rows).to_excel(writer, sheet_name="色卡与补充图", index=False)
    pd.DataFrame(recommend_rows).to_excel(writer, sheet_name="推荐商品图", index=False)
    pd.DataFrame(copy_rows).to_excel(writer, sheet_name="日文文案", index=False)

summary = {
    "pack_dir": str(PACK_DIR),
    "main_dir": str(MAIN_DIR),
    "detail_dir": str(DETAIL_DIR),
    "color_dir": str(COLOR_DIR),
    "recommend_dir": str(REC_DIR),
    "xlsx": str(XLSX_PATH),
    "main_count": len(main_rows),
    "detail_count": len(detail_rows),
    "color_count": len(color_rows),
    "recommend_count": len(recommend_rows),
    "gap": "现有最大缺口是高质量男模同款上身正拍；其它槽位已能直接落图。",
}
JSON_PATH.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

MD_PATH.write_text(
    "\n".join(
        [
            "# T01 直接上架素材包",
            "",
            f"- 目录：`{PACK_DIR}`",
            "- 目标：不是复刻页面，而是让素材直接匹配你成功商品页的真实槽位。",
            f"- 商品主图：`{len(main_rows)}`张",
            f"- 三联卖点图：`{len(detail_rows)}`张",
            f"- 色卡与补充图：`{len(color_rows)}`张",
            f"- 推荐商品图：`{len(recommend_rows)}`张",
            "- 当前唯一明显缺口：更高质量的男模同款上身图。",
        ]
    ),
    encoding="utf-8",
)
