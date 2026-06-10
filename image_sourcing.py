# 三元结构理论 案例分析引擎 — 图片来源管理器
# Copyright (C) 2026 李政恒 (Li Zhengheng)
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as published
# by the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

"""
零成本版权合规图片来源管理器
v6.1 — 链式降级策略：Wikimedia Commons → 官媒配图 → Unsplash
默认不搜索，仅在用户明确触发时执行。

合同与红线：
- 不爬取个人微博/公众号/自媒体配图（肖像权风险）
- 不使用无法验证许可证的图片
- 所有嵌入图片必须标注 [来源：机构 | 许可证 | 原始链接]
"""

import os
import re
import json
import urllib.request
import urllib.parse
import urllib.error
from datetime import datetime


class ImageSourcer:
    """
    版权合规的图片来源管理器。

    链式降级：
        1. Wikimedia Commons  → CC/PD 许可明确，API 直接返回许可证字段（零成本）
        2. 官媒新闻配图       → 引述性使用，标注来源（需 WebSearch 配合）
        3. Unsplash           → 免费商用许可（需 UNSPLASH_ACCESS_KEY 环境变量）
        4. 以上均无结果       → 跳过，不出图

    Usage:
        sourcer = ImageSourcer(report_dir="/path/to/report")
        results = sourcer.search("煤矿安全事故", limit=3)
        for r in results:
            print(r["local_path"], r["attribution"])
    """

    WIKIMEDIA_API = "https://commons.wikimedia.org/w/api.php"
    UNSPLASH_API  = "https://api.unsplash.com/search/photos"

    # 允许嵌入使用的许可证类型
    ALLOWED_LICENSES = frozenset([
        "cc-by", "cc-by-sa", "cc0", "public domain",
        "cc-by-4.0", "cc-by-sa-4.0",
        "cc-by-3.0", "cc-by-sa-3.0",
        "cc-by-2.5", "cc-by-sa-2.5",
        "cc-by-2.0", "cc-by-sa-2.0",
        "pd-old", "pd-us", "pd-usgov",
    ])

    HTTP_HEADERS = {"User-Agent": "CaseAnalysisEngine/1.0 (research; contact@example.com)"}

    def __init__(self, report_dir):
        self.report_dir = report_dir
        self.images_dir = os.path.join(report_dir, "images")
        os.makedirs(self.images_dir, exist_ok=True)

    # ------------------------------------------------------------------
    #  公开 API
    # ------------------------------------------------------------------

    def search(self, query, limit=3):
        """
        链式降级搜索图片。

        Args:
            query: 搜索关键词（建议用英文或中英混合）
            limit: 最多返回图片数量

        Returns:
            list[dict]: 每项包含 local_path, source, license, source_url, attribution
            空列表表明所有来源均无合规结果
        """
        # Tier 1: Wikimedia Commons（首选，许可证最明确）
        results = self._search_wikimedia(query, limit)
        if results:
            return results

        # Tier 2: Unsplash（需环境变量 UNSPLASH_ACCESS_KEY）
        results = self._search_unsplash(query, limit)
        if results:
            return results

        # Tier 3: 官媒配图需 WebSearch 配合，不在本模块实现
        # 调用方可通过 search_official_media() 单独调用
        return []

    def search_official_media(self, query, limit=3):
        """
        搜索官方媒体配图（预留接口）。
        需要外部提供 URLs 列表 → 本方法下载并标注来源。
        （当前由 AI 通过 WebSearch/WebFetch 在分析阶段完成）

        Args:
            query: 搜索关键词
            limit: 最多返回图片数量

        Returns:
            list[dict]: 同 search() 格式
        """
        # 此处为预留接口，实际搜图逻辑由 AI 在分析阶段通过 WebSearch 完成
        # 本方法仅负责下载返回的 URL 并标注来源
        return []

    def download_and_register(self, img_url, source_name, license_text, source_url="", attribution=""):
        """
        下载单张图片并注册版权信息。
        供外部（如官媒搜图后）直接调用。
        """
        local_path = self._download_image(img_url, re.sub(r'[^a-zA-Z0-9]', '_', source_name)[:20])
        if not local_path:
            return None
        return {
            "local_path":  local_path,
            "source":      source_name,
            "license":     license_text,
            "source_url":  source_url,
            "attribution": attribution or source_name,
        }

    # ------------------------------------------------------------------
    #  Wikimedia Commons
    # ------------------------------------------------------------------

    def _search_wikimedia(self, query, limit=3):
        """搜索 Wikimedia Commons，返回许可证明确的图片"""
        params = urllib.parse.urlencode({
            "action":           "query",
            "format":           "json",
            "generator":        "search",
            "gsrsearch":        query,
            "gsrnamespace":     6,        # File 命名空间
            "gsrlimit":         min(limit * 3, 15),
            "prop":             "imageinfo",
            "iiprop":           "url|extmetadata|size",
            "iiextmetadatafilter": "LicenseShortName|LicenseUrl|Artist|Attribution",
            "iiurlwidth":       800,
        })

        url = f"{self.WIKIMEDIA_API}?{params}"
        try:
            req = urllib.request.Request(url, headers=self.HTTP_HEADERS)
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except Exception as e:
            print(f"  [ImageSourcer] Wikimedia 搜索失败: {e}")
            return []

        pages = data.get("query", {}).get("pages", {})
        results = []

        for page_id, page in pages.items():
            if len(results) >= limit:
                break

            imageinfo = page.get("imageinfo", [{}])[0]
            extmeta   = imageinfo.get("extmetadata", {})

            license_short = (extmeta.get("LicenseShortName", {}) or {}).get("value", "")
            license_url   = (extmeta.get("LicenseUrl", {}) or {}).get("value", "")
            artist        = (extmeta.get("Artist", {}) or {}).get("value", "")
            attribution   = (extmeta.get("Attribution", {}) or {}).get("value", "")

            # 许可证过滤
            license_key = license_short.lower().replace(" ", "-")
            if not any(a in license_key for a in self.ALLOWED_LICENSES):
                continue

            img_url = imageinfo.get("thumburl") or imageinfo.get("url")
            if not img_url:
                continue

            local_path = self._download_image(img_url, f"wm_{page_id[:8]}")
            if local_path:
                results.append({
                    "local_path":  local_path,
                    "source":      "wikimedia",
                    "license":     license_short or "CC BY-SA",
                    "source_url":  imageinfo.get("descriptionurl", ""),
                    "attribution": attribution or artist or "Wikimedia Commons",
                })

        return results[:limit]

    # ------------------------------------------------------------------
    #  Unsplash
    # ------------------------------------------------------------------

    def _search_unsplash(self, query, limit=3):
        """搜索 Unsplash（需环境变量 UNSPLASH_ACCESS_KEY）"""
        access_key = os.environ.get("UNSPLASH_ACCESS_KEY", "")
        if not access_key:
            return []

        params = urllib.parse.urlencode({"query": query, "per_page": limit})
        url = f"{self.UNSPLASH_API}?{params}"

        try:
            req = urllib.request.Request(url, headers={
                **self.HTTP_HEADERS,
                "Authorization": f"Client-ID {access_key}",
            })
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except Exception as e:
            print(f"  [ImageSourcer] Unsplash 搜索失败: {e}")
            return []

        results = []
        for photo in data.get("results", []):
            img_url = photo.get("urls", {}).get("regular", "")
            if not img_url:
                continue
            photographer = photo.get("user", {}).get("name", "Unsplash")
            local_path   = self._download_image(img_url, f"us_{photo.get('id', 'img')[:8]}")
            if local_path:
                results.append({
                    "local_path":  local_path,
                    "source":      "unsplash",
                    "license":     "Unsplash License (free commercial)",
                    "source_url":  photo.get("links", {}).get("html", ""),
                    "attribution": f"Photo by {photographer} on Unsplash",
                })
        return results[:limit]

    # ------------------------------------------------------------------
    #  下载工具
    # ------------------------------------------------------------------

    def _download_image(self, img_url, prefix):
        """下载图片到 images/ 目录，返回绝对路径"""
        try:
            ext = ".jpg"
            lowered = img_url.lower()
            if ".png" in lowered:
                ext = ".png"
            elif ".svg" in lowered:
                ext = ".svg"
            elif ".webp" in lowered:
                ext = ".webp"

            ts       = datetime.now().strftime("%H%M%S%f")[:12]
            filename = f"{prefix}_{ts}{ext}"
            filepath = os.path.join(self.images_dir, filename)

            req = urllib.request.Request(img_url, headers=self.HTTP_HEADERS)
            with urllib.request.urlopen(req, timeout=30) as resp:
                with open(filepath, "wb") as f:
                    f.write(resp.read())

            return os.path.abspath(filepath)

        except Exception as e:
            print(f"  [ImageSourcer] 下载失败 [{img_url[:60]}…]: {e}")
            return None


# ====================================================================
#  模块自测
# ====================================================================
if __name__ == "__main__":
    import tempfile
    test_dir = os.path.join(tempfile.gettempdir(), "image_sourcing_test")
    sourcer = ImageSourcer(test_dir)
    print(f"测试目录: {test_dir}")
    results = sourcer.search("Beijing cityscape", limit=2)
    if results:
        for i, r in enumerate(results):
            print(f"\n  [{i+1}] {r['local_path']}")
            print(f"       来源: {r['source']}")
            print(f"       许可: {r['license']}")
            print(f"       署名: {r['attribution']}")
    else:
        print("  未找到合规图片。")
