import unittest
from html.parser import HTMLParser
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
GUIDES = (
    "DnS_Auto_MCP_사용자_설명서(AI_연동용).html",
    "DnS_Auto_사용자_설명서(직접_실행용).html",
    "DnS_Auto_빠른_시작_가이드(직접_실행용).html",
    "DnS_Auto_통합_설명서.html",
    "DnS_Auto_통합_설명서(빠른_시작).html",
)
LEGACY_NAMES = (
    "MCP_GUIDE.html",
    "GUI_GUIDE.html",
    "GUI_QUICK_START.html",
    "USER_GUIDE.html",
    "QUICK_START.txt",
)


class LinkCollector(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links = []
        self.images = []

    def handle_starttag(self, tag, attrs):
        values = dict(attrs)
        if tag == "a" and values.get("href"):
            self.links.append(values["href"])
        if tag == "img" and values.get("src"):
            self.images.append(values["src"])


class PortableGuideTests(unittest.TestCase):
    def test_every_guide_links_all_five_guides(self):
        for name in GUIDES:
            parser = LinkCollector()
            parser.feed((ROOT / name).read_text(encoding="utf-8"))
            with self.subTest(guide=name):
                self.assertTrue(set(GUIDES).issubset(parser.links))

    def test_legacy_names_are_not_referenced(self):
        for name in GUIDES:
            text = (ROOT / name).read_text(encoding="utf-8")
            with self.subTest(guide=name):
                for legacy in LEGACY_NAMES:
                    self.assertNotIn(legacy, text)

    def test_guide_images_resolve_under_internal_assets(self):
        for name in GUIDES:
            parser = LinkCollector()
            parser.feed((ROOT / name).read_text(encoding="utf-8"))
            for image in parser.images:
                with self.subTest(guide=name, image=image):
                    self.assertTrue(image.startswith("_internal/assets/guide/"))
                    source_image = ROOT / image.removeprefix("_internal/")
                    self.assertTrue(source_image.is_file(), source_image)


if __name__ == "__main__":
    unittest.main()
