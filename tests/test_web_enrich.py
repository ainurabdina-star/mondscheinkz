import html as htmlmod
import unittest

from catalog_agent import web_enrich as we

GLM_25_23_TEXT = """PRO GLM 25-23
Лазерный дальномер
Номер заказа 0.601.072.W00

ТЕХНИЧЕСКИЕ ХАРАКТЕРИСТИКИ
ДОПОЛНИТЕЛЬНЫЕ ДАННЫЕ
Лазерный диод
635 нм, < 1 мВт
Диапазон измерений*
0,150 – 25,000 м
Вес, ок.*
0,09 кг
Класс лазера
2
* Подробнее о погрешности можно прочитать по следующей ссылке: Технические характеристики изделий
PRO GLM 25-23: ДОПОЛНИТЕЛЬНЫЕ СВЕДЕНИЯ
Лазерный дальномер GLM 25-23 Professional - удобное решение..."""


def _make_html(text: str) -> str:
    escaped = htmlmod.escape(text)
    return "<html><body><div>" + escaped.replace("\n", "</div><div>") + "</div></body></html>"


class TestBuildUrl(unittest.TestCase):
    def test_url_uses_placeholder_slug(self):
        self.assertEqual(
            we.build_product_url("0601072W00"),
            "https://www.bosch-professional.com/kz/ru/products/x-0601072W00",
        )


class TestParseProductPage(unittest.TestCase):
    def test_extracts_model_category_weight_description(self):
        html_doc = _make_html(GLM_25_23_TEXT)
        result = we.parse_product_page(html_doc, "0601072W00")
        self.assertIsNotNone(result)
        self.assertEqual(result.model, "GLM 25-23")
        self.assertEqual(result.category, "Лазерный дальномер")
        self.assertAlmostEqual(result.weight_kg, 0.09)
        self.assertIn("Лазерный диод - 635 нм, < 1 мВт", result.short_description)

    def test_handles_literal_less_than_in_content_without_corrupting_pairs(self):
        # "< 1 мВт" в тексте не должен ломать HTML-парсинг соседних строк —
        # это баг, который был исправлен: наивный regex-стриппер считал
        # экранированный '<' началом тега.
        html_doc = _make_html(GLM_25_23_TEXT)
        result = we.parse_product_page(html_doc, "0601072W00")
        self.assertAlmostEqual(result.weight_kg, 0.09)  # а не None из-за сдвига пар

    def test_rejects_when_order_number_does_not_match(self):
        html_doc = _make_html(GLM_25_23_TEXT)
        result = we.parse_product_page(html_doc, "9999999999")
        self.assertIsNone(result)

    def test_rejects_when_no_order_number_present(self):
        result = we.parse_product_page("<html><body>Товар не найден</body></html>", "0601072W00")
        self.assertIsNone(result)

    def test_order_number_regex_does_not_cross_lines(self):
        # регрессия: жадный \s в захвате артикула перескакивал через перенос
        # строки и хватал текст следующего блока.
        html_doc = _make_html("Номер заказа 0.601.072.W00\nСЛЕДУЮЩИЙ БЛОК ТЕКСТА")
        m = we.ORDER_NUMBER_RE.search(we._strip_html(html_doc))
        self.assertIsNotNone(m)
        self.assertEqual(we._normalize_order_number(m.group(1)), "0601072W00")


class FakeResponse:
    def __init__(self, status_code, text):
        self.status_code = status_code
        self.text = text


class FakeSession:
    def __init__(self, responses):
        self.responses = responses  # {url: FakeResponse}
        self.calls = []

    def get(self, url, timeout=None, headers=None):
        self.calls.append(url)
        return self.responses.get(url, FakeResponse(404, "not found"))


class TestFetchProduct(unittest.TestCase):
    def test_returns_none_on_404(self):
        session = FakeSession({})
        result = we.fetch_product("0601072W00", session)
        self.assertIsNone(result)

    def test_returns_parsed_product_on_200(self):
        url = we.build_product_url("0601072W00")
        session = FakeSession({url: FakeResponse(200, _make_html(GLM_25_23_TEXT))})
        result = we.fetch_product("0601072W00", session)
        self.assertIsNotNone(result)
        self.assertEqual(result.model, "GLM 25-23")


class TestEnrichRecords(unittest.TestCase):
    def test_fills_only_empty_fields_never_overwrites(self):
        records = [
            {
                "артикул": "0601072W00",
                "модель": "УЖЕ ЕСТЬ, НЕ ТРОГАТЬ",
                "Наименование": None,
                "Краткое техническое описание": None,
                "ВЕС нетто": None,
                "Код ТНВЭД": None,
            }
        ]

        import catalog_agent.web_enrich as we_module

        original_fetch = we_module.fetch_product
        we_module.fetch_product = lambda article, session, timeout=15.0: we.ScrapedProduct(
            article=article,
            model="ДОЛЖНО БЫТЬ ПРОИГНОРИРОВАНО",
            category="Лазерный дальномер",
            short_description="описание с сайта",
            weight_kg=0.09,
        )
        try:
            report = we_module.enrich_records(records, delay=0)
        finally:
            we_module.fetch_product = original_fetch

        rec = records[0]
        self.assertEqual(rec["модель"], "УЖЕ ЕСТЬ, НЕ ТРОГАТЬ")  # не перезаписано
        self.assertEqual(rec["Наименование"], "Лазерный дальномер")  # дополнено
        self.assertEqual(rec["Краткое техническое описание"], "описание с сайта")
        self.assertEqual(rec["ВЕС нетто"], 0.09)
        self.assertIsNone(rec["Код ТНВЭД"])  # сайт этого не даёт — так и осталось пусто
        self.assertEqual(report["filled"], 1)

    def test_skips_records_where_all_target_fields_already_filled(self):
        records = [
            {
                "артикул": "X",
                "модель": "M",
                "Наименование": "N",
                "Краткое техническое описание": "D",
                "ВЕС нетто": 1.0,
            }
        ]

        import catalog_agent.web_enrich as we_module

        called = []
        original_fetch = we_module.fetch_product
        we_module.fetch_product = lambda *a, **kw: called.append(1)
        try:
            report = we_module.enrich_records(records, delay=0)
        finally:
            we_module.fetch_product = original_fetch

        self.assertEqual(called, [])
        self.assertEqual(report["checked"], 0)


if __name__ == "__main__":
    unittest.main()
