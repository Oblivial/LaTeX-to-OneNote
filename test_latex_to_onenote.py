import unittest

from docx import Document
from lxml import etree

from latex_to_onenote import (
    MATH_NAMESPACE,
    _paragraph_to_html,
    _prepare_markdown,
)


class PrepareMarkdownTests(unittest.TestCase):
    def test_wraps_a_formula_without_delimiters(self):
        self.assertEqual(_prepare_markdown('  x^2 + 1  '), '$x^2 + 1$')

    def test_preserves_delimited_formula_in_text(self):
        text = 'The result is $x^2 + 1$ and then $y$.'
        self.assertEqual(_prepare_markdown(text), text)


class ParagraphHtmlTests(unittest.TestCase):
    def test_keeps_text_and_multiple_inline_math_in_order(self):
        paragraph = Document().add_paragraph('The result is ')
        math = etree.fromstring(
            f'<m:oMath xmlns:m="{MATH_NAMESPACE}"><m:r><m:t>x</m:t></m:r></m:oMath>'
        )
        paragraph._p.append(math)
        paragraph.add_run(' and then ')
        second_math = etree.fromstring(
            f'<m:oMath xmlns:m="{MATH_NAMESPACE}"><m:r><m:t>y</m:t></m:r></m:oMath>'
        )
        paragraph._p.append(second_math)
        paragraph.add_run('.')

        result = _paragraph_to_html(paragraph)

        self.assertIn('<p>The result is ', result)
        self.assertIn('</m:oMath><![endif]--> and then ', result)
        self.assertTrue(result.endswith('.</p>'))

    def test_escapes_text_and_wraps_standalone_math_for_onenote(self):
        paragraph = Document().add_paragraph()
        paragraph.add_run('\u200b')
        math = etree.fromstring(
            f'<m:oMath xmlns:m="{MATH_NAMESPACE}"><m:r><m:t>x</m:t></m:r></m:oMath>'
        )
        paragraph._p.append(math)

        result = _paragraph_to_html(paragraph)
        text_paragraph = Document().add_paragraph('A < B & C')

        self.assertIn('<m:oMathPara', result)
        self.assertEqual(_paragraph_to_html(text_paragraph), '<p>A &lt; B &amp; C</p>')


if __name__ == '__main__':
    unittest.main()
