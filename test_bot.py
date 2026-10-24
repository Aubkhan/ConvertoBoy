import io
import unittest

from openpyxl import Workbook

from bot import (
    convert,
    excel_table_candidates,
    excel_table_correct_index,
    parse_questions,
    questions_from_excel_table,
    read_excel_tables,
)


def workbook_bytes(sheets):
    workbook = Workbook()
    workbook.remove(workbook.active)
    for name, rows in sheets:
        worksheet = workbook.create_sheet(name)
        for row in rows:
            worksheet.append(row)
    stream = io.BytesIO()
    workbook.save(stream)
    return stream.getvalue()


class ExcelParserTests(unittest.TestCase):
    def test_multiple_sheets_and_skips_empty_and_repeated_headers(self):
        data = workbook_bytes([
            ('Russian', [
                ['Вопрос', 'Правильный ответ', 'Неправильный ответ 1', 'Неправильный ответ 2'],
                ['Столица Франции?', 'Париж', 'Лион', 'Марсель'],
                [None, None, None, None],
                ['Вопрос', 'Правильный ответ', 'Неправильный ответ 1', 'Неправильный ответ 2'],
                ['Столица Италии?', 'Рим', 'Милан', 'Неаполь'],
            ]),
            ('Uzbek', [
                ['Savol', 'To’g’ri javob', 'Noto’g’ri javob'],
                ['O‘zbekiston poytaxti?', 'Toshkent', 'Samarqand'],
            ]),
        ])

        tables = read_excel_tables(data)

        self.assertEqual([table['sheet'] for table in tables], ['Russian', 'Uzbek'])
        self.assertEqual(excel_table_correct_index(tables[0]), 1)
        self.assertEqual(excel_table_correct_index(tables[1]), 1)
        questions = [
            question
            for table in tables
            for question in questions_from_excel_table(table, excel_table_correct_index(table))
        ]
        self.assertEqual(len(questions), 3)
        self.assertEqual(questions[0], {
            'q': 'Столица Франции?',
            'a': [('+', 'Париж'), ('-', 'Лион'), ('-', 'Марсель')],
        })
        self.assertEqual(questions[2]['a'], [('+', 'Toshkent'), ('-', 'Samarqand')])

    def test_ambiguous_answer_columns_are_left_for_user_selection(self):
        data = workbook_bytes([('Quiz', [
            ['Question', 'Option A', 'Option B', 'Option C'],
            ['Unnumbered question', 'Alpha', 'Beta', 'Gamma'],
        ])])
        table = read_excel_tables(data)[0]

        self.assertIsNone(excel_table_correct_index(table))
        self.assertEqual(excel_table_candidates(table), [1, 2, 3])
        self.assertEqual(
            questions_from_excel_table(table, excel_table_candidates(table)[1]),
            [{'q': 'Unnumbered question', 'a': [('-', 'Alpha'), ('+', 'Beta'), ('-', 'Gamma')]}],
        )

    def test_recognizes_russian_correctness_variants(self):
        data = workbook_bytes([('Quiz', [
            ['Вопрос', 'Верный ответ', 'Ошибочный ответ'],
            ['Без нумерации', 'Да', 'Нет'],
        ])])
        table = read_excel_tables(data)[0]

        self.assertEqual(excel_table_correct_index(table), 1)
        self.assertEqual(
            questions_from_excel_table(table, 1)[0]['a'],
            [('+', 'Да'), ('-', 'Нет')],
        )

    def test_existing_numbered_text_conversion_is_unchanged(self):
        questions = parse_questions('1. Question?\n+ Correct\n- Incorrect')

        self.assertEqual(
            convert(questions, 'marks'),
            '# Question?\n+ Correct\n- Incorrect\n',
        )


if __name__ == '__main__':
    unittest.main()
