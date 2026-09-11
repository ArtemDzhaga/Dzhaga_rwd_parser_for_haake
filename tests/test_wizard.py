from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import patch

from haake_rheo.wizard import Prompts, run, select_workflow, custom_import, custom_plot


class WizardTests(TestCase):
    def test_legacy_arguments_do_not_trigger_prompts(self):
        with patch('builtins.input', side_effect=AssertionError('Unexpected prompt')):
            self.assertIsNone(select_workflow(['template.xlsx', 'structure.txt', 'data.rwd'], 'import'))

    def test_selection_order_and_defaults(self):
        with patch('builtins.input', side_effect=['custom', 'B_draft', 'en']):
            self.assertEqual(select_workflow([], 'import'), ('custom', 'B_draft', 'en'))
        with patch('builtins.input', return_value=''):
            self.assertEqual(select_workflow([], 'plot'), ('auto', None, 'ru'))

    def test_invalid_choice_retries(self):
        with patch('builtins.input', side_effect=['unknown', 'k_draft']):
            self.assertEqual(Prompts('en').choice('Подход', 'Profile', ['B_draft', 'K_draft'], 'K_draft'), 'K_draft')

    def test_pasted_and_dragged_paths_with_spaces(self):
        with TemporaryDirectory() as d:
            path = Path(d) / 'my report.xlsx'
            path.touch()
            for text in [str(path), f"'{path}'", str(path).replace(' ', r'\ ')]:
                with patch('builtins.input', return_value=text):
                    self.assertEqual(Prompts().paths('Путь', 'Path', suffixes={'.xlsx'}), path.resolve())

    def test_eof_and_interrupt_do_not_print_traceback(self):
        for error in [EOFError, KeyboardInterrupt]:
            with patch('builtins.input', side_effect=error):
                self.assertEqual(run('import', []), 130)

    def test_custom_flags_skip_only_initial_questions(self):
        with patch('haake_rheo.wizard.custom_import', return_value=0) as importer:
            self.assertEqual(run('import', ['--mode', 'custom', '--profile', 'K_draft', '--language', 'en']), 0)
            importer.assert_called_once_with('K_draft', 'en')

    def test_k_draft_import_dialog_constructs_valid_legacy_options(self):
        with TemporaryDirectory() as d:
            root = Path(d)
            template, ascii_file, rwd = [root / name for name in ['t.xlsx', 'a.txt', 'data.rwd']]
            for path in [template, ascii_file, rwd]:
                path.touch()
            answers = [str(template), str(ascii_file), str(rwd), '', '', '', str(root/'new.xlsx'), '', '']
            with patch('builtins.input', side_effect=answers), patch('haake_rheo.excel_import.main', return_value=0) as main:
                self.assertEqual(custom_import('K_draft', 'en'), 0)
            from haake_rheo.excel_import import parse_args
            args = parse_args(main.call_args.args[0])
            self.assertTrue(args.recursive)
            self.assertFalse(args.series_folders)
            self.assertEqual(args.sheet, 'input_data')

    def test_k_draft_plot_dialog_constructs_valid_options(self):
        with TemporaryDirectory() as d:
            workbook = Path(d)/'report.xlsx'; workbook.touch()
            answers = [str(workbook), str(Path(d)/'plots'), '', '', 'yes', '', '', '', '', '', '', '']
            with patch('builtins.input', side_effect=answers), patch('haake_rheo.plot.main', return_value=0) as main:
                self.assertEqual(custom_plot('K_draft', 'en'), 0)
            from haake_rheo.plot import parse_args
            args = parse_args(main.call_args.args[0])
            self.assertTrue(args.no_fit)
            self.assertTrue(args.split_by_series)
            self.assertEqual(args.language, 'en')
