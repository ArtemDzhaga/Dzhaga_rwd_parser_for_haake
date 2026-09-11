from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import patch
from types import SimpleNamespace

from openpyxl import load_workbook
from haake_rheo.b_draft import identity, import_report, read_results, chart_specs, plot_report


class BDraftTests(TestCase):
    def test_identity_does_not_guess_field_or_shaft(self):
        self.assertEqual(identity(Path('PMS400_99_Fe3O4_1_custom_val_control.rwd')), ('Fe3O4', 1, 'Alum', 'control'))
        with self.assertRaisesRegex(RuntimeError, 'Ambiguous'):
            identity(Path('PMS400_96_Fe3O4_4_magnit_val_control.rwd'))

    def build(self, root, language='ru', missing=False):
        rwd = root/'PMS400_99_Fe3O4_1_custom_val_control.rwd'
        ascii_path = root/'structure.txt'
        rwd.touch(); ascii_path.touch()
        structure = SimpleNamespace(headers=['segment_point', 'Eta in Pas', 'T in °C'])
        def extract(record, reference, bindings):
            # Measured temperature intentionally disagrees with nominal temperature.
            rows = []
            for segment in range(1, 6):
                rows.extend([[f'{segment}|1', 999, 99], [f'{segment}|50', 2, 99]])
                if not missing:
                    rows.append([f'{segment}|100', 4, 99])
            return rows
        output = root/'report.xlsx'
        with patch('haake_rheo.b_draft.source.extract_ascii_record', return_value=structure), \
             patch('haake_rheo.b_draft.source.extract_rwd_record', return_value=object()), \
             patch('haake_rheo.b_draft.source.select_ascii_record_for_csv', return_value=(structure, [], 'test')), \
             patch('haake_rheo.b_draft.source.measurement_rows', side_effect=extract):
            import_report([rwd], [ascii_path], True, [20,30,40,50,60], (50,100), output, language)
        return output

    def test_nominal_temperature_and_two_point_mean_survive_save_reopen(self):
        for language in ['ru', 'en']:
            with self.subTest(language=language), TemporaryDirectory() as d:
                path = self.build(Path(d), language)
                workbook, names, values = read_results(path)
                self.assertEqual([v[4] for v in values], [20,30,40,50,60])
                self.assertEqual([v[5] for v in values], [3]*5)
                self.assertIsNone(workbook[names[3]]['A1'].value)
                self.assertIsNone(workbook[names[1]]['G2'].value)
                workbook.close()
                plot_report(path, Path(d)/'plots', ['xlsx'], language)
                plotted = load_workbook(Path(d)/'plots'/f'report_charts_{language}.xlsx')
                self.assertEqual(len(plotted[names[2]]._charts), 2)
                self.assertTrue(all(c.__class__.__name__ == 'ScatterChart' for c in plotted[names[2]]._charts))
                self.assertEqual(plotted[names[2]]._charts[0].series[0].marker.symbol, 'circle')
                plotted.close()

    def test_missing_target_point_stops_before_saving(self):
        with TemporaryDirectory() as d:
            with self.assertRaisesRegex(RuntimeError, 'Missing/invalid Eta'):
                self.build(Path(d), missing=True)
            self.assertFalse((Path(d)/'report.xlsx').exists())

    def test_all_concentrations_and_matched_magnetic_pairs(self):
        rows = []
        for c in range(1, 6):
            for t in [20, 30]:
                rows.extend([('Fe3O4', c, 'Alum', 'control', t, 2.), ('Fe3O4', c, 'Alum', 'magnet', t, 3.)])
        specs = chart_specs(rows, 'ru')
        self.assertEqual(len([s for s in specs if s[0].endswith('_temperature')]), 5)
        magnetic = next(s for s in specs if s[0].endswith('_magnetic_change'))
        self.assertTrue(all(y == 50 for points in magnetic[4].values() for x,y in points))
