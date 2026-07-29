from unittest import TestCase

import numpy as np

from haake_rheo.plot import fitted_curve


class FittedCurveTests(TestCase):
    def test_reduces_unstable_polynomial_order(self) -> None:
        xs = np.array([0.001, 0.002, 0.003, 0.01, 0.02, 0.03])
        ys = np.array([1000.0, 900.0, 850.0, 800.0, 790.0, 780.0])

        _, curve_ys = fitted_curve(xs, ys, order=3, curve_points=120)

        self.assertGreater(len(curve_ys), 0)
        self.assertLessEqual(float(curve_ys.max()), float(ys.max()) * 2)
        self.assertGreaterEqual(float(curve_ys.min()), float(ys.min()) / 2)

    def test_collapses_duplicate_gamma_before_fitting(self) -> None:
        xs = np.array([0.001, 0.001, 0.01, 0.1])
        ys = np.array([1000.0, 1100.0, 800.0, 500.0])

        curve_xs, curve_ys = fitted_curve(xs, ys, order=3, curve_points=40)

        self.assertEqual(len(curve_xs), 40)
        self.assertEqual(len(curve_ys), 40)
