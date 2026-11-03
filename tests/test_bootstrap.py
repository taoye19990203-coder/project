"""Regression tests use disposable environments; no network or real audit documents."""
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import unittest
import venv
from bootstrap import ensure_pip, pip_is_working


class PipRepairTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.env = Path(self.tmp.name) / 'venv'
        venv.EnvBuilder(with_pip=False).create(self.env)
        self.python = self.env / ('Scripts/python.exe' if os.name == 'nt' else 'bin/python')

    def test_missing_pip_is_bootstrapped_and_healthy_install_is_unchanged(self):
        self.assertFalse(pip_is_working(self.python))
        self.assertTrue(ensure_pip(self.python, lambda _: None))
        self.assertTrue(pip_is_working(self.python))
        self.assertFalse(ensure_pip(self.python, lambda _: None))

    def test_partial_pip_is_reinstalled_even_with_existing_version_metadata(self):
        ensure_pip(self.python, lambda _: None)
        result = subprocess.check_output([
            str(self.python), '-I', '-c',
            'import pip; from pathlib import Path; print(Path(pip.__file__).parent / "_internal" / "utils")'
        ], text=True).strip()
        damaged = Path(result)
        self.assertTrue(damaged.is_relative_to(self.env))
        shutil.rmtree(damaged)
        failure = subprocess.run([str(self.python), '-I', '-m', 'pip', '--version'], capture_output=True, text=True)
        self.assertNotEqual(failure.returncode, 0)
        self.assertIn('pip._internal.utils', failure.stderr)
        self.assertTrue(ensure_pip(self.python, lambda _: None))
        self.assertTrue(pip_is_working(self.python))
        self.assertTrue(damaged.is_dir())


if __name__ == '__main__':
    unittest.main()
