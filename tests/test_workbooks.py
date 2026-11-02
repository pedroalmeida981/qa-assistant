"""A project with more than one workbook (technique-excel.md section 1.1): each run log goes to the one it names.

python -m unittest discover -s tests   (from the root of qa-assistant)
Does not need LibreOffice: never recalculates.
"""

import contextlib
import copy
import hashlib
import io
import json
import os
import sys
import tempfile
import unittest
import unittest.mock

import yaml

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path[:0] = [ROOT, os.path.join(ROOT, "scripts"), os.path.join(ROOT, "tests", "golden")]

import validate_project
import workspace_ci as wc
from build_fixture import build

from adapters.excel import ConfigError, ExcelAdapter, InvalidBatch, workbook_for
from adapters.excel import apply as excel_apply

with open(os.path.join(ROOT, "adapters", "excel", "example.yaml"), encoding="utf-8") as _f:
    CFG = yaml.safe_load(_f)

MOBILE = {"file": "shop-mobile.xlsx", "batches": ["batch_mobile_*"]}


def excel(**workbooks):
    return {**copy.deepcopy(CFG), "file": "shop.xlsx", "workbooks": workbooks}


class Choice(unittest.TestCase):
    """run.workbook, --workbook and the file name agree, or the batch is refused; none of them is the default."""

    def test_nothing_named_goes_to_the_default(self):
        self.assertEqual(workbook_for(excel(mobile=MOBILE), "runs/batch_20261102_01.json", {"results": []}), "default")
        self.assertEqual(workbook_for(excel(), "runs/batch_mobile_01.json"), "default")

    def test_the_file_name_the_run_block_or_the_option_choose_it(self):
        cfg = excel(mobile=MOBILE)
        self.assertEqual(workbook_for(cfg, "runs/batch_mobile_01.json"), "mobile")
        self.assertEqual(workbook_for(cfg, "runs/batch_x.json", {"run": {"workbook": "mobile"}}), "mobile")
        self.assertEqual(workbook_for(cfg, "runs/batch_x.json", None, "mobile"), "mobile")
        self.assertEqual(workbook_for(cfg, "runs/batch_mobile_01.json", {"run": {"workbook": "mobile"}}, "mobile"),
                         "mobile")

    def test_sources_that_disagree_are_refused(self):
        cfg = excel(mobile=MOBILE)
        for path, doc, requested in [("runs/batch_mobile_01.json", {"run": {"workbook": "default"}}, None),
                                     ("runs/batch_mobile_01.json", None, "default"),
                                     ("runs/batch_x.json", {"run": {"workbook": "mobile"}}, "default")]:
            with self.subTest(path=path, doc=doc, requested=requested), self.assertRaises(InvalidBatch) as cm:
                workbook_for(cfg, path, doc, requested)
            self.assertIn("the workbook is ambiguous", str(cm.exception))

    def test_an_unknown_name_or_two_matching_workbooks_are_refused(self):
        with self.assertRaises(InvalidBatch) as cm:
            workbook_for(excel(mobile=MOBILE), "runs/batch_x.json", {"run": {"workbook": "tablet"}})
        self.assertIn("run.workbook 'tablet' is not a workbook of the project (known: default, mobile)", str(cm.exception))
        both = excel(mobile=MOBILE, tablet={"file": "shop-tablet.xlsx", "batches": ["batch_*_01.json"]})
        with self.assertRaises(InvalidBatch) as cm:
            workbook_for(both, "runs/batch_mobile_01.json")
        self.assertIn("matches the batches of more than one workbook: mobile, tablet", str(cm.exception))

    def test_a_malformed_configuration_is_a_config_error(self):
        for bad in [{"default": MOBILE}, {"mobile": {"batches": ["x"]}}, {"mobile": {"file": "a.xlsx", "batches": "x"}},
                    {"mobile": {**MOBILE, "workbooks": {}}}]:
            with self.subTest(bad=bad), self.assertRaises(ConfigError):
                workbook_for(excel(**bad), "runs/batch_x.json")


class Workspace(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.ws = self.tmp.name
        self.write_project(excel(mobile=MOBILE))
        build(os.path.join(self.ws, "shop.xlsx"))
        build(os.path.join(self.ws, "shop-mobile.xlsx"))
        os.makedirs(os.path.join(self.ws, "runs"))

    def tearDown(self):
        self.tmp.cleanup()

    def write_project(self, section):
        with open(os.path.join(self.ws, "project.yaml"), "w", encoding="utf-8") as f:
            yaml.safe_dump({"tracker": "excel", "excel": section}, f, allow_unicode=True)

    def run_log(self, name, status="Fail", **top):
        path = os.path.join(self.ws, "runs", f"{name}.json")
        with open(path, "w", encoding="utf-8") as f:
            json.dump({**top, "results": [{"tc": "TC-SHOP-001", "status": status, "actual_result": "seen on " + name}]},
                      f)
        return path

    def digest(self, name):
        with open(os.path.join(self.ws, name), "rb") as f:
            return hashlib.sha256(f.read()).hexdigest()

    def status(self, name):
        ad = ExcelAdapter.from_project(self.ws, name)
        ad.load()
        return next(str(r[CFG["cases"]["columns"]["status"]]) for rows in ad.case_rows().values() for r in rows
                    if str(r[CFG["cases"]["columns"]["id"]]) == "TC-SHOP-001")

    def apply(self, *args):
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
            code = excel_apply.main([self.ws, *args, "--no-recalc", "--now", "2026-11-02T12:00:00"])
        return code, out.getvalue()


class ApplyToTheNamedWorkbook(Workspace):
    def test_a_run_log_of_the_other_workbook_never_touches_the_default(self):
        default_before = self.digest("shop.xlsx")
        code, out = self.apply(self.run_log("batch_mobile_01"))
        self.assertEqual(code, 0, out)
        self.assertIn("shop-mobile.xlsx (mobile)", out)
        self.assertEqual(self.digest("shop.xlsx"), default_before)
        self.assertEqual(self.status("mobile"), "Fail")
        mobile_after = self.digest("shop-mobile.xlsx")
        code, out = self.apply(self.run_log("batch_mobile_01"))
        self.assertEqual(code, 0, out)
        self.assertIn("Nothing applied", out)
        self.assertEqual(self.digest("shop-mobile.xlsx"), mobile_after)

    def test_the_default_still_takes_every_other_run_log(self):
        mobile_before = self.digest("shop-mobile.xlsx")
        code, out = self.apply(self.run_log("batch_20261102_01"))
        self.assertEqual(code, 0, out)
        self.assertEqual(self.status("default"), "Fail")
        self.assertEqual(self.digest("shop-mobile.xlsx"), mobile_before)

    def test_an_ambiguous_run_log_writes_nothing_anywhere(self):
        before = (self.digest("shop.xlsx"), self.digest("shop-mobile.xlsx"))
        for args in [(self.run_log("batch_mobile_02", run={"workbook": "default"}),),
                     (self.run_log("batch_mobile_03"), "--workbook", "default"),
                     (self.run_log("batch_04", run={"workbook": "tablet"}),)]:
            with self.subTest(args=args):
                code, out = self.apply(*args)
                self.assertEqual(code, 4, out)
                self.assertIn("NOTHING APPLIED", out)
        self.assertEqual((self.digest("shop.xlsx"), self.digest("shop-mobile.xlsx")), before)

    def test_the_other_workbook_can_override_a_key_of_the_default(self):
        self.write_project(excel(mobile={**MOBILE, "defects": {**CFG["defects"], "id_format": "MOB-{n}"}}))
        ad = ExcelAdapter.from_project(self.ws, "mobile")
        self.assertEqual(ad.cfg["defects"]["id_format"], "MOB-{n}")
        self.assertEqual(ExcelAdapter.from_project(self.ws).cfg["defects"]["id_format"], CFG["defects"]["id_format"])
        self.assertTrue(ad.path.endswith("shop-mobile.xlsx"))
        with self.assertRaises(ConfigError):
            ExcelAdapter.from_project(self.ws, "tablet")


class ValidateEveryWorkbook(Workspace):
    def test_a_missing_other_workbook_is_a_problem(self):
        warnings = []
        self.assertEqual(validate_project.other_workbook_problems(self.ws, warnings), [])
        os.remove(os.path.join(self.ws, "shop-mobile.xlsx"))
        problems = validate_project.other_workbook_problems(self.ws, warnings)
        self.assertEqual(len(problems), 1)
        self.assertTrue(problems[0].startswith("excel.workbooks.mobile: "), problems)


class ApplyWorkflowRoutesEachRunLog(Workspace):
    """The workspace's apply workflow runs the same adapter: each run log reaches its own workbook, an ambiguous one
    fails the job and nothing of it is written."""

    def test_through_the_real_adapter(self):
        calls = []

        def real_adapter(workspace, tracker, path):
            out = io.StringIO()
            with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
                code = excel_apply.main([workspace, path, "--no-recalc", "--now", "2026-11-02T12:00:00"])
            calls.append((os.path.basename(path), code))
            return code, out.getvalue().strip()

        default_before = self.digest("shop.xlsx")
        self.run_log("batch_mobile_01")
        self.run_log("batch_mobile_02", status="Pass", run={"workbook": "default"})
        with unittest.mock.patch.object(wc, "run_adapter", real_adapter), \
                unittest.mock.patch.object(wc, "merged_run_logs",
                                           lambda *a: ["runs/batch_mobile_01.json", "runs/batch_mobile_02.json"]), \
                unittest.mock.patch.object(wc.wg, "git_ok"), \
                unittest.mock.patch.object(wc.wg, "git", lambda *a: unittest.mock.Mock(returncode=0)), \
                unittest.mock.patch.object(wc, "slug_of", lambda ws: "owner/demo"), \
                unittest.mock.patch.object(wc, "gh", lambda *a, **k: None), \
                contextlib.redirect_stdout(io.StringIO()) as out:
            code = wc.cmd_apply_merged(self.ws, unittest.mock.Mock(before="a", after="b"))
        self.assertEqual(code, 1, out.getvalue())
        self.assertEqual(calls, [("batch_mobile_01.json", 0), ("batch_mobile_02.json", 4)])
        self.assertIn("`runs/batch_mobile_02.json`: adapter exit 4", out.getvalue())
        self.assertEqual(self.digest("shop.xlsx"), default_before)
        self.assertEqual(self.status("mobile"), "Fail")


class ApplyWorkflowCommitsEveryWorkbook(unittest.TestCase):
    def test_the_files_staged_after_applying(self):
        proj = {"excel": {"file": "plan.xlsx", "workbooks": {"mobile": {"file": "plan-mobile.xlsx"},
                                                             "copy": {"file": "plan.xlsx"}}}}
        self.assertEqual(wc.workbook_files(proj), ["plan.xlsx", "plan-mobile.xlsx"])
        self.assertEqual(wc.workbook_files({"excel": {"file": "plan.xlsx"}}), ["plan.xlsx"])
        self.assertEqual(wc.workbook_files({}), [])


if __name__ == "__main__":
    unittest.main()
