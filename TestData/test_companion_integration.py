# -*- coding: utf-8 -*-
import json
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "Scripts")
if SCRIPTS not in sys.path:
    sys.path.insert(0, SCRIPTS)

from companion_tools import EXTERNAL_FORK_ARGS, find_external_merge_tools, register_current_app
from fork_integration import MERGE_ARGS, install_fork_integration, integration_status


class CompanionIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.old_env = {key: os.environ.get(key) for key in ("LOCALAPPDATA", "MERGE_TOOLS_HUB_DIR", "EXCEL_MERGE_FORK_LAUNCHER_EXE")}
        os.environ["LOCALAPPDATA"] = os.path.join(self.temp.name, "local")
        os.environ["MERGE_TOOLS_HUB_DIR"] = os.path.join(self.temp.name, "hub")
        self.excel = self._file("apps/ExcelMergeFork.exe")
        self.external = self._file("apps/ExternalMergeTools.exe")
        os.environ["EXCEL_MERGE_FORK_LAUNCHER_EXE"] = self.excel
        self.settings = os.path.join(self.temp.name, "settings.json")

    def tearDown(self):
        for key, value in self.old_env.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
        self.temp.cleanup()

    def _file(self, relative):
        path = os.path.join(self.temp.name, relative)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "wb") as stream:
            stream.write(b"tool")
        return path

    def _write_settings(self, value):
        with open(self.settings, "w", encoding="utf-8") as stream:
            json.dump(value, stream)

    def _read_settings(self):
        with open(self.settings, "r", encoding="utf-8") as stream:
            return json.load(stream)

    def test_excel_install_preserves_external_dispatcher(self):
        dispatcher = {"Type": "Custom", "ApplicationPath": self.external, "Arguments": EXTERNAL_FORK_ARGS}
        self._write_settings({
            "MergeTool": dispatcher,
            "ExternalMergeTools": [{"Type": "Custom", "Name": "ExternalMergeTools", "Path": self.external, "Arguments": EXTERNAL_FORK_ARGS, "IsPrimary": True}],
            "ExternalDiffTools": [],
        })
        result = install_fork_integration(self.excel, settings_path=self.settings, backup=False)
        value = self._read_settings()
        self.assertEqual(dispatcher, value["MergeTool"])
        excel_tool = next(item for item in value["ExternalMergeTools"] if item.get("Name") == "ExcelMergeFork")
        self.assertFalse(excel_tool.get("IsPrimary", False))
        self.assertEqual(os.path.abspath(self.excel), value["ExternalDiffTool"]["ApplicationPath"])
        self.assertTrue(result["installed"])
        self.assertTrue(integration_status(self.excel, self.settings)["merge_configured"])

    def test_excel_remains_primary_without_dispatcher(self):
        self._write_settings({"MergeTool": None, "ExternalMergeTools": [], "ExternalDiffTools": []})
        install_fork_integration(self.excel, settings_path=self.settings, backup=False)
        value = self._read_settings()
        self.assertEqual(os.path.abspath(self.excel), value["MergeTool"]["ApplicationPath"])
        self.assertEqual(MERGE_ARGS, value["MergeTool"]["Arguments"])

    def test_shared_registration_finds_external_app(self):
        registry = {
            "schema_version": 1,
            "apps": {"external_merge_tools": {"launcher_path": self.external, "version": "1.0", "updated_at": "now"}},
        }
        os.makedirs(os.environ["MERGE_TOOLS_HUB_DIR"], exist_ok=True)
        with open(os.path.join(os.environ["MERGE_TOOLS_HUB_DIR"], "registry.json"), "w", encoding="utf-8") as stream:
            json.dump(registry, stream)
        self.assertEqual(os.path.abspath(self.external), find_external_merge_tools())
        self.assertTrue(register_current_app("2.81"))


if __name__ == "__main__":
    unittest.main()
