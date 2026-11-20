from pathlib import Path
import importlib.util
import json
import sys
from types import SimpleNamespace
import pandas as pd
import pytest
import requests
from build_research_workbook import build_workbook
from notebook_runtime import load_openai_key
from notebook_helpers import ROOT, notebook_namespace


def load_client():
    spec = importlib.util.spec_from_file_location("youtube_client", ROOT / "01 yt_channel_download_cyr_api.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(autouse=True)
def no_network(monkeypatch):
    def fail(*args, **kwargs):
        raise AssertionError("Release tests must not contact external services")
    monkeypatch.setattr(requests.sessions.Session, "request", fail)


@pytest.fixture
def research_workbook(tmp_path):
    source = tmp_path / "exports"
    source.mkdir()
    tables = {
        "example__channel": [{"channel_id": "UC_TEST", "channel_title": "Example Ukraine", "description": "Публічний аудит корупції", "channel_url": "https://www.youtube.com/@example"}],
        "example__uploads": [{"video_id": "VID_TEST", "channel_id": "UC_TEST", "channel_title": "Example Ukraine", "video_title": "Розслідування корупції в Україні", "description": "Публічний аудит закупівель", "published_at": "2024-06-01T12:00:00Z", "duration_seconds": 900, "video_url": "https://www.youtube.com/watch?v=VID_TEST", "api_section": "videos"}],
        "example__playlists": [{"playlist_id": "PL_TEST", "channel_id": "UC_TEST", "channel_title": "Example Ukraine", "playlist_title": "Investigations", "description": "Public accountability"}],
        "playlist_video_membership": [{"playlist_id": "PL_TEST", "video_id": "VID_TEST", "playlist_title": "Investigations", "video_title": "Розслідування", "video_description": "Публічний аудит", "video_published_at": "2024-06-01T12:00:00Z"}],
    }
    for name, rows in tables.items():
        pd.DataFrame(rows).to_excel(source / f"{name}.xlsx", index=False)
    output = tmp_path / "youtube_data.xlsx"
    counts = build_workbook(source, output)
    assert counts == {"channel": 1, "uploads": 1, "playlists": 1, "podcasts": 0, "playlist_video_membership": 1}
    return output


def test_duration_and_channel_normalization():
    client = load_client()
    assert client.parse_iso8601_duration_seconds("PT1H2M3S") == 3723
    assert client.parse_iso8601_duration_seconds("P1DT1S") == 86401
    assert client.normalize_channel_source("https://www.youtube.com/@example/videos") == "https://www.youtube.com/@example"


def test_api_export_is_literal_text(tmp_path):
    client = load_client()
    path = client.save_rows([{"title": "=1+1", "description": "example"}], tmp_path / "sample", "uploads")
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=False)
    assert wb.active["A2"].data_type == "s"
    wb.close()


def test_environment_credentials(monkeypatch):
    monkeypatch.setenv("OPENAI_API_KEY", "test-credential")
    assert load_openai_key() == "test-credential"
    monkeypatch.delenv("OPENAI_API_KEY")
    with pytest.raises(ValueError, match="OPENAI_API_KEY"):
        load_openai_key()


@pytest.mark.parametrize("step", range(1, 10))
def test_notebook_default_is_offline(step, monkeypatch, tmp_path):
    monkeypatch.chdir(ROOT)
    monkeypatch.setenv("SPPACEU_DATA_DIR", str(tmp_path))
    ns = notebook_namespace(f"yt_channel_AI-analysis_step_{step}.ipynb", as_main=True)
    assert ns["RUN_ANALYSIS"] is False
    assert not list(tmp_path.glob("*.xlsx"))


@pytest.mark.parametrize("step", range(1, 8))
def test_notebook_inputs_match_api_workbook(step, research_workbook, monkeypatch):
    monkeypatch.chdir(ROOT)
    ns = notebook_namespace(f"yt_channel_AI-analysis_step_{step}.ipynb")
    ns["INPUT_XLSX"] = research_workbook
    ns["OUTPUT_XLSX"] = research_workbook.parent / f"out{step}.xlsx"
    ns["LOG_XLSX"] = research_workbook.parent / f"log{step}.xlsx"
    ns["DIAGNOSTICS_XLSX"] = research_workbook.parent / f"diag{step}.xlsx"
    ns["CHANNEL_AI_XLSX"] = research_workbook.parent / "missing-channel.xlsx"
    ns["PLAYLIST_AI_XLSX"] = research_workbook.parent / "missing-playlist.xlsx"
    if step == 1:
        frame = pd.read_excel(research_workbook, sheet_name=ns["SHEET_CHANNEL"])
        ns["validate_input"](frame)
    elif step == 2:
        frame = ns["prepare_data"]()
    else:
        frame = ns["prepare_selection_data"]()
    assert len(frame) == 1


def test_mocked_channel_response():
    ns = notebook_namespace("yt_channel_AI-analysis_step_1.ipynb")
    class Completions:
        def create(self, **kwargs):
            return SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(content=json.dumps({"channel_topic": "public audit"})))])
    client = SimpleNamespace(chat=SimpleNamespace(completions=Completions()))
    result = ns["call_model"](client, pd.Series({"channel_title": "Example", "description": "Public audit"}))
    assert result["channel_topic"] == "public audit"


def test_empty_source_collection_does_not_load_embedding_model(tmp_path):
    for script in ["sources_LLM_search.py", "sources_LLM_embeddings.py"]:
        import subprocess
        result = subprocess.run([sys.executable, str(ROOT / script), "--data-dir", str(tmp_path)], capture_output=True, text=True)
        assert result.returncode == 2
        assert "No nonempty UTF-8" in result.stderr


def test_channel_name_with_dot_preserves_export_section(tmp_path):
    client = load_client()
    path = client.save_rows([{"video_id": "example"}], tmp_path / "channel.name__uploads", "uploads")
    assert path.name == "channel.name__uploads.xlsx"
    assert (tmp_path / "channel.name__uploads.json").is_file()


def test_published_notebooks_are_clean_and_valid():
    import nbformat
    from notebook_helpers import ROOT
    for path in ROOT.glob("*.ipynb"):
        notebook = nbformat.read(path, as_version=4)
        nbformat.validate(notebook)
        for cell in notebook.cells:
            if cell.cell_type == "code":
                assert not cell.outputs
                assert cell.execution_count is None


def test_network_errors_do_not_expose_api_key(monkeypatch):
    module = load_client()
    client = module.YouTubeDataAPI("private-test-key")
    def connection_error(*args, **kwargs):
        raise requests.ConnectionError("https://example.invalid?key=private-test-key")
    monkeypatch.setattr(client.session, "get", connection_error)
    with pytest.raises(module.YouTubeAPIError) as error:
        client.get("channels", {})
    assert "private-test-key" not in str(error.value)
    assert "ConnectionError" in str(error.value)
