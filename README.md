# SPPACEU — Public Audit of Corruption in Ukraine

Trace public practices of exposing corruption in Ukraine with YouTube data collection, contextual AI coding, and transcript analysis.

**Social Practices of Public Audit of Corruption Exposure in Ukraine**

## Setup

Python 3.12 or newer is recommended. Run these commands from the repository directory:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt -c constraints.txt
```

On Windows, activate the environment with `.venv\Scripts\activate`.
`constraints.txt` records the dependency versions used for the local release checks.

## 1. Collect public metadata

Copy `channels.example.txt` to `channels.txt` and replace the example with channel handles, IDs, or URLs, one per line. Set your YouTube Data API key in `YOUTUBE_API_KEY`, then run:

```sh
python "01 yt_channel_download_cyr_api.py" --input channels.txt --output-dir outputs/youtube
```

Alternatively, pass `--api-key-file /path/to/key.txt`. Run with `--help` for pagination-related controls, request delay, channel limits, and playlist options. Exports contain channel, upload, playlist, and membership data as JSON and Excel files. Possible Shorts are duration-based candidates requiring validation, not an authoritative format label.

## 2. Build the analysis workbook

```sh
python build_research_workbook.py --input-dir outputs/youtube --output outputs/youtube_data.xlsx
```

This preserves the sheets expected by the notebooks: `channel`, `uploads`, `playlists`, `podcasts`, and `playlist_video_membership`. It avoids duplicating separate videos/streams/shorts exports and deduplicates playlist membership by playlist/video pair.

`02 yt_channel_merge.py` remains available for a **flat legacy video table**, with `--input-dir` and `--output` arguments. Its output is not the multi-sheet analysis workbook.

## 3. Review and execute the notebooks

```sh
jupyter lab
```

Open notebooks from this repository. Google Drive mounting and system-package installation are no longer performed automatically. For Colab, upload the repository support file `notebook_runtime.py`, mount Drive yourself, and point the environment variables below to your data directories.

| Notebook | Purpose |
| --- | --- |
| Step 1 | Channel context and relevance. |
| Step 2 | Playlist context using membership data. |
| Steps 3–7 | Alternative video-screening versions; select one appropriate to your research design. Step 7 is the latest included iteration. |
| Step 8 | MP3 audio downloads for selected rows in `records/base.xlsx`, sheet `videos`, columns `video_url`, `task`, and `lang`. |
| Step 9 | Coding of prepared transcripts from `Final_dataset_with_transcripts.xlsx`, sheet `dataset_with_transcripts`. |

All notebooks start with `RUN_ANALYSIS = False`. Review paths and inputs before setting it to `True`. Video-screening variants additionally default to `DRY_RUN_ONLY = True`; retain that setting for selection diagnostics and change it to `False` only when ready for API processing. Each screening variant writes its own output files.

For AI stages, set `OPENAI_API_KEY` and `OPENAI_MODEL` to credentials and a compatible chat-completions model available in your API account. Neither is bundled. Text and metadata are sent to OpenAI during live analysis and requests may incur charges.

Directory configuration is shared through `notebook_runtime.py`:

| Environment variable | Default |
| --- | --- |
| `SPPACEU_DATA_DIR` | `outputs/` alongside the support module |
| `SPPACEU_RECORDS_DIR` | `<SPPACEU_DATA_DIR>/records/` |
| `OPENAI_MODEL` | No default; explicitly select a model before AI execution. |
| `DENO_PATH` | Detected on the executable search path when available. |

Stage 8 requires the **FFmpeg system executable**. Install a supported JavaScript runtime such as Deno if yt-dlp requires challenge solving for your sources. An optional `cookies.txt` may be placed in the records directory; never publish that file. Transcription between stages 8 and 9 is a separate preparation step; this repository does not include a speech-to-text service integration for that step.

## Optional semantic source search

```sh
python -m pip install -r requirements-embeddings.txt
python sources_LLM_search.py --data-dir data/txt --query "публічний аудит"
python sources_LLM_embeddings.py --data-dir data/txt
```

These utilities load models only after finding nonempty source files. The first run downloads model weights. The search utility uses multilingual embeddings; the comparison utility uses `all-MiniLM-L6-v2`. Dependencies and model downloads for this optional path are separate from the offline checks.

## Research scope

Research datasets, channel lists, transcripts, model weights, and credentials are not bundled. AI relevance judgments require researcher validation. Downloaded media and generated workbooks remain under ignored output directories.

## Offline checks

```sh
python -m pip install -r requirements-dev.txt -c constraints.txt
python -m pytest -q
```

The test suite does not require API credentials. GitHub Actions is configured to repeat these checks on Python 3.12, 3.13, and 3.14 after push.

## License

[MIT](LICENSE).
