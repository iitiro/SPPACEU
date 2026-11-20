# SPPACEU — Public Audit of Corruption in Ukraine

Trace public practices of exposing corruption in Ukraine with YouTube data collection, contextual AI coding, and transcript analysis.

**Social Practices of Public Audit of Corruption Exposure in Ukraine**

## Research focus

A collection and analysis workflow for investigating how corruption exposure becomes a public social practice on YouTube. It brings together channel and video metadata, channel and playlist context, relevance screening, and transcript-based interpretation.

## Workflow

| Component | Purpose |
| --- | --- |
| `01 yt_channel_download_cyr_api.py` | Collect channel, upload, and playlist metadata through the YouTube Data API. |
| `02 yt_channel_merge.py` | Normalize and merge channel exports into an Excel dataset. |
| `yt_channel_AI-analysis_step_1.ipynb` | Analyze channel context. |
| `yt_channel_AI-analysis_step_2.ipynb` | Analyze playlist context. |
| `yt_channel_AI-analysis_step_3.ipynb` through `yt_channel_AI-analysis_step_7.ipynb` | Explore successive video-screening variants, including date, relevance, and short-video filters. |
| `yt_channel_AI-analysis_step_8.ipynb` | Download audio for selected videos. |
| `yt_channel_AI-analysis_step_9.ipynb` | Analyze a prepared dataset containing transcripts. |
| `sources_LLM_embeddings.py` and `sources_LLM_search.py` | Compare text sources and search them with sentence embeddings. |

The video-screening notebooks contain related iterations. Review their filters and output filenames before choosing a variant; they are not nine interchangeable batch jobs. Audio downloading and transcript analysis require an intervening transcription step.

## Getting started

Use the downloader's built-in help to inspect its options:

```sh
python3 -m pip install pandas openpyxl requests
python3 "01 yt_channel_download_cyr_api.py" --help
```

Supply a channel list and YouTube API credentials, and set input/output paths for your environment. Review the merger's configured directories before running it.

The analysis notebooks use Google Colab/Drive paths and OpenAI API calls. Configure credentials, workbook sheets, paths, and filters in each selected notebook before execution. The embedding scripts expect text files under `data/txt` and require `sentence-transformers` and `scikit-learn`.

## Scope

Research datasets, channel lists, credentials, and transcript inputs are not bundled as a complete runnable study. AI relevance judgments and duration-based Shorts classifications require validation against the research design.

## License

[MIT](LICENSE).
