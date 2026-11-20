"""Combine API-client exports into the multi-sheet workbook used by AI notebooks."""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

SCHEMAS = {
    "channel": ["channel_id", "channel_title", "description", "channel_url"],
    "uploads": ["video_id", "channel_id", "channel_title", "video_title", "description", "published_at", "duration_seconds", "video_url"],
    "playlists": ["playlist_id", "channel_id", "channel_title", "playlist_title", "description"],
    "podcasts": ["playlist_id", "channel_id", "channel_title", "playlist_title", "description"],
    "playlist_video_membership": ["playlist_id", "video_id", "playlist_title", "video_title", "video_description", "video_published_at"],
}
KEYS = {
    "channel": ["channel_id"],
    "uploads": ["video_id"],
    "playlists": ["playlist_id"],
    "podcasts": ["playlist_id"],
    "playlist_video_membership": ["playlist_id", "video_id"],
}


def export_section(path: Path) -> str | None:
    if path.name.startswith("~$"):
        return None
    if path.stem == "playlist_video_membership" or "__playlists__" in path.stem or "__podcasts__" in path.stem:
        return "playlist_video_membership"
    for section in ("channel", "uploads", "playlists", "podcasts"):
        if path.stem.endswith("__" + section):
            return section
    return None


def build_workbook(input_dir: Path, output: Path) -> dict[str, int]:
    if not input_dir.is_dir():
        raise FileNotFoundError(f"API export directory not found: {input_dir}")
    sections: dict[str, list[pd.DataFrame]] = {key: [] for key in SCHEMAS}
    for path in sorted(input_dir.rglob("*.xlsx")):
        if path.resolve() == output.resolve():
            continue
        section = export_section(path)
        if section is None:
            continue
        frame = pd.read_excel(path, keep_default_na=False)
        for column in KEYS[section]:
            if column not in frame:
                raise ValueError(f"{path.name}: missing required identifier {column}")
            frame[column] = frame[column].astype(str).str.strip()
            if frame[column].eq("").any():
                raise ValueError(f"{path.name}: empty identifier in {column}")
        frame["source_file"] = str(path.relative_to(input_dir))
        sections[section].append(frame)
    if not sections["channel"] or not sections["uploads"]:
        raise ValueError("Expected __channel.xlsx and __uploads.xlsx API exports; section-only or legacy yt-dlp files cannot supply the research workbook.")
    output.parent.mkdir(parents=True, exist_ok=True)
    counts = {}
    with pd.ExcelWriter(output, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False, "strings_to_formulas": False}}) as writer:
        for section, frames in sections.items():
            frame = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=SCHEMAS[section])
            frame = frame.reindex(columns=list(dict.fromkeys(SCHEMAS[section] + list(frame.columns))), fill_value="")
            frame = frame.drop_duplicates(subset=KEYS[section], keep="last")
            frame.to_excel(writer, sheet_name=section, index=False)
            counts[section] = len(frame)
    return counts


def main() -> None:
    base = Path(__file__).resolve().parent / "outputs"
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-dir", type=Path, default=base / "youtube")
    parser.add_argument("--output", type=Path, default=base / "youtube_data.xlsx")
    args = parser.parse_args()
    try:
        counts = build_workbook(args.input_dir, args.output)
    except (FileNotFoundError, ValueError) as exc:
        parser.error(str(exc))
    print(f"Saved {args.output}: {counts}")


if __name__ == "__main__":
    main()
