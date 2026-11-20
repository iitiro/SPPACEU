# Release verification

- 24 offline regression tests passed.
- All nine notebooks executed in their default preview mode through Jupyter kernels without credentials.
- Workbook assembly and the input-preparation functions of steps 1–7 were exercised with synthetic API-shaped exports.
- Channel response parsing was exercised with a mocked model response.
- Checks cover duration parsing, export filenames containing dots, literal Excel text, credential loading, and omission of API keys from network-error messages.
- Live YouTube collection, OpenAI model quality, authenticated media downloads, and optional embedding-model downloads were not executed.

## Verification environment

Checked locally on 2026-11-20 with Python 3.14.2 on macOS. Runtime and test dependency versions are recorded in `constraints.txt`. Dependency resolution was also checked for the SPPACEU requirements targeting Python 3.12; the GitHub Actions platform/version matrix has been configured but has not run as part of this local audit.

## Repeat the offline checks

```sh
python -m pip install -r requirements-dev.txt -c constraints.txt
python -m pytest -q
```

No paid API calls were made. Tests use synthetic data and mocked service responses where external inputs are required. No Git commit, release tag, push, or GitHub Release was created by this preparation.
