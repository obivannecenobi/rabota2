# Web Novel Planner

A simple planner for tracking web‑novel releases and statistics.

## Installation

```sh
python -m venv .venv
. .venv/bin/activate  # or .venv\Scripts\activate on Windows
pip install -r requirements.txt
```

The project can also be installed as a package:

```sh
pip install .
```

## Running

```sh
python run.py
```

## Settings

Runtime settings are stored in `data/config.json`. Important keys:

- `theme` – UI theme name;
- `monochrome` and `mono_saturation` – grayscale options;
- `neon*` – neon effect parameters;
- `glass_*` – parameters of the glass effect;
- `header_font`/`text_font` – font families;
- `save_path` – directory for generated data.

## Examples

Legacy Excel examples are kept in the `examples/` directory for reference only.
