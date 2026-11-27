# Taboo! tools

Generate a Taboo! PowerPoint deck from a JSON card list while keeping the
original deck's design, navigation and slide layout.

## Setup

First, download the code [here](https://github.com/MeltingReactor/Taboo/archive/refs/tags/v1.0.0.zip)

> [!IMPORTANT]
> A template is required. An example template can be found [here](https://github.com/MeltingReactor/Taboo/releases/latest/download/Taboo.pptx)

### Dependences

```powershell
pip install python-pptx tqdm
```

`tqdm` is optional — the generator falls back to a built-in bar if it is not
installed. Requires Python 3.10+.

## Run

```powershell
python taboo_generator.py taboo_sections.json
```

or

```powershell
.\generate.bat
```

Useful overrides:

```powershell
# Keep the exact order in the JSON file
python taboo_generator.py taboo_sections.json --no-shuffle

# Different random order
python taboo_generator.py taboo_sections.json --seed 7

# Different template/output
python taboo_generator.py taboo_sections.json --template "Taboo!.pptx" --output "My deck.pptx"

# Progress bar controls
python taboo_generator.py taboo_sections.json --no-progress
python taboo_generator.py taboo_sections.json --force-progress
python taboo_generator.py taboo_sections.json --fallback-bar
```

The bar uses `tqdm` by default when stderr is a terminal. `--force-progress`
shows it when output is redirected, `--no-progress` hides it, and
`--fallback-bar` forces the dependency-free bar instead of `tqdm`. When a bar
is shown, the terminal is cleared before the final summary is printed.

Paths in the JSON are resolved relative to the JSON file, not the current
folder.

## JSON structure

```json
{
  "settings": {
    "template": "Taboo!.pptx",
    "output": "Taboo!_Generated.pptx",
    "seed": 20261127,
    "show_progress": true,
    "force_progress": false,
    "fallback_bar": false,
    "shuffle_sections": true,
    "shuffle_cards": true,
    "banned_count": 4,
    "number_start": 1,
    "number_box_width": 600000,
    "footer_format": "Taboo • {section}",
    "intro_section_name": "Taboo!",
    "end_section_name": "End",
    "front_slide_indices": [1, 2, 3],
    "final_slide_index": -1,
    "begin_button_text": "Begin!"
  },
  "sections": [
    {
      "name": "General",
      "cards": [
        {
          "word": "Umbrella",
          "banned": ["Rain", "Canopy", "Wet", "Collapsible"],
          "source_word": "Park"
        }
      ]
    }
  ]
}
```

### Card fields

- `word` — the guess word shown in red.
- `banned` — exactly `banned_count` banned words (default 4).
- `source_word` *(optional)* — reuse a slide from the template whose title
  matches this value. Useful for keeping the original cards.
- `source_slide` *(optional)* — reuse a specific 1-based template slide number
  instead of `source_word`.

If neither `source_word` nor `source_slide` is supplied, the generator clones
the template's word-card layout.
