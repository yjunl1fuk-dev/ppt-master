# Defense PPT Academic Restyle / 答辩 PPT 高级学术风格优化

This folder contains a non-destructive restyling helper for an existing thesis-defense `.pptx`. It is designed for decks like the provided long preview image: many dense scientific slides, a strong dark-blue header/footer system, gold highlights, experimental figures, equations, and chart-heavy pages.

## Design direction

- **Keep content unchanged**: no text rewriting, no figure replacement, no chart/data edits.
- **Make it more academic**: replace generic saturated blue with deep ink/navy, warm paper backgrounds, restrained gold accents, and cyan scientific highlights.
- **Avoid “template sameness”**: the script cycles several page rhythms instead of applying one identical card layout to every page.
- **Improve readability**: lighten dense slide backgrounds, normalize text color, and add subtle visual hierarchy around title/section areas.

## Usage

Install the project dependencies first if your environment does not already include `python-pptx`:

```bash
pip install -r requirements.txt
```

Then run:

```bash
python3 projects/defense_academic_restyle/restyle_pptx.py input.pptx output_academic_restyle.pptx
```

Optional: use a stronger dark academic mode for title/section-heavy decks:

```bash
python3 projects/defense_academic_restyle/restyle_pptx.py input.pptx output_academic_restyle.pptx --mode dark
```

## What the script changes

1. Inserts refined slide backgrounds and accent structures behind existing objects.
2. Normalizes text styling to a calmer academic palette.
3. Adds varied, low-noise visual rhythms across slides.
4. Preserves all existing slide objects, text strings, images, tables, and charts.

## After running

Open the generated PPTX and review several figure-heavy slides. Because PowerPoint files vary in stacking order and placeholder structure, a final manual pass may still be useful for slides with very large full-bleed images or unusual masters.
