# svg-to-animation-ppt

**Turn an SVG into a PowerPoint slide where the logo draws itself, stroke by stroke.**

[한국어 README](README.ko.md)

PowerPoint has no "draw this path" animation for shapes, but it does have one for **ink**:
the *Replay* effect that redraws pen strokes. This tool converts your SVG paths into native
PowerPoint ink (InkML), drops it on a slide, and attaches a Replay animation with the
duration and easing you choose. The result is a normal `.pptx` with a normal animation:
sequence it with other effects, copy it into another deck, present it anywhere PowerPoint
runs. No add-in, no macros, nothing to install on the presenting machine.

```
python svg_to_animation_ppt.py logo.svg --duration 3 --ease-in 0.2 --ease-out 0.3
```

Or use the GUI: drop the SVG, reorder the paths by dragging, tune the sliders, click *Done*.

## Features

- **Hand-drawn input**: drop a `.pptx` that contains pen ink (drawn with Apple Pencil in PowerPoint
  for iPad, or any pen in PowerPoint) and the strokes are imported, one path per pen stroke. Keep the
  original drawing rhythm or normalise to uniform speed.
- **Any SVG geometry**: `path`, `rect`, `circle`, `ellipse`, `line`, `polyline`, `polygon`,
  groups with transforms. Illustrator exports work out of the box.
- **Disconnected paths** (multiple `M` commands, compound paths) become separate strokes
  drawn in order.
- **Stroke colour and width** come from the SVG; fill-only shapes are drawn as outlines.
- **Uniform pen speed**: every stroke is resampled by arc length, so long and short strokes
  draw at the same speed.
- **Ease-in / ease-out** (smooth start / smooth end), written into the animation.
- **Drawing order control**: reorder paths and flip stroke direction in the GUI, or pass an
  order from code.
- **Split mode**: one ink object per SVG path, each with its own animation node, so you can
  retime them individually in PowerPoint's Animation Pane.
- **Template support**: append the slide to an existing deck so it inherits the theme.
- **Fallback image**: a PNG is embedded for apps that cannot render ink (Keynote, older
  Office, Google Slides show a static picture).

## Requirements

- Python 3.10+
- PowerPoint 2016 or later (Microsoft 365 recommended) to *play* the animation.
  Verified on PowerPoint for Mac 365.
- macOS is used for the GUI launcher; the CLI and core are cross-platform.

## Install

```bash
git clone https://github.com/tansan7271/svg-to-animation-ppt.git
cd svg-to-animation-ppt
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

## CLI

```bash
.venv/bin/python svg_to_animation_ppt.py logo.svg                      # -> logo.pptx
.venv/bin/python svg_to_animation_ppt.py logo.svg -o out.pptx --width-cm 12 --duration 3
.venv/bin/python svg_to_animation_ppt.py logo.svg --template deck.pptx # append a slide to deck.pptx
```

| Option | Meaning | Default |
| --- | --- | --- |
| `-o, --output` | Output `.pptx` | `<svg name>.pptx` |
| `--template` | Existing deck to append the slide to | new 16:9 blank deck |
| `--width-cm`, `--height-cm` | Logo size on the slide (aspect ratio is kept) | 60 % of slide width, max 80 % of height |
| `--duration` | Total drawing time in seconds | `2` |
| `--ease-in`, `--ease-out` | Fraction of the duration spent accelerating / decelerating, 0 to 1 | `0` (linear) |
| `--ease-mode` | `keyframes` writes the easing curve as 61 keyframes; `attr` uses the `accel`/`decel` attributes | `keyframes` |
| `--stroke-width` | Force pen width in cm | proportional to SVG `stroke-width` |
| `--color` | Force one colour for all strokes (`#RRGGBB`) | SVG colours |
| `--split path` | One ink object per SVG path, played one after another | single ink object |
| `--step` | Sampling distance in SVG units (smaller = more points) | longest side / 600 |
| `--timing` | Per-point timing info written into the ink: `none`, `offset`, `channel` | `channel` |
| `--keep-rhythm` | For `.pptx` ink input: keep the pen's original sample points so the replay reproduces how fast you drew | off (uniform speed) |
| `--keep-png` | Keep the fallback PNG next to the output | off |

From Python:

```python
from pathlib import Path
import svg_to_animation_ppt as core

opt = core.ConvertOptions(duration=3, ease_in=0.2, ease_out=0.3,
                          path_order=[2, 0, 1], reversed_paths={1})
core.convert(Path("logo.svg"), Path("logo.pptx"), opt)
```

### Hand-drawn input from an iPad

1. Open PowerPoint on the iPad, Draw tab, pick a pen, draw with the Apple Pencil on a blank slide, save.
2. Feed that `.pptx` to the tool instead of an SVG:

```bash
.venv/bin/python svg_to_animation_ppt.py drawing.pptx -o out.pptx --duration 4 --keep-rhythm
```

Each pen stroke becomes one entry in the drawing-order list, so you can reorder or reverse them like SVG paths.
Without `--keep-rhythm` the strokes are resampled to a uniform pen speed.

## GUI (macOS)

Double-click `SVG to Animation PPT.app`. It must stay inside the project folder because it launches
`.venv/bin/python app.py` relative to itself. From a terminal: `.venv/bin/python app.py`.

1. Drop an SVG, or a pptx with pen ink, on the window (or click to browse).
2. The preview shows every path with its order number at the start point and a dot at the end.
3. Drag entries in the **Drawing order** list to reorder. Select one to highlight it.
   *Reverse* flips a path's direction, *Reset* restores document order.
4. Adjust duration, ease-in / ease-out (the little graph shows the curve), size, pen width,
   colour, split mode, template deck. For pptx input, *Keep drawing rhythm* toggles between the original
   pen timing and uniform speed.
5. Click **Done**, choose where to save, and the deck opens in PowerPoint.

Errors are logged to `/tmp/svg-to-animation-ppt.log`.

## How it works

A pptx stores pen strokes as an InkML part (`ppt/ink/inkN.xml`) referenced from the slide
by a `p:contentPart` element. The Replay effect is an ordinary timing node
(`presetID="63"`) that animates the `drawProgress` attribute from 0 to 1.
The converter:

1. Parses the SVG with `svgelements`, flattens transforms, and turns every shape into path
   segments.
2. Splits each path at `M` commands and samples every segment by arc length into points.
3. Writes the points as InkML traces (units of 1/1000 cm, first-difference encoded) with one
   brush per colour / width.
4. Adds the ink part, a PNG fallback, and the timing tree to a `python-pptx` slide.

Things that were established by testing and are easy to get wrong:

- PowerPoint splits replay time by **point count**, so all strokes must be sampled uniformly
  by length. A 2-point straight line would otherwise draw very slowly.
- The brush property `fitToCurve` must be off; with sparse points it turns polygons into blobs
  and lets the line overshoot while drawing.
- Both `accel`/`decel` attributes and multi-keyframe `tavLst` are honoured by the Replay
  effect, even though PowerPoint's UI does not expose an easing control for it.
- `timeOffset` on traces and a `T` channel are accepted but not required.

## Limitations

- Ink is strokes only. Fills, gradients, dashes, and arrowheads are not carried over.
  Fill-only shapes are drawn as outlines.
- Easing cannot be changed inside PowerPoint afterwards; regenerate with new values.
- Pen width and colour are per stroke, with a round tip.
- Copying the ink object between decks carries its animation; whether custom easing survives
  a paste has not been verified. Using `--template` to generate directly into the target deck
  is the safe route.

## Project layout

```
svg_to_animation_ppt.py   core + CLI
app.py                    Tkinter GUI
SVG to Animation PPT.app  macOS launcher (shell script bundle, no Python inside)
examples/                 sample SVGs
internal/                 dev-only decks and test outputs (git-ignored)
```

## License

MIT. See [LICENSE](LICENSE).

## Contributing

Issues and pull requests are welcome, especially reports from PowerPoint for Windows and
PowerPoint for the web, which have not been tested.
