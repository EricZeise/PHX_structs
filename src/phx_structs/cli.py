"""CLI for phx-structs.

    phpp-struct-ref build <filled.xlsx> -o crossref.json --phpp-version EN_10_6_IP
    phpp-struct-ref assemblies <filled.xlsx> --phpp-version EN_10_6_IP
    phpp-struct-ref windows <filled.xlsx> --phpp-version EN_10_6_IP

--phpp-version resolves against PHX_pyxl's own phpp-field-mapping/<version>.md
(the sibling repo's field maps are reused as-is, never copied) -- e.g.
EN_10_6_IP or EN_10_6_SI. --field-map overrides with a direct path.
"""

from __future__ import annotations

import json
from pathlib import Path

import click

from phx_structs.crossref import build_crossref, list_assemblies, list_windows
from phx_structs.sibling_import import PHX_PYXL_SRC

FIELD_MAP_DIR = PHX_PYXL_SRC.parent / "phpp-field-mapping"
DEFAULT_PHPP_VERSION = "EN_10_6_IP"
OUTPUT_DIR = Path(__file__).resolve().parent.parent.parent / "output"


def _resolve_field_map(phpp_version: str, field_map: str | None) -> str:
    if field_map:
        return field_map
    path = FIELD_MAP_DIR / f"{phpp_version}.md"
    if not path.exists():
        raise click.ClickException(
            f"No field map found for --phpp-version {phpp_version!r} "
            f"(expected {path})")
    return str(path)


def _resolve_output_path(output: str) -> Path:
    """A bare filename (no directory component) lands in OUTPUT_DIR, matching
    phpp-shape-sync's convention of always writing generated output to its
    own output/ directory rather than scattering it around the repo root.
    An explicit path (e.g. "subdir/x.json" or "/tmp/x.json") is respected
    as given -- this only supplies a default location, it doesn't override
    a location the caller actually specified.
    """
    path = Path(output)
    return OUTPUT_DIR / path if path.parent == Path(".") else path


@click.group()
@click.version_option(package_name="phx-structs")
def main() -> None:
    """Cross-reference PHPP Components/Windows/Areas by their internal IDs."""


@main.command()
@click.argument("workbook", type=click.Path(exists=True, dir_okay=False))
@click.option("-o", "--output", type=click.Path(dir_okay=False),
              help="Output JSON file path. A bare filename (no directory) lands in output/.")
@click.option("--phpp-version", default=DEFAULT_PHPP_VERSION, show_default=True,
              help=f"PHPP version/variant, resolved to {FIELD_MAP_DIR}/<version>.md.")
@click.option("--field-map", type=click.Path(exists=True, dir_okay=False), default=None,
              help="Path to a specific field map file, overriding --phpp-version.")
def build(workbook: str, output: str | None, phpp_version: str, field_map: str | None) -> None:
    """Build a Components/Windows/Areas cross-reference from a filled PHPP workbook."""
    resolved_map = _resolve_field_map(phpp_version, field_map)
    result = build_crossref(workbook, resolved_map)
    json_str = json.dumps(result, indent=2, default=str)

    if output:
        out_path = _resolve_output_path(output)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json_str, encoding="utf-8")
        n_unresolved = sum(len(v) for v in result["unresolved"].values())
        click.echo(f"Written to {out_path}")
        click.echo(
            f"{len(result['windows'])} windows, {len(result['areas'])} areas, "
            f"{len(result['assemblies'])} assemblies, "
            f"{len(result['thermal_bridges'])} thermal bridges, "
            f"{len(result['hvac']['units'])} vent units, {len(result['hvac']['rooms'])} vent rooms, "
            f"{len(result['hvac']['ducts'])} ducts, "
            f"{n_unresolved} unresolved reference(s)."
        )
    else:
        click.echo(json_str)


@main.command()
@click.argument("workbook", type=click.Path(exists=True, dir_okay=False))
@click.option("-o", "--output", type=click.Path(dir_okay=False),
              help="Output JSON file path (a bare filename lands in output/). Without it, prints a summary table instead.")
@click.option("--phpp-version", default=DEFAULT_PHPP_VERSION, show_default=True,
              help=f"PHPP version/variant, resolved to {FIELD_MAP_DIR}/<version>.md.")
@click.option("--field-map", type=click.Path(exists=True, dir_okay=False), default=None,
              help="Path to a specific field map file, overriding --phpp-version.")
def assemblies(workbook: str, output: str | None, phpp_version: str, field_map: str | None) -> None:
    """List every R-Values assembly with its construction detail and referencing areas grouped together."""
    resolved_map = _resolve_field_map(phpp_version, field_map)
    result = build_crossref(workbook, resolved_map)
    grouped = list_assemblies(result)

    if output:
        out_path = _resolve_output_path(output)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(grouped, indent=2, default=str), encoding="utf-8")
        click.echo(f"Written to {out_path}")
        click.echo(f"{len(grouped)} assemblies.")
        return

    click.echo(f"{len(grouped)} assembl{'y' if len(grouped) == 1 else 'ies'} identified:\n")
    for assembly in grouped:
        layer_desc = " + ".join(layer["sec_1_description"] for layer in assembly["layers"])
        n_areas = len(assembly["areas"])
        click.echo(
            f"  {assembly.get('id')}: {assembly.get('display_name')!r}\n"
            f"      result_val={assembly.get('result_val')!r}\n"
            f"      layers=[{layer_desc}]\n"
            f"      used by {n_areas} area(s), total area={assembly['area_total']!r}\n"
        )


@main.command()
@click.argument("workbook", type=click.Path(exists=True, dir_okay=False))
@click.option("-o", "--output", type=click.Path(dir_okay=False),
              help="Output JSON file path (a bare filename lands in output/). Without it, prints a summary table instead.")
@click.option("--phpp-version", default=DEFAULT_PHPP_VERSION, show_default=True,
              help=f"PHPP version/variant, resolved to {FIELD_MAP_DIR}/<version>.md.")
@click.option("--field-map", type=click.Path(exists=True, dir_okay=False), default=None,
              help="Path to a specific field map file, overriding --phpp-version.")
def windows(workbook: str, output: str | None, phpp_version: str, field_map: str | None) -> None:
    """List every window with its resolved frame and glazing components inlined."""
    resolved_map = _resolve_field_map(phpp_version, field_map)
    result = build_crossref(workbook, resolved_map)
    grouped = list_windows(result)

    if output:
        out_path = _resolve_output_path(output)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(grouped, indent=2, default=str), encoding="utf-8")
        click.echo(f"Written to {out_path}")
        click.echo(f"{len(grouped)} windows.")
        return

    click.echo(f"{len(grouped)} window{'' if len(grouped) == 1 else 's'} identified:\n")
    for window in grouped:
        # COMPONENTS.frames/glazings name this field "description", not
        # "display_name" (unlike assemblies, which do use "display_name").
        frame = window.get("frame") or {}
        glazing = window.get("glazing") or {}
        click.echo(
            f"  row {window.get('_row')}: {window.get('description')!r} (host: {window.get('host')!r})\n"
            f"      frame={frame.get('description')!r}  glazing={glazing.get('description')!r}\n"
            f"      u_w={window.get('u_w')!r}  window_area={window.get('window_area')!r}\n"
        )


if __name__ == "__main__":
    main()
