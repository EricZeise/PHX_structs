"""CLI for phx-structs.

    phpp-struct-ref build <filled.xlsx> -o crossref.json --phpp-version EN_10_6_IP

--phpp-version resolves against PHX_pyxl's own phpp-field-mapping/<version>.md
(the sibling repo's field maps are reused as-is, never copied) -- e.g.
EN_10_6_IP or EN_10_6_SI. --field-map overrides with a direct path.
"""

from __future__ import annotations

import json
from pathlib import Path

import click

from phx_structs.crossref import build_crossref
from phx_structs.sibling_import import PHX_PYXL_SRC

FIELD_MAP_DIR = PHX_PYXL_SRC.parent / "phpp-field-mapping"
DEFAULT_PHPP_VERSION = "EN_10_6_IP"


def _resolve_field_map(phpp_version: str, field_map: str | None) -> str:
    if field_map:
        return field_map
    path = FIELD_MAP_DIR / f"{phpp_version}.md"
    if not path.exists():
        raise click.ClickException(
            f"No field map found for --phpp-version {phpp_version!r} "
            f"(expected {path})")
    return str(path)


@click.group()
@click.version_option(package_name="phx-structs")
def main() -> None:
    """Cross-reference PHPP Components/Windows/Areas by their internal IDs."""


@main.command()
@click.argument("workbook", type=click.Path(exists=True, dir_okay=False))
@click.option("-o", "--output", type=click.Path(dir_okay=False),
              help="Output JSON file path.")
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
        Path(output).parent.mkdir(parents=True, exist_ok=True)
        Path(output).write_text(json_str, encoding="utf-8")
        n_unresolved = sum(len(v) for v in result["unresolved"].values())
        click.echo(f"Written to {output}")
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


if __name__ == "__main__":
    main()
