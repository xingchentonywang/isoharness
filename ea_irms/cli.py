"""Turn one Isodat export into a correction workbook.

    python -m ea_irms run.xls -o run_corrected.xlsx [--config lab.json]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .parse import ExportFormatError, parse_export
from .pipeline import Config, config_from_dict, config_to_dict, process
from .workbook import VERSION, build_workbook


def main(argv=None):
    ap = argparse.ArgumentParser(prog="ea_irms", description=__doc__)
    ap.add_argument("export", nargs="?", help="Isodat export (.xls, .xlsx or .csv)")
    ap.add_argument("-o", "--out", help="workbook to write (default: alongside the export)")
    ap.add_argument("--config", help="lab configuration JSON")
    ap.add_argument("--write-config", help="write the default configuration and stop")
    ap.add_argument("--no-recalc", action="store_true",
                    help="skip computing the formulas with LibreOffice; the workbook then shows "
                         "values only once Excel opens it")
    ap.add_argument("--version", action="version", version=f"ea_irms {VERSION}")
    args = ap.parse_args(argv)

    if args.write_config:
        Path(args.write_config).write_text(json.dumps(config_to_dict(Config()), indent=2))
        print(f"wrote {args.write_config}")
        return 0

    if not args.export:
        ap.error("an export is required unless you are writing a configuration")

    config = Config()
    if args.config:
        try:
            config = config_from_dict(json.loads(Path(args.config).read_text()))
        except (OSError, ValueError) as exc:
            print(f"cannot read the configuration {args.config}: {exc}", file=sys.stderr)
            return 2
    export = Path(args.export)
    out = Path(args.out) if args.out else export.with_name(export.stem + "_corrected.xlsx")

    try:
        positions = parse_export(export)
    except ExportFormatError as exc:
        print(f"cannot read {export.name}: {exc}", file=sys.stderr)
        return 2

    positions, fits = process(positions, config)
    build_workbook(export, out, config, run_name=export.stem)
    if args.no_recalc:
        baked = False
    else:
        from .workbook import bake_values
        baked = bake_values(out)

    settled = fits["settling"]["settled_from"]
    flagged = sum(1 for p in positions if p.flags)
    print(f"{len(positions)} positions, {flagged} flagged")
    print(f"calibration anchors: {fits['cal_n']['n']}, drift model: {fits['drift_n']['model']}")
    for key, label in (("lin_n", "N linearity"), ("lin_c", "C linearity"),
                       ("drift_n", "N drift"), ("drift_c", "C drift")):
        f = fits[key]["fit"]
        p = "no fit" if not f or f["p"] is None else f"p = {f['p']:.3f}"
        print(f"{label}: {p}, {'applied' if fits[key]['in_force'] else 'not applied'}")
    if fits["checks"]["n"] == 0:
        print("no check standard in this run: nothing independent of the anchors tests the calibration")
    for el, name in (("n", "N"), ("c", "C")):
        if fits[f"drift_{el}"]["outliers"]:
            print(f"{name} drift monitor outlier at position(s) "
                  f"{', '.join(str(r) for r in fits[f'drift_{el}']['outliers'])}")
    if settled > 2:
        print(f"positions 1 to {settled - 1} ran before the mass-28 background settled")
    if fits["carryover"]["flagged"]:
        print("possible carryover from an oversized sample at position(s) "
              + ", ".join(str(r) for r in fits["carryover"]["flagged"]))
    print(f"wrote {out}" + ("" if baked else
          " (formulas only: Excel computes them on opening; a preview pane shows them empty)"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
