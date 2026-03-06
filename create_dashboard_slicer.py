import argparse
import shutil
from pathlib import Path


def _find_table_in_workbook(workbook, table_name: str):
    for ws in workbook.Worksheets:
        try:
            return ws.ListObjects(table_name)
        except Exception:
            continue
    return None


def _extract_slicer_blueprint(source_wb):
    blueprint = []
    for cache_index in range(1, source_wb.SlicerCaches.Count + 1):
        cache = source_wb.SlicerCaches(cache_index)
        table_name = None
        field_name = None

        try:
            table_name = str(cache.ListObject.Name)
        except Exception:
            table_name = None

        try:
            field_name = str(cache.SourceName)
        except Exception:
            field_name = None

        slicers = []
        for slicer_index in range(1, cache.Slicers.Count + 1):
            slicer = cache.Slicers(slicer_index)
            slicers.append(
                {
                    "name": str(slicer.Name),
                    "caption": str(slicer.Caption),
                    "sheet": str(slicer.Shape.Parent.Name),
                    "left": float(slicer.Left),
                    "top": float(slicer.Top),
                    "width": float(slicer.Width),
                    "height": float(slicer.Height),
                    "columns": int(slicer.NumberOfColumns),
                }
            )

        if table_name and field_name and slicers:
            blueprint.append(
                {
                    "table_name": table_name,
                    "field_name": field_name,
                    "slicers": slicers,
                }
            )
    return blueprint


def _clear_existing_slicers(target_wb):
    for cache_index in range(target_wb.SlicerCaches.Count, 0, -1):
        cache = target_wb.SlicerCaches(cache_index)
        for slicer_index in range(cache.Slicers.Count, 0, -1):
            cache.Slicers(slicer_index).Delete()


def recreate_all_slicers(input_file: str, output_file: str) -> int:
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        raise RuntimeError("This script requires pywin32. Install with: pip install pywin32") from exc

    in_path = Path(input_file).resolve()
    out_path = Path(output_file).resolve()

    if in_path != out_path:
        shutil.copy2(in_path, out_path)

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        source_wb = excel.Workbooks.Open(str(in_path), ReadOnly=True)
        slicer_blueprint = _extract_slicer_blueprint(source_wb)
        source_wb.Close(SaveChanges=False)

        target_wb = excel.Workbooks.Open(str(out_path))
        _clear_existing_slicers(target_wb)

        created = 0
        for cache_def in slicer_blueprint:
            target_table = _find_table_in_workbook(target_wb, cache_def["table_name"])
            if target_table is None:
                continue

            new_cache = target_wb.SlicerCaches.Add2(target_table, cache_def["field_name"])
            for slicer_def in cache_def["slicers"]:
                try:
                    slicer_sheet = target_wb.Worksheets(slicer_def["sheet"])
                except Exception:
                    slicer_sheet = target_wb.Worksheets.Add(After=target_wb.Worksheets(target_wb.Worksheets.Count))
                    slicer_sheet.Name = slicer_def["sheet"]

                new_slicer = new_cache.Slicers.Add(
                    slicer_sheet,
                    Name=slicer_def["name"],
                    Caption=slicer_def["caption"],
                    Left=slicer_def["left"],
                    Top=slicer_def["top"],
                    Width=slicer_def["width"],
                    Height=slicer_def["height"],
                )
                new_slicer.NumberOfColumns = slicer_def["columns"]
                created += 1

        target_wb.Save()
        target_wb.Close(SaveChanges=True)
        return created
    finally:
        excel.Quit()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Recreate all slicers from source workbook into generated workbook (generic)."
    )
    parser.add_argument("--input", required=True, help="Input workbook path")
    parser.add_argument("--output", required=True, help="Output workbook path")
    args = parser.parse_args()

    created = recreate_all_slicers(input_file=args.input, output_file=args.output)
    print(f"Recreated {created} slicer(s) in: {args.output}")


if __name__ == "__main__":
    main()
