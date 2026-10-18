# geo_excel

[Japanese / 日本語: `README_JP.md`](README_JP.md)

A lightweight wrapper plugin that makes the layer-action scripts from
[QGIS-exportExcel](https://github.com/yamamoto-ryuzo/QGIS-exportExcel)
**runnable from the QGIS plugin menu** (built for QGIS Live Harness).

## Usage (as a plugin)

1. Activate the target vector layer (select features for single-record output).
2. Run from the menu `Plugins → geo_excel`:
   - **Export list report (active layer)** — equivalent of `exportList.py`
   - **Export single record (selected feature)** — equivalent of
     `exportSingle.py` (runs per selected feature)
   - **Export single records in bulk (selected or all features)** — equivalent
     of `exportSingleBulk.py`
3. Layer variables (`xlsout_*`) are configured the same way as described in
   the upstream documentation below.

At execution time, `[% @layer_id %]` (active layer) and `[% $id %]` (selected
feature) are substituted automatically. The script bodies are bundled
unmodified from upstream; only Qt6 rewrites (`PyQt5`→`qgis.PyQt`,
`Qt.WaitCursor`→`Qt.CursorShape.WaitCursor`) are applied just before exec.

**Prerequisites**: `win32com` (pywin32) and Excel must be installed.

## AI/external control API (bridge)

Via `plugins['geo_excel'].bridge` (`GeoExcelBridge`), the plugin can be
operated from QGIS Live Driver's `bridge_call` / `describe_plugins`
(MCP `qgis_bridge_call`, REST, AI assistant panel). Methods return JSON
strings:

- `pluginInfo()` / `status()` — self-description / runtime state (active
  layer, selection count, win32com availability, variable configuration)
- `listScripts()` — list of report scripts (id/file/requires selection/
  required layer variables)
- `layerVariables(layer_id)` / `setLayerVariables(vars, layer_id)` —
  get/set `xlsout_*`/`dpi` layer variables
- `run(script, feature_id=None, confirm=False)` — run a report.
  `script` accepts `list`/`single`/`bulk`/filename/index. `single` uses the
  selected feature or `feature_id`. Modal QMessageBoxes inside the scripts
  are auto-replied only during execution; `confirm=True` answers Ok to
  overwrite/continue prompts (default False = abort). Returns the execution
  result (`excel_result`/`output_path`) and collected messages. Blocks the
  GUI thread while Excel is generating.

Example: `qgis_bridge_call pluginId="geo_excel" method="run" args=["list", null, true]`

---

Below is the upstream (QGIS-exportExcel) README.

# QGIS-exportExcel

This repository provides Python scripts for exporting vector layer
attribute information from QGIS as Excel reports.

# Script overview

## exportList.py
- A script for exporting a list report, used as a "layer scope" action on a
  vector layer.
- Specify the template path and output directory via layer variables;
  information for all features is written to an Excel template.
- `##ListInsert::<field name>` inside the template enables repeated output.

### Usage
1. Set the following layer variables on the target layer in QGIS:
   - `xlsout_list_template_path`: path to the Excel template file
   - `xlsout_list_output_path_fixed`: output directory path
   - `xlsout_list_output_path_variable`: output file name or subdirectory
     name (expressions and field references allowed)
2. Write `##ListInsert::<field name>` in the template Excel file.
3. Register `exportList.py` as a layer action, passing `[% @layer_id %]` as
   the action field.
4. Running the action exports the report.

---

## exportSingle.py
- A script for exporting a single-record report, used as a "feature scope"
  action on a vector layer.
- Writes the attribute values of one selected feature into an Excel
  template.
- `##Attach::<field name>` inside the template inserts attribute values.

### Usage
1. Set the following layer variables on the target layer in QGIS:
   - `xlsout_template_path`: path to the Excel template file
   - `xlsout_output_path_fixed`: output directory path
   - `xlsout_output_path_variable`: output file name (expressions and field
     references allowed)
2. Write `##Attach::<field name>` in the template Excel file.
3. Register `exportSingle.py` as a layer action, passing `[% @layer_id %]`
   and `[% $id %]` as action fields.
4. Running the action exports the report.

---

## exportSingleBulk.py
- A script for bulk export of single-record reports, used as a "feature
  scope" action on a vector layer.
- Exports individual Excel files for selected features or all features at
  once.
- Also supports image embedding, theme switching, and scale specification.

### Usage
1. Set the following layer variables on the target layer in QGIS:
   - `xlsout_template_path`: path to the Excel template file
   - `xlsout_output_path_fixed`: output directory path
   - `xlsout_output_path_variable`: output file name (expressions and field
     references allowed)
2. Write `##Attach::<field name>` or
   `##AttachFitImage::<theme name>::<scale>` in the template Excel file.
3. Register `exportSingleBulk.py` as a layer action, passing
   `[% @layer_id %]` and `[% $id %]` as action fields.
4. Running the action bulk-exports reports for selected or all features.

---

# Common usage

1. Set the required layer variables on the target layer in your QGIS
   project.
2. Prepare an Excel template file and write `##Attach::<field name>` or
   `##ListInsert::<field name>` where you want values inserted.
3. Register a script as a layer action and pass the required fields
   (e.g. `[% @layer_id %]`, `[% $id %]`).
4. Running the action creates an Excel report at the specified output
   location.

# Notes

- Excel automation requires a Windows environment and `win32com.client`.
- The output directory must be created beforehand (some scripts create it
  automatically).
- Template and output paths may be absolute or relative, but mind the path
  separators (`\\` recommended).
- Image embedding and map theme switching use QGIS's map theme feature.

# See also

- See the comments at the top of each script and the function comments.
- Depending on the QGIS version or Python environment, the scripts may not
  work.

Initial development: OrbitalNet — https://www.orbitalnet.jp/
2023/02/01 Added support for QDATE and QDATETIME.
