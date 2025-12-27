````md
_⚠️ Potential issue_ | _🟠 Major_

**Pin the Codecov action to a full commit SHA for security.**

The action `codecov/codecov-action@v4` should be pinned to a full-length commit SHA to use it as an immutable release. This prevents potential supply chain attacks if the tag is moved.

<details>
<summary>🔎 Recommended fix: Pin to commit SHA</summary>

Visit the [codecov-action releases page](https://github.com/codecov/codecov-action/releases) to find the commit SHA for v4 and update:

```diff
-      - name: Upload coverage to Codecov
-        if: runner.os == 'Linux' && matrix.python-version == '3.12'
-        uses: codecov/codecov-action@v4
-        with:
-          files: coverage.xml
-          fail_ci_if_error: true
+      - name: Upload coverage to Codecov
+        if: runner.os == 'Linux' && matrix.python-version == '3.12'
+        uses: codecov/codecov-action@125fc84a9a348dbcf27191600683ec096ec9021c  # v4.x.x
+        with:
+          files: coverage.xml
+          fail_ci_if_error: true
```

</details>

<!-- suggestion_start -->

<details>
<summary>📝 Committable suggestion</summary>

> ‼️ **IMPORTANT**
> Carefully review the code before committing. Ensure that it accurately replaces the highlighted code, contains no missing lines, and has no issues with indentation. Thoroughly test & benchmark the code to ensure it meets the requirements.

```suggestion
      - name: Upload coverage to Codecov
        if: runner.os == 'Linux' && matrix.python-version == '3.12'
        uses: codecov/codecov-action@125fc84a9a348dbcf27191600683ec096ec9021c  # v4.x.x
        with:
          files: coverage.xml
          fail_ci_if_error: true
```

</details>

<!-- suggestion_end -->

<details>
<summary>🧰 Tools</summary>

<details>
<summary>🪛 GitHub Check: Codacy Static Code Analysis</summary>

[warning] 43-43: .github/workflows/pytest.yml#L43
An action sourced from a third-party repository on GitHub is not pinned to a full length commit SHA. Pinning an action to a full length commit SHA is currently the only way to use an action as an immutable release.

</details>

</details>

<details>
<summary>🤖 Prompt for AI Agents</summary>

```
.github/workflows/pytest.yml around lines 41 to 46: the workflow currently uses
the mutable tag codecov/codecov-action@v4 which should be pinned to an immutable
full commit SHA; update the uses line to reference the specific full-length
commit SHA corresponding to v4 from the codecov-action releases (e.g.,
codecov/codecov-action@<full-commit-sha>) so the action is pinned, commit the
change, and verify the workflow runs successfully.
```

</details>

<!-- fingerprinting:phantom:poseidon:puma -->

<!-- This is an auto-generated comment by CodeRabbit -->
````

````md
_⚠️ Potential issue_ | _🟡 Minor_

**Duplicate requirement ID: `[MOD-01]` appears twice.**

Line 202 defines `[MOD-01] build_workbook_data...` but `[MOD-01]` is already used in section 3 (line 97) for "すべてのモデルは `BaseModel` 継承". Consider renaming these to unique IDs, e.g., `[BLD-01]` or `[PIPE-MOD-01]` for the pipeline modeling requirements.

<details>
<summary>🔎 Suggested fix</summary>

```diff
-## 2.6 Pipeline
+## 2.6 Pipeline

 - [PIPE-01] build*pre_com_pipeline は include*\* と mode に応じて必要なステップのみ含む
 ...
-- [MOD-01] build_workbook_data は raw コンテナから WorkbookData/SheetData を構築する
-- [MOD-02] collect_sheet_raw_data は抽出済みデータを raw コンテナにまとめる
+- [PIPE-MOD-01] build_workbook_data は raw コンテナから WorkbookData/SheetData を構築する
+- [PIPE-MOD-02] collect_sheet_raw_data は抽出済みデータを raw コンテナにまとめる
```

</details>

<!-- suggestion_start -->

<details>
<summary>📝 Committable suggestion</summary>

> ‼️ **IMPORTANT**
> Carefully review the code before committing. Ensure that it accurately replaces the highlighted code, contains no missing lines, and has no issues with indentation. Thoroughly test & benchmark the code to ensure it meets the requirements.

```suggestion
- [PIPE-MOD-01] build_workbook_data は raw コンテナから WorkbookData/SheetData を構築する
- [PIPE-MOD-02] collect_sheet_raw_data は抽出済みデータを raw コンテナにまとめる
```

</details>

<!-- suggestion_end -->

<details>
<summary>🤖 Prompt for AI Agents</summary>

```
In docs/agents/TEST_REQUIREMENTS.md around lines 202-203, the requirement ID
[MOD-01] duplicates an earlier use (line 97); rename these duplicate IDs to
unique identifiers (e.g., [BLD-01] or [PIPE-MOD-01]) and update any
cross-references or index entries accordingly so each requirement ID is globally
unique; ensure the new ID naming is consistent with the section naming and run a
quick search to replace all occurrences of the old duplicate if needed.
```

</details>

<!-- fingerprinting:phantom:medusa:ocelot -->

<!-- This is an auto-generated comment by CodeRabbit -->
````

````md
_⚠️ Potential issue_ | _🔴 Critical_

**Same coordinate mismatch as openpyxl_backend.**

The `extract_print_areas` method uses `_parse_print_area_range` which returns zero-based coordinates, but `PrintArea` expects 1-based values (per model definition).

<details>
<summary>🔎 Proposed fix</summary>

```diff
             for part in str(raw).split(","):
                 parsed = _parse_print_area_range(part)
                 if not parsed:
                     continue
                 r1, c1, r2, c2 = parsed
                 areas.setdefault(sheet.name, []).append(
-                    PrintArea(r1=r1, c1=c1, r2=r2, c2=c2)
+                    PrintArea(r1=r1 + 1, c1=c1 + 1, r2=r2 + 1, c2=c2 + 1)
                 )
```

</details>

<details>
<summary>🧰 Tools</summary>

<details>
<summary>🪛 GitHub Check: Codacy Static Code Analysis</summary>

[warning] 37-37: src/exstruct/core/backends/com_backend.py#L37
Try, Except, Continue detected.

</details>

</details>

<details>
<summary>🤖 Prompt for AI Agents</summary>

```
In src/exstruct/core/backends/com_backend.py around lines 17 to 49, the parsed
coordinates from _parse_print_area_range are zero-based while the PrintArea
model expects 1-based values; update the code to convert the parsed r1,c1,r2,c2
to 1-based before constructing PrintArea (e.g., increment each by 1), ensuring
they remain ints and preserving the existing filtering logic for empty/invalid
parses.
```

</details>

<!-- fingerprinting:phantom:medusa:ocelot -->

<!-- This is an auto-generated comment by CodeRabbit -->
````

````md
_⚠️ Potential issue_ | _🔴 Critical_

<details>
<summary>🧩 Analysis chain</summary>

🏁 Script executed:

```shell
# Find PrintArea class definition
find . -type f -name "*.py" | xargs grep -l "class PrintArea"
```

Repository: harumiWeb/exstruct

Length of output: 120

---

🏁 Script executed:

```shell
# Examine the full context of the problematic method
cat -n src/exstruct/core/backends/com_backend.py | sed -n '70,135p'
```

Repository: harumiWeb/exstruct

Length of output: 3486

---

🏁 Script executed:

```shell
# Search for all PrintArea usages to understand coordinate patterns
rg "PrintArea\(" --context=3
```

Repository: harumiWeb/exstruct

Length of output: 5457

---

🏁 Script executed:

```shell
# Get PrintArea class definition
cat -n src/exstruct/models/__init__.py | grep -A 15 "class PrintArea"
```

Repository: harumiWeb/exstruct

Length of output: 1540

---

🏁 Script executed:

```shell
# Check the other extract_print_area method in com_backend for comparison
cat -n src/exstruct/core/backends/com_backend.py | sed -n '36,75p'
```

Repository: harumiWeb/exstruct

Length of output: 1747

---

</details>

**Inconsistent coordinate adjustment: columns are adjusted by -1 but rows are not.**

The `PrintArea` model expects all coordinates to be 1-based (per field descriptions: "Start row (1-based)", "Start column (1-based)", etc.). Line 124 applies `-1` adjustment to `c1` and `c2` but not to `r1` and `r2`, despite both being derived from 1-based Excel COM values. This is inconsistent with `extract_print_area()` (line 47), which passes `PrintArea(r1=r1, c1=c1, r2=r2, c2=c2)` without any adjustments. Remove the `-1` adjustments from both column values:

```diff
                        for j in range(len(cols) - 1):
                            c1, c2 = cols[j], cols[j + 1] - 1
                            results.setdefault(sheet.name, []).append(
-                                PrintArea(r1=r1, c1=c1 - 1, r2=r2, c2=c2 - 1)
+                                PrintArea(r1=r1, c1=c1, r2=r2, c2=c2)
                             )
```

<!-- suggestion_start -->

<details>
<summary>📝 Committable suggestion</summary>

> ‼️ **IMPORTANT**
> Carefully review the code before committing. Ensure that it accurately replaces the highlighted code, contains no missing lines, and has no issues with indentation. Thoroughly test & benchmark the code to ensure it meets the requirements.

```suggestion
                    for i in range(len(rows) - 1):
                        r1, r2 = rows[i], rows[i + 1] - 1
                        for j in range(len(cols) - 1):
                            c1, c2 = cols[j], cols[j + 1] - 1
                            results.setdefault(sheet.name, []).append(
                                PrintArea(r1=r1, c1=c1, r2=r2, c2=c2)
                            )
```

</details>

<!-- suggestion_end -->

<details>
<summary>🤖 Prompt for AI Agents</summary>

```
In src/exstruct/core/backends/com_backend.py around lines 119 to 125, the code
incorrectly subtracts 1 from column coordinates when creating PrintArea
instances while leaving row coordinates unchanged; PrintArea expects 1-based
coordinates and extract_print_area uses unadjusted values, so remove the "-1"
adjustments for c1 and c2 and pass the raw Excel-derived r1, c1, r2, c2 (all
1-based) to PrintArea to make coordinate handling consistent.
```

</details>

<!-- fingerprinting:phantom:medusa:ocelot -->

<!-- This is an auto-generated comment by CodeRabbit -->
````

````md
_⚠️ Potential issue_ | _🔴 Critical_

**Coordinate system mismatch: PrintArea expects 1-based, but zero-based values are passed.**

The `PrintArea` model fields are documented as 1-based (per `src/exstruct/models/__init__.py` lines 95-101), but `_parse_print_area_range` returns zero-based coordinates from `parse_range_zero_based`. This will cause off-by-one errors in downstream consumers.

<details>
<summary>🔎 Proposed fix</summary>

```diff
 def _append_print_areas(areas: PrintAreaData, sheet_name: str, range_str: str) -> None:
     for part in str(range_str).split(","):
         parsed = _parse_print_area_range(part)
         if not parsed:
             continue
         r1, c1, r2, c2 = parsed
-        areas.setdefault(sheet_name, []).append(PrintArea(r1=r1, c1=c1, r2=r2, c2=c2))
+        areas.setdefault(sheet_name, []).append(
+            PrintArea(r1=r1 + 1, c1=c1 + 1, r2=r2 + 1, c2=c2 + 1)
+        )
```

</details>

<!-- fingerprinting:phantom:medusa:ocelot -->

<!-- This is an auto-generated comment by CodeRabbit -->
````

````md
_⚠️ Potential issue_ | _🟡 Minor_

**Misleading error message for generic COM pipeline failure.**

The inner `except` block catches all exceptions from `run_com_pipeline` and `collect_sheet_raw_data`, but the fallback message specifically mentions "Shape extraction failed". This could mislead debugging efforts when the actual failure is in charts, print areas, or data collection.

<details>
<summary>🔎 Suggested fix</summary>

```diff
             except Exception as exc:
                 return _fallback(
-                    f"Shape extraction failed ({exc!r}).",
-                    FallbackReason.SHAPE_EXTRACTION_FAILED,
+                    f"COM pipeline step failed ({exc!r}).",
+                    FallbackReason.COM_PIPELINE_FAILED,
                 )
```

Note: This would require adding `COM_PIPELINE_FAILED` to `FallbackReason` or renaming the existing reason.

</details>

> Committable suggestion skipped: line range outside the PR's diff.

<details>
<summary>🤖 Prompt for AI Agents</summary>

```
In src/exstruct/core/pipeline.py around lines 642-646, the except block
currently returns a fallback that always says "Shape extraction failed" even
though it catches any exception from the COM pipeline steps; update the fallback
to use a new or renamed reason that reflects a generic COM pipeline failure
(e.g., add COM_PIPELINE_FAILED to FallbackReason or rename the existing member)
and change the message to something like "COM pipeline failed ({exc!r})" so
failures in charts, print areas, or data collection are accurately reported;
ensure you update any type/enum definitions and imports accordingly.
```

</details>

<!-- fingerprinting:phantom:medusa:ocelot -->

<!-- This is an auto-generated comment by CodeRabbit -->
````

````md
_⚠️ Potential issue_ | _🟡 Minor_

**Unused `reason` parameter.**

The `reason` parameter is defined and documented but never used in the function body. Either use it (e.g., for logging) or remove it.

<details>
<summary>🔎 Option 1: Use the parameter for logging</summary>

```diff
 def build_cells_tables_workbook(
     *,
     inputs: ExtractionInputs,
     artifacts: ExtractionArtifacts,
     reason: str,
 ) -> WorkbookData:
     """..."""
+    logger.debug("Building fallback workbook: %s", reason)
     backend = OpenpyxlBackend(inputs.file_path)
```

</details>

<details>
<summary>🔎 Option 2: Remove unused parameter</summary>

```diff
 def build_cells_tables_workbook(
     *,
     inputs: ExtractionInputs,
     artifacts: ExtractionArtifacts,
-    reason: str,
 ) -> WorkbookData:
```

And update the call sites accordingly.

</details>

> Committable suggestion skipped: line range outside the PR's diff.

<details>
<summary>🤖 Prompt for AI Agents</summary>

```
In src/exstruct/core/pipeline.py around lines 654 to 669, the function
build_cells_tables_workbook declares and documents a reason parameter that is
never used; either remove the parameter and its doc mention and update all call
sites to stop passing reason, or use reason (recommended) by adding a concise
log statement inside the function (e.g., processLogger.info/debug or
logger.info) that includes the reason when falling back to table_candidates, and
keep the parameter and docs; choose one approach and make matching changes
across callers and tests.
```

</details>

<!-- fingerprinting:phantom:medusa:ocelot -->

<!-- This is an auto-generated comment by CodeRabbit -->
````

````md
_⚠️ Potential issue_ | _🟡 Minor_

<details>
<summary>🧩 Analysis chain</summary>

🏁 Script executed:

```shell
rg "SKIP_COM_TESTS" --type py src/exstruct/core/
```

Repository: harumiWeb/exstruct

Length of output: 467

---

🏁 Script executed:

```shell
rg "SKIP_COM_TESTS" --type py tests/
```

Repository: harumiWeb/exstruct

Length of output: 405

---

🏁 Script executed:

```shell
cat -n src/exstruct/core/pipeline.py | head -100
```

Repository: harumiWeb/exstruct

Length of output: 4075

---

🏁 Script executed:

```shell
rg -B 2 -A 2 "SKIP_COM_TESTS" src/exstruct/core/pipeline.py
```

Repository: harumiWeb/exstruct

Length of output: 733

---

🏁 Script executed:

```shell
cat -n src/exstruct/core/pipeline.py | sed -n '200,300p'
```

Repository: harumiWeb/exstruct

Length of output: 4292

---

🏁 Script executed:

```shell
cat -n tests/test_pipeline.py | sed -n '35,56p'
```

Repository: harumiWeb/exstruct

Length of output: 949

---

🏁 Script executed:

```shell
cat -n src/exstruct/core/pipeline.py | sed -n '140,180p'
```

Repository: harumiWeb/exstruct

Length of output: 1699

---

🏁 Script executed:

```shell
cat -n tests/test_pipeline.py | sed -n '58,90p'
```

Repository: harumiWeb/exstruct

Length of output: 1358

---

🏁 Script executed:

```shell
cat -n tests/test_pipeline.py | sed -n '93,107p'
```

Repository: harumiWeb/exstruct

Length of output: 739

---

</details>

**Remove unnecessary `monkeypatch.delenv("SKIP_COM_TESTS")` calls.**

Lines 19 and 38 remove the `SKIP_COM_TESTS` environment variable, but both calls are unnecessary based on the pipeline logic:

- **Line 19** (test_build_pre_com_pipeline_respects_flags): mode="standard" with include_colors_map=False. The colors_map step requires both `include_colors_map=True` AND `SKIP_COM_TESTS` to be enabled. Since include_colors_map is False, the step is never included regardless of the environment variable.

- **Line 38** (test_build_pre_com_pipeline_includes_colors_map_for_light): mode="light" with include_colors_map=True. The light mode pipeline does not check `SKIP_COM_TESTS` at all—it includes colors_map based only on the flag. The delenv call has no effect.

Remove both lines 19 and 38 to simplify the tests and reduce unnecessary setup.

<details>
<summary>🤖 Prompt for AI Agents</summary>

```
In tests/test_pipeline.py around lines 16 to 32 (and also the related call at
line 38), remove the unnecessary monkeypatch.delenv("SKIP_COM_TESTS",
raising=False) calls: they have no effect on the pipeline branches exercised by
these tests (one uses include_colors_map=False in standard mode, the other is
light mode which doesn't consult SKIP_COM_TESTS), so delete those two lines to
simplify the tests and avoid needless environment manipulation.
```

</details>

<!-- fingerprinting:phantom:poseidon:puma -->

<!-- This is an auto-generated comment by CodeRabbit -->
````
