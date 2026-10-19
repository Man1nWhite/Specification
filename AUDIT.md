# Repository audit — Генератор ТКП (Приложение 3 «Спецификация оборудования»)

Audit date: 2026-10-19
Scope: `generate_tkp.py`, `Module1.bas`, `tools/sync_assets.py`, `tests/test_tkp.py`,
`.github/workflows/tests.yml`, `README.md`, `requirements-dev.txt`, and the two binary
assets `Gen_TKP.xlsm` / `ТКП_шаблон.dotx`.

Method: full source read, dependency install in a clean venv, test-suite execution,
byte-level inspection of both OOXML binaries, and differential testing — the VBA macro
(executed through `pyopenvba`) and the Python generator were run against **identical**
inputs and their outputs compared row by row.

---

## Verdict

The project is in good shape for its size: dependencies resolve, all 6 tests pass locally,
the sync tool is idempotent, and the generators produce a correct specification document
end to end. The dominant risk is **drift between the two implementations** of the same
specification-building algorithm (`Module1.bas` for Excel/VBA, `generate_tkp.py` for
Python). They are maintained in two languages with no shared table of rules, and they have
already diverged in four places — two of which put wrong data into a customer-facing
document.

One correction to this report, found while implementing the fixes: the opening draft
counted "CI is wired up" as a healthy signal. It is not, and never was — see **H3** below.
The workflow existed and reported a status, but it had never executed a single test.

Severity counts: **3 high, 3 medium, 6 low** (H3 was added during remediation).

---

## Resolution status (updated after the fixes below)

| ID | Finding | Sev | State |
|---|---|---|---|
| H1 | VBA leaks «Итого»/«МАРЖА» rows written into column C | High | **Fixed** |
| H2 | `clean_text()` deletes any value starting with `#` | High | **Fixed** |
| H3 | CI has never executed the test suite | High | **Fixed** |
| M1 | `format_qty()` emits scientific notation | Medium | **Fixed** |
| M2 | Personal data committed in the binary assets | Medium | **Fixed** |
| M3 | `_eval_cell()` has no cycle guard | Medium | **Fixed** |
| L1 | Hard-coded `xl/worksheets/sheet26.xml` | Low | **Fixed** |
| L2 | Boolean cell values handled inconsistently | Low | **Fixed** |
| L3 | Test normalisation truncates fractional quantities | Low | **Fixed** |
| L4 | Lint findings, no linter in CI | Low | **Fixed** |
| L5 | `.gitignore` misses Office lock files | Low | **Fixed** |
| L6 | `block_Defs` violates PEP 8 | Low | **Fixed** |

Every code fix is covered by a regression test that was verified to **fail against the
unfixed code** before being accepted. Test count went from 6 to 11.

---

## High

### H1 — VBA does not filter «Итого»/«МАРЖА» rows in column C; totals can leak into the specification

`Module1.bas:218` inspects **only column B** (the caption):

```vba
ElseIf StrComp(Left$(caption, Len("Итого")), "Итого", vbTextCompare) = 0 Then
    currentBlock = 0
```

The Python generator additionally inspects **column C** (the description) and also
recognises «МАРЖА» — `generate_tkp.py:225-231`:

```python
if (
    caption.lower().startswith("итого")
    or description.lower().startswith("итого")
    or description.lower().startswith("маржа")
):
    currentBlock = 0
```

Consequence: a total or margin row typed into column C with its caption cell (B) left
blank keeps `currentBlock` alive in VBA, so the row is emitted into Appendix 3 as a normal
line item. This directly contradicts the documented rule in `README.md`
(«Цены, себестоимость, наценки, строки «Итого...» и работы в Приложение 3 не переносятся»).

Reproduced (same input, both engines):

| input | VBA output | Python output |
|---|---|---|
| `B5` empty, `C5='Итого по шкафу'`, `D5='ART-ITOGO'`, `G5=5` | emits `2. \| Итого по шкафу \| ART-ITOGO \| шт. \| 5` | row correctly omitted |
| `B5` empty, `C5='МАРЖА по шкафу'`, `D5='ART-MAR'`, `G5=7` | emits `2. \| МАРЖА по шкафу \| ART-MAR \| шт. \| 7` | row correctly omitted |

This is not hypothetical: the shipped workbook already stores totals in column C with an
empty column B (e.g. sheet `1` rows 51–53 hold «Итоговая СЕБЕСТОИМОСТЬ…», «Итоговая
СТОИМОСТЬ…», «МАРЖА» with `B` empty). Those rows happen to survive today only because
row 48 carries an «Итого…» caption that resets `currentBlock` first — the guard works by
accident, not by design.

**Fix:** in `ОбработатьЛист`, read the description cell alongside the caption and apply the
same three-way test the Python side uses. Alternatively, hoist the skip rules into a single
shared table so the two engines cannot drift again.

### H2 — `clean_text()` deletes any value starting with `#`, silently dropping real articles

`generate_tkp.py:94-96`:

```python
if text.startswith("#"):
    return ""
```

The intent is to blank Excel error values (`#REF!`, `#N/A`, `#VALUE!`, …). The
implementation blanks **everything** beginning with `#`. Because `clean_text` is applied to
the article (column D) as well as the description, a legitimate article such as `#A-100`
is emptied out of the generated document.

Reproduced:

| input | VBA | Python |
|---|---|---|
| `C4='Контроллер'`, `D4='#A-100'`, `G4=3` | `Контроллер \| #A-100 \| шт. \| 3` | `Контроллер \|  \| шт. \| 3` ← article lost |

The article column is a primary key for the ЗИП `SUMIF` roll-up, so a silently blanked
article also degrades the spare-parts quantity calculation.

**Fix:** match a closed set of Excel error literals instead of the `#` prefix, e.g.

```python
_EXCEL_ERRORS = frozenset(
    {"#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "#GETTING_DATA"}
)
...
if text in _EXCEL_ERRORS:
    return ""
```

### H3 — CI reported a status but never ran a single test

Found while implementing the fixes above, not in the original pass. Every run of the
`tests` workflow — including the one for the merged PR #4 and the merge commit on `main` —
has failed after a few seconds at the **Set up Python** step, before installing anything.
The runner annotation explains it:

```
No file in /home/runner/work/Specification/Specification matched to
[**/requirements.txt or **/pyproject.toml], make sure you have checked out
the target repository
```

`actions/setup-python` with `cache: pip` hashes a dependency file to key its cache, and its
default glob matches only `requirements.txt` or `pyproject.toml`. This repository ships
`requirements-dev.txt`, which neither pattern matches, so setup failed and every later step
was skipped.

The consequence is larger than the misconfiguration: the test suite was introduced in PR #4
together with CI, so **no test in this repository has ever been executed by CI**. The suite
passes locally, but nothing was gating merges. Both findings above could have reached a
delivered proposal without a red build.

**Fix:** set `cache-dependency-path: requirements-dev.txt` on both jobs. The checkouts and
`set up python` steps were also bumped (`checkout@v4 → v6`, `setup-python@v5 → v6`) because
the runner was warning that the pinned versions target the deprecated Node.js 20 and are
being force-run on Node.js 24.

---

## Medium

### M1 — `format_qty()` emits scientific notation and loses precision

`generate_tkp.py:99-103` uses `%g`, which is capped at **6 significant digits** and falls
back to exponent form:

```python
return f"{qty:g}"
```

Measured output, which flows straight into the Word table:

| value | rendered |
|---|---|
| `2.5` | `2.5` ok |
| `1/3` | `0.333333` — silently rounded |
| `1234567.5` | `1.23457e+06` — **scientific notation in a customer document** |
| `10000000.5` | `1e+07` — **worse** |

**Fix:** `f"{qty:.10g}"` at minimum, or better `format(Decimal(str(qty)).normalize(), "f")`
to guarantee plain decimal notation.

### M2 — Personal data committed in the binary assets

Both binaries are tracked in a public repository and carry identifying metadata:

- `Gen_TKP.xlsm` → `xl/workbook.xml` embeds the author's absolute working path
  `C:\Users\ABagdasaryan\Генераторы\Генератор ТКП\` (via `x15ac:absPath`);
  `docProps/core.xml` has `dc:creator` and `cp:lastModifiedBy` = `ABagdasaryan`.
- `ТКП_шаблон.dotx` → `docProps/core.xml` has `dc:creator` =
  `Шустваль Людмила Александровна` (a full third-party personal name) and
  `cp:lastModifiedBy` = `ABagdasaryan`.

**Fix:** strip `docProps/core.xml` fields and the `absPath` element before committing, or
regenerate the assets from a sanitised source. Note that `tools/sync_assets.py` does not
touch these parts, so a plain re-run will not clean them.

### M3 — `_eval_cell()` has no cycle guard; circular references crash the generator

`generate_tkp.py:137-186` recurses through cross-sheet cell references with no visited-set.
Two cells pointing at each other produce an unhandled `RecursionError`:

```python
b['C1'] = "=Y!C1"; c['C1'] = "=X!C1"
g._eval_cell(b, b, wb, wb, 1, 3)   # -> RecursionError
```

Real workbooks do acquire accidental reference cycles. A crash here is reported as a raw
Python traceback with no context about which cell was being resolved.

**Fix:** thread a `frozenset`/`set` of `(sheet, row, col)` keys through the recursion and
return `""` when a key repeats.

---

## Low

### L1 — Hard-coded worksheet part name `xl/worksheets/sheet26.xml`

`tools/sync_assets.py:235` and `tests/test_tkp.py:87` both address the preview sheet by its
ZIP part name rather than by its sheet name `Спец_ТКП`. The mapping is correct today, but
inserting, deleting, or reordering any sheet before position 26 will make the sync tool
overwrite an unrelated worksheet **and** the test will keep passing while checking the wrong
sheet.

**Fix:** resolve the part via `xl/workbook.xml` → `r:id` → `xl/_rels/workbook.xml.rels`.

### L2 — Boolean cell values handled inconsistently

`clean_text` returns `""` for `bool` (`generate_tkp.py:88-89`), whereas VBA's `ЧистыйТекст`
returns the string `"True"`. Divergence is cosmetic but is a second instance of the drift in
H1/H2.

### L3 — Test normalisation truncates fractional quantities

`tests/test_tkp.py:166` coerces the VBA quantity with `int(v[4])` before comparing against
the Python rows. Every quantity in the current fixtures is a whole number, so this is
invisible today, but it means the VBA↔Python equivalence test cannot see non-integer
quantities at all — it would either falsely fail (`int(2.9)=2` vs `2.9`) or falsely pass.
Given the two engines differ in exactly this area (H1, H2), the comparison should be
stricter, not looser.

### L4 — Lint findings

`ruff check .` reports 11 issues at default settings, none fatal but all trivially fixable:

- `I001` unsorted imports in all three Python files
- `UP035` `Iterable` imported from `typing` instead of `collections.abc`
- `B008` `RGBColor(0x00, 0x00, 0x00)` evaluated as a default argument
  (`generate_tkp.py:467`) — works because `RGBColor` is immutable, but is a trap
- `PYI041` / `RUF046` `float | int` should be `float`; `int(round(x))` is redundant
- `PIE810` `startswith` called twice where a tuple would do (`generate_tkp.py:196`)
- `SIM102` collapsible `if` (`tools/sync_assets.py:206`)
- `EXE001` shebang present but `generate_tkp.py` is not executable

There is no linter or formatter in CI, so none of this is enforced.

### L5 — `.gitignore` misses Office lock files

Only the generated `.docx` output is ignored. Opening `Gen_TKP.xlsm` in Excel alongside a
working copy creates `~$Gen_TKP.xlsm`, which will show up as an untracked binary in
`git status` and is easy to commit by accident. Suggest adding `~$*`, `*.tmp`, `.DS_Store`.

### L6 — Minor style

`generate_tkp.py:293` declares `block_Defs` in mixed case, which reads as a class name and
violates PEP 8 (`block_defs`).

---

## Verified healthy

These were checked rather than assumed:

- **Tests pass.** All 6 tests in `tests/test_tkp.py` pass on Python 3.11 in a clean venv.
- **Dependencies resolve.** `pyOpenVBA==6.5.0` and `lxml==6.0.2` both exist on PyPI;
  `pip install -r requirements-dev.txt` succeeds cleanly.
- **Tests do not corrupt the binary assets.** `test_generation_groups_sums_skips_and_refreshes`
  writes into the workbook and runs macros, which looks like it mutates the tracked
  `Gen_TKP.xlsm`. It does not: `pyopenvba.ExcelApplication` is a pure-Python in-memory VBA
  interpreter and only writes on an explicit `.save()`. `md5sum` of `Gen_TKP.xlsm` and
  `ТКП_шаблон.dotx` is byte-identical before and after a full test run.
- **`generate_tkp.py` does not mutate the source workbook**, confirming the README claim
  about in-memory demo data.
- **`tools/sync_assets.py` is idempotent** — a second run leaves both assets byte-identical.
- **No stale external references.** Neither binary contains an `externalLink` part, and
  there are no defined names pointing outside the workbook.
- **No homoglyph risk.** The Cyrillic keywords the parsers match on
  («Спецификация », «Итого», «НВТ», «Эксара», «ИнСАТ», «3-х поставщиков») were compared
  codepoint-by-codepoint against the captions stored in the workbook. All match; the 61
  `ruff` `RUF001` warnings are legitimate Cyrillic and should be suppressed project-wide
  rather than "fixed".
- **ЗИП `SUMIF` emulation is numerically correct.** The Python re-implementation of
  `ROUNDUP((Σ SUMIF(...))/10, 0)` returns 3 for a cabinet quantity of 25, matching what
  real Excel computes.
- **End-to-end document generation is correct** — 7 sections / 16 positions, correct
  section numbering («1.» … «7.»), per-section item renumbering, bold section headers with
  `Компл.` / `1`, and empty blocks correctly omitted (Шкаф 2 Эксара).

---

## What was changed

**H1 — `Module1.bas`.** `ОбработатьЛист` now reads column C alongside column B and routes
both through a new `ЭтоСлужебнаяСтрока()` helper, so «Итого» in either column and «МАРЖА»
in the description column end the block. The VBA was re-embedded into `Gen_TKP.xlsm` with
the project's own `tools/sync_assets.py`; only `xl/vbaProject.bin` changed, and all 31 VBA
modules survived.

**H2 — both engines.** `generate_tkp.py` gained an `EXCEL_ERRORS` frozen set
(`#REF!`, `#N/A`, `#VALUE!`, `#DIV/0!`, `#NAME?`, `#NULL!`, `#NUM!`, `#SPILL!`, `#CALC!`,
`#GETTING_DATA`) that replaces every `#`-prefix test in `clean_text`,
`_parse_numeric_qty`, `_eval_cell` and the article fallback. `Module1.bas` mirrors it with
`ЭтоОшибкаExcel()`. Articles such as `#A-100` now survive in both engines.

**M1 — `generate_tkp.py`.** `format_qty` renders via `f"{qty:.10f}"` with trailing zeros
stripped, so `1234567.5` stays `1234567.5` instead of becoming `1.23457e+06`.

**M2 — both binaries.** Removed `dc:creator` / `cp:lastModifiedBy` from both
`docProps/core.xml` files and dropped the `mc:AlternateContent` block carrying
`C:\Users\ABagdasaryan\…` from `xl/workbook.xml`. Every other ZIP entry was copied through
byte-for-byte. Both assets still open, and a full re-scan finds no residual personal data.

**M3 — `generate_tkp.py`.** `_eval_cell` threads a `_seen` set of `(sheet, row, col)` keys
through its recursion; a repeat visit returns `""`. Cycles and self-references degrade to
an empty value, while ordinary reference chains still resolve.

**L1/L2/L3/L6 — tests and tooling.** `tools/sync_assets.py` resolves the preview sheet by
name through `xl/workbook.xml` → rels (verified to follow the sheet name when the parts are
reordered). The test reads it by name via openpyxl instead of by part index. The VBA↔Python
comparison now collapses integral floats instead of truncating with `int()`, and
`ЧистыйТекст` drops booleans to match `clean_text`. `block_Defs` renamed to `block_defs`.

**L4/L5.** Added `ruff.toml` pinning the rule selection (Cyrillic ambiguity rules
`RUF001`/`RUF003` are ignored — they are false positives in a Russian-language project),
pinned `ruff==0.17.1` in `requirements-dev.txt`, added a Ruff job to CI, and extended
`.gitignore` with Office lock files, `.ruff_cache/` and OS noise. `ruff check .` is clean.

**H3 — `.github/workflows/tests.yml`.** Added `cache-dependency-path:
requirements-dev.txt` to both jobs, which is what actually unblocks CI, and bumped
`actions/checkout` to v6 and `actions/setup-python` to v6 to clear the Node.js 20
deprecation warning.

Verification: clean-venv install, `ruff check .` clean, 11/11 tests pass, and the generated
document is unchanged in structure (7 sections / 16 positions). The pinned dependency
versions (`ruff==0.17.1`, `pyOpenVBA==6.5.0`, `lxml==6.0.2`) were confirmed to exist on
PyPI, since CI resolves them there rather than against a local index.

---

## Still open (structural, not fixed here)

The block/skip rules (caption → block number, fallback row ranges, «Итого»/«МАРЖА»
exclusions, quantity formatting) still exist **twice**, in two languages. They are now
aligned and pinned by differential tests, so drift will be caught rather than shipped — but
the duplication itself remains. Expressing the rules once as data and generating or
validating both implementations against it would remove the drift surface entirely, and is
the natural next step if this codebase keeps growing.
