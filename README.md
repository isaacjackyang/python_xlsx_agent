# xl-agent-core

`xl-agent-core` 是一個用 Python 打造的 spreadsheet agent substrate，目標是讓 Agent / LLM 能更安全地理解與操作試算表。
`xl-agent-core` is a Python spreadsheet agent substrate designed to help Agents / LLMs understand and operate on spreadsheets safely.

它的核心不是做另一個「Excel 讀檔工具」，而是建立一層面向 agent 的 spreadsheet interaction layer。
Its goal is not to be "another Excel reader," but to provide an agent-facing spreadsheet interaction layer.

目前版本已同時涵蓋 Phase 1 的結構理解，以及 Phase 2 的安全修改、重算與驗證工作流。
The current version covers both Phase 1 structure understanding and Phase 2 safe mutation, recalculation, and verification workflows.

這個專案同時是 Python 套件，也是命令列工具；它不是只有人類手動使用的 CLI，也不是只有內部 `.py` 模組。
This project is both a Python package and a command-line tool; it is not only a human-operated CLI, and it is not only a set of internal `.py` modules.

目前最主要的兩種使用方式是：用 CLI 輸出穩定 JSON 給 agent / automation 使用，或直接從 Python 程式中 import service classes 來調用。
The two main usage modes today are: using the CLI to emit stable JSON for agents / automation, or importing the service classes directly from Python code.

如果你在問「它會輸出 JSON 給 agent 看嗎？」答案是會。
If you are asking, "Does it output JSON for an agent to consume?" the answer is yes.

像 `xl probe ...`、`xl read table ...`、`xl verify diff ...` 這些命令都會回傳固定結構的 JSON envelope，包含 `operation`、`status`、`data`、`warnings`、`sources`、`metadata`。
Commands such as `xl probe ...`, `xl read table ...`, and `xl verify diff ...` return a fixed JSON envelope containing `operation`, `status`, `data`, `warnings`, `sources`, and `metadata`.

如果你在問「它也有 `.py` 檔給 agent 或開發者直接用嗎？」答案也是會。
If you are asking, "Does it also provide `.py` files that an agent or developer can call directly?" the answer is also yes.

你可以直接 import `xl_agent_core.core.recon`、`xl_agent_core.core.readers`、`xl_agent_core.core.formulas`、`xl_agent_core.core.mutate`、`xl_agent_core.core.verify` 這些模組，把它當成 Python SDK 使用。
You can directly import modules such as `xl_agent_core.core.recon`, `xl_agent_core.core.readers`, `xl_agent_core.core.formulas`, `xl_agent_core.core.mutate`, and `xl_agent_core.core.verify`, and use the project as a Python SDK.

目前已經包含第一版 MCP server，能把現有的 read / mutate / recalc / verify 能力直接暴露成 MCP tools。
The project now includes a first MCP server that exposes the existing read / mutate / recalc / verify capabilities as MCP tools.

下面是 CLI 給 agent 使用時的典型輸出型態。
Here is the typical output shape when the CLI is used by an agent.

```json
{
  "operation": "read.table",
  "status": "ok",
  "data": {
    "sheet": "Model",
    "region": "A1:C4"
  },
  "warnings": [],
  "sources": [
    {
      "workbook_path": "C:/path/to/workbook.xlsx",
      "sheet": "Model",
      "range_ref": "A1:C4"
    }
  ],
  "metadata": {}
}
```

下面是 Python 直接調用的典型寫法。
Here is the typical pattern for using the project directly from Python.

```python
from xl_agent_core.core.readers import ReadService

reader = ReadService()
result = reader.read_table("workbook.xlsx", "Model", "auto")
print(result.region)
print(result.rows)
```

---

## 設計原則 / Design Principles

這個專案延續的是 spreadsheet agent 的操作哲學，而不是單純照抄 CLI 功能表。
This project follows spreadsheet-agent operating principles rather than copying a CLI feature list.

- 先漸進式揭露，再做精讀
- Progressive disclosure before deep reads

- 以 region-aware 操作取代整張 sheet dump
- Region-aware operations instead of whole-sheet dumping

- 把公式關係當成一等公民資料
- Formula relationships as first-class data

- 每個 payload 都附 warnings 與 source refs
- Attach warnings and source refs to every payload

- 預設採用 non-destructive copy-on-write 修改流程
- Default to non-destructive copy-on-write mutation workflows

---

## 安裝 / Installation

使用 editable install 來安裝本地開發版本。
Install the local development version in editable mode.

```powershell
.\install.cmd
.\.venv\Scripts\activate
```

### 安裝前提 / Prerequisites

建議在 Windows 上使用 Python 3.10 以上版本；如果你要使用 `excel_com` 重算後端，機器上還需要安裝 Microsoft Excel。
Use Python 3.10 or newer on Windows; if you want the `excel_com` recalculation backend, Microsoft Excel must also be installed.

`install.cmd` 會自動偵測 `py -3` 或 `python`，建立 `.venv`，升級打包工具，然後以 editable mode 安裝專案。
`install.cmd` automatically detects `py -3` or `python`, creates `.venv`, upgrades packaging tools, and installs the project in editable mode.

### 安裝選項 / Installation Options

如果你想重建虛擬環境，可以加入 `--recreate`。
Use `--recreate` if you want to rebuild the virtual environment from scratch.

如果你要指定特定 Python，可使用 `--python`。
Use `--python` when you want to point to a specific Python executable.

```powershell
.\install.cmd --recreate
.\install.cmd --python "C:\Python314\python.exe"
```

---

## 快速測試 / Smoke Test

可以用下面的指令快速驗證目前骨架是否可運作。
Use the following command to quickly verify that the current scaffold works.

```powershell
$env:PYTHONPATH = "src"
python -m unittest -q
```

你也可以直接用 `start.cmd test`，它會自動使用 `.venv` 來執行測試。
You can also use `start.cmd test`, which automatically runs tests inside `.venv`.

```powershell
.\start.cmd test
```

---

## MCP 伺服器 / MCP Server

如果你要把這個專案當成真正的 spreadsheet MCP server 來用，現在已經可以。
If you want to use this project as a real spreadsheet MCP server, you can now do that.

目前第一版 MCP 以 stdio 為主要 transport，另外也支援 `streamable-http` 啟動模式。
The first MCP version primarily targets stdio transport, and it also supports `streamable-http`.

最簡單的啟動方式是直接用 `start.cmd mcp`。
The simplest way to start it is with `start.cmd mcp`.

```powershell
.\start.cmd mcp
```

如果你要直接使用虛擬環境中的 entry point，也可以用 `xl-mcp`。
If you want to use the virtual environment entry point directly, you can also run `xl-mcp`.

```powershell
.\.venv\Scripts\xl-mcp --help
.\.venv\Scripts\xl-mcp --transport stdio
```

如果你要用 HTTP transport，可以這樣啟動。
If you want to use HTTP transport, you can start it like this.

```powershell
.\start.cmd mcp --transport streamable-http --host 127.0.0.1 --port 8000 --path /mcp
```

這個 MCP server 目前會提供 workbook reconnaissance、exact reads、formula tracing、copy-on-write mutation、recalc、proof、diff 等工具。
This MCP server currently exposes tools for workbook reconnaissance, exact reads, formula tracing, copy-on-write mutation, recalculation, proof, and diff workflows.

你可以先呼叫 `server_info` 看它目前提供哪些工具與重算後端。
You can call `server_info` first to see which tools and recalculation backends are currently available.

---

## 專案結構 / Project Structure

這個 package 不是依照檔案格式細節拆分，而是依照 agent 最容易出錯的工作流來拆分。
This package is organized around agent failure modes instead of file-format trivia.

- `core/contracts.py`: 定義穩定且可預測的 response models
- `core/contracts.py`: deterministic response models

- `core/loader.py`: workbook 載入與共用 helper
- `core/loader.py`: workbook loading helpers

- `core/recon.py`: workbook overview 與 sheet reconnaissance
- `core/recon.py`: workbook overview and sheet reconnaissance

- `core/regions.py`: region detection heuristics
- `core/regions.py`: region detection heuristics

- `core/readers.py`: region-grounded reads
- `core/readers.py`: region-grounded reads

- `core/layout.py`: layout-aware cell inspection
- `core/layout.py`: layout-aware cell inspection

- `core/formulas.py`: formula parsing 與 trace helpers
- `core/formulas.py`: formula parsing and trace helpers

- `core/mutate.py`: copy、write-cells、replace-formulas
- `core/mutate.py`: copy, write-cells, and replace-formulas

- `core/recalc.py`: `auto / excel_com / openpyxl` 重算後端
- `core/recalc.py`: `auto / excel_com / openpyxl` recalculation backends

- `core/verify.py`: proof 與 diff 驗證
- `core/verify.py`: proof and diff verification

- `cli/main.py`: CLI transport
- `cli/main.py`: CLI transport

- `mcp/server.py`: MCP transport 與 tool wrappers
- `mcp/server.py`: MCP transport and tool wrappers

---

## 目前範圍 / Current Scope

Phase 1 已在這個 repo 中實作完成，重點是讓 agent 先看懂 workbook，再精讀正確區塊。
Phase 1 is implemented in this repository and focuses on helping an agent understand a workbook before reading exact regions.

- `probe` 工作簿結構
- `probe` workbook structure

- `read sheets` 讀取工作表清單
- `read sheets` list workbook sheets

- `read names` 讀取 named items
- `read names` list named items

- `detect regions` 偵測候選區塊
- `detect regions` detect candidate regions

- `read table` 精讀表格區域
- `read table` read a table region

- `read layout` 讀取版面與顯示脈絡
- `read layout` inspect layout-aware range details

- `inspect cells` 精讀指定儲存格
- `inspect cells` inspect exact cells

- `analyze formula-trace` 追蹤公式關聯
- `analyze formula-trace` trace formula relationships

Phase 2 目前也已經有可用 MVP，聚焦在 non-destructive editing、recalc 與 verify。
Phase 2 now has a working MVP focused on non-destructive editing, recalculation, and verification.

- `workbook copy` 建立 copy-on-write 工作副本
- `workbook copy` create a copy-on-write workbook copy

- `mutate write-cells` 安全寫入 literal values
- `mutate write-cells` safely write literal values

- `mutate replace-formulas` 安全替換公式
- `mutate replace-formulas` safely replace formulas

- `workbook recalc` 執行重算後端
- `workbook recalc` run a recalculation backend

- `verify proof` 驗證指定 targets 的前後變化
- `verify proof` compare explicit proof targets

- `verify diff` 比對整體 workbook 變化
- `verify diff` compare workbook-wide changes

---

## 命令列介面 / CLI

下面這些指令代表目前的主要 interaction surface。
These commands represent the current interaction surface.

### 啟動方式 / Startup Modes

`start.cmd` 是最簡單的入口；沒帶參數時，它會顯示 CLI 說明。
`start.cmd` is the easiest entry point; without arguments, it shows the CLI help output.

`start.cmd shell` 會開一個已啟用 `.venv` 的命令列視窗。
`start.cmd shell` opens a command prompt with `.venv` already activated.

`start.cmd [xl arguments]` 會把後面的參數直接轉交給 `python -m xl_agent_core`。
`start.cmd [xl arguments]` forwards the remaining arguments directly to `python -m xl_agent_core`.

```powershell
.\start.cmd
.\start.cmd shell
.\start.cmd probe workbook.xlsx
.\start.cmd verify diff before.xlsx after.xlsx
```

### 第一階段 / Phase 1

```powershell
xl probe workbook.xlsx
xl read sheets workbook.xlsx
xl read names workbook.xlsx
xl detect regions workbook.xlsx --sheet Model
xl read table workbook.xlsx --sheet Model --region auto
xl read layout workbook.xlsx --sheet Model --range A1:H30
xl inspect cells workbook.xlsx --sheet Model --refs B2 D10:F12
xl analyze formula-trace workbook.xlsx --sheet Model --cell C10 --direction precedents
```

第一階段建議使用順序是 `probe -> read sheets / read names -> detect regions -> read table / read layout -> inspect cells -> analyze formula-trace`。
The recommended Phase 1 sequence is `probe -> read sheets / read names -> detect regions -> read table / read layout -> inspect cells -> analyze formula-trace`.

如果你不確定哪一塊表才是對的，先用 `detect regions`，再把回傳的 `range_ref` 或 `region_x` 標籤帶進 `read table`。
If you are not sure which table is the right one, start with `detect regions`, then pass the returned `range_ref` or `region_x` label into `read table`.

### 第二階段 / Phase 2

```powershell
xl workbook copy in.xlsx out.xlsx
xl mutate write-cells out.xlsx --sheet Model --set B2=5 --output out_values.xlsx --recalc-backend auto
xl mutate replace-formulas out.xlsx --sheet Model --set C2==B2*12 --output out_formulas.xlsx --recalc-backend auto
xl workbook recalc out_values.xlsx --backend auto
xl verify proof in.xlsx out_values.xlsx --targets Model!C2 Summary!B2
xl verify diff in.xlsx out_values.xlsx
```

第二階段預設走 non-destructive 流程，也就是先複製 workbook，再修改副本，最後再驗證變化。
Phase 2 defaults to a non-destructive workflow: copy the workbook first, mutate the copy, then verify the result.

`mutate write-cells` 適合寫入 literal values，例如數字、字串、布林值或 JSON 可解析的標量。
`mutate write-cells` is for literal values such as numbers, strings, booleans, or JSON-parsable scalars.

`mutate replace-formulas` 則是直接改公式，輸入時如果公式本身有 `=`，請使用 `REF==FORMULA` 這種寫法，因為指令會在第一個 `=` 位置切開。
`mutate replace-formulas` directly updates formulas; if the formula itself starts with `=`, use the `REF==FORMULA` form because the command splits on the first `=`.

```powershell
.\start.cmd mutate write-cells budget.xlsx --sheet Model --set B2=5 --set B3=7 --output budget_edit.xlsx --dry-run
.\start.cmd mutate write-cells budget.xlsx --sheet Model --set B2=5 --output budget_edit.xlsx --recalc-backend auto
.\start.cmd mutate replace-formulas budget.xlsx --sheet Model --set C2==B2*12 --output budget_formula.xlsx --recalc-backend excel_com
```

### 常見工作流 / Common Workflows

如果你只是要先理解一本 workbook，建議先跑 reconnaissance workflow。
If your goal is to understand a workbook first, begin with the reconnaissance workflow.

```powershell
.\start.cmd probe workbook.xlsx
.\start.cmd read sheets workbook.xlsx
.\start.cmd detect regions workbook.xlsx --sheet Model
.\start.cmd read table workbook.xlsx --sheet Model --region auto --pretty
```

如果你要安全地修改並驗證結果，建議走 copy -> mutate -> recalc -> verify 的順序。
If you need a safe edit workflow, follow the order copy -> mutate -> recalc -> verify.

```powershell
.\start.cmd workbook copy in.xlsx work.xlsx
.\start.cmd mutate write-cells work.xlsx --sheet Model --set B2=5 --output work_edited.xlsx --recalc-backend auto
.\start.cmd verify proof in.xlsx work_edited.xlsx --targets Model!C2 Summary!B2 --pretty
.\start.cmd verify diff in.xlsx work_edited.xlsx --pretty
```

### 實戰範例 / Real Example

假設你手上有一份 `budget.xlsx`，其中 `Model!B2` 是輸入數量，會影響 `Model!C2` 和 `Summary!B2`。
Assume you have a workbook named `budget.xlsx`, where `Model!B2` is an input quantity that affects `Model!C2` and `Summary!B2`.

**步驟 1 / Step 1**

先看懂 workbook 結構，不要直接修改。
Understand the workbook structure before making any changes.

```powershell
.\start.cmd probe budget.xlsx --pretty
.\start.cmd read sheets budget.xlsx
.\start.cmd detect regions budget.xlsx --sheet Model --pretty
.\start.cmd read table budget.xlsx --sheet Model --region auto --pretty
```

先確認 `sheet_order`、`range_ref`、`warnings`，確定你要改的是正確區塊。
Check `sheet_order`, `range_ref`, and `warnings` first to confirm that you are targeting the correct region.

**步驟 2 / Step 2**

先做 dry-run，確認系統理解的修改計畫是正確的。
Run a dry-run first so you can confirm the planned mutation before writing anything.

```powershell
.\start.cmd mutate write-cells budget.xlsx --sheet Model --set B2=5 --output budget_edit.xlsx --dry-run --pretty
```

這一步要看 `targets`、`before_value`、`after_value`、`note`，確認它是改 `B2`，而不是改到別的儲存格。
At this step, inspect `targets`, `before_value`, `after_value`, and `note` to make sure the plan changes `B2` and nothing else.

**步驟 3 / Step 3**

確認 dry-run 沒問題後，再對副本執行實際修改並重算。
After the dry-run looks correct, apply the real edit to a copied workbook and recalculate it.

```powershell
.\start.cmd workbook copy budget.xlsx budget_work.xlsx
.\start.cmd mutate write-cells budget_work.xlsx --sheet Model --set B2=5 --output budget_edit.xlsx --recalc-backend auto --pretty
```

如果機器上有 Excel，`auto` 會優先使用 `excel_com`，這樣公式快取值也會一起更新。
If Excel is installed, `auto` prefers `excel_com`, which also refreshes cached formula values.

**步驟 4 / Step 4**

用 `verify proof` 驗證你真正關心的輸出儲存格。
Use `verify proof` to validate the exact output cells you care about.

```powershell
.\start.cmd verify proof budget.xlsx budget_edit.xlsx --targets Model!C2 Summary!B2 --pretty
```

這一步要看每個 target 的 `before_value`、`after_value`、`classification`，確認修改是否帶來你預期的結果。
Look at each target's `before_value`, `after_value`, and `classification` to confirm that the change produced the expected result.

**步驟 5 / Step 5**

最後用 `verify diff` 看整體波及範圍，避免只改對一格、卻連帶破壞別的地方。
Finish with `verify diff` to understand the full impact and avoid fixing one cell while accidentally breaking others.

```powershell
.\start.cmd verify diff budget.xlsx budget_edit.xlsx --pretty
```

通常你會看到像 `value_changed`、`formula_changed` 或 `recalc_impact` 這類分類；如果出現超出預期的變化，就回頭檢查 `warnings`、region 選擇與公式依賴。
You will usually see classifications such as `value_changed`, `formula_changed`, or `recalc_impact`; if the impact is broader than expected, review `warnings`, region selection, and formula dependencies.

**步驟 6 / Step 6**

如果結果不如預期，可以回到公式追蹤層重新定位依賴來源。
If the result is not what you expected, go back to formula tracing to locate the dependency chain.

```powershell
.\start.cmd analyze formula-trace budget.xlsx --sheet Summary --cell B2 --direction precedents --pretty
.\start.cmd analyze formula-trace budget_edit.xlsx --sheet Model --cell C2 --direction dependents --pretty
```

這能幫你確認是輸入值改錯、公式改錯，還是其實影響鏈比想像中更長。
This helps you determine whether the issue comes from the wrong input, the wrong formula, or a longer dependency chain than expected.

所有指令都會輸出統一且穩定的 JSON envelope。
All commands return a stable JSON envelope.

- `operation`: 操作名稱
- `operation`: operation name

- `status`: 執行狀態
- `status`: execution status

- `data`: 主要資料內容
- `data`: primary payload

- `warnings`: 機器可讀的警告訊息
- `warnings`: machine-readable warnings

- `sources`: 來源引用資訊
- `sources`: source references

- `metadata`: 補充中繼資料
- `metadata`: supplemental metadata

`warnings` 很重要，它會告訴你 hidden rows、merged cells、待重算、或 openpyxl roundtrip 等風險。
`warnings` is important because it reports risks such as hidden rows, merged cells, pending recalculation, or openpyxl roundtrip concerns.

`sources` 則是給 agent 或上層系統引用的來源資訊，方便把回答綁回 workbook、sheet 與 range。
`sources` provides source references for agents or higher-level systems so answers can be grounded back to workbook, sheet, and range.

---

## 重算策略 / Recalculation Strategy

`auto` 會優先使用 `excel_com`，這在 Windows 且可用 Excel COM 時可以真正更新 cached formula results。
`auto` prefers `excel_com`, which can genuinely refresh cached formula results on Windows when Excel COM is available.

如果沒有可用的 Excel COM，`openpyxl` backend 會把 workbook 標記為待重算，但不會真的計算公式結果。
If Excel COM is not available, the `openpyxl` backend marks the workbook for recalculation but does not actually compute formula results.

這代表 `verify proof` 和 `verify diff` 在公式值驗證上，最好搭配 `excel_com` 使用。
That means `verify proof` and `verify diff` are most reliable for formula-value validation when used with `excel_com`.
