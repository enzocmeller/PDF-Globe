# Soybean rankings — Power Query setup

Adds **Soybeans, Soybean Meal, Soybean Oil** top-5 tables (Producers / Exporters /
Importers / Consumers), built the same way as your wheat/corn/barley tables.

Columns are plain marketing years — **`23/24  24/25  25/26  26/27`** (no May/Jun
suffix). Each query reads **one** data file, so you get a given month by pointing
it at that month's file:

| File | Month / release |
|------|-----------------|
| `psd_oilseeds_latest.csv` | newest release (e.g. June) |
| `psd_oilseeds_prior.csv`  | the release before it (e.g. May) |

So the `26/27` column shows **June's** 26/27 number when you read the *latest*
file, and **May's** 26/27 number when you read the *prior* file. To compare months,
make one rankings query per file (see step 4).

---

## 1. Prerequisite — get the data

Run `update_psd.py` once. It creates:

```
data\psd_oilseeds_latest.csv
data\psd_oilseeds_prior.csv
```

(Commodities inside: `Oilseed, Soybean`, `Meal, Soybean`, `Oil, Soybean`.)

---

## 2. Create the queries

In Excel: **Data → Get Data → Launch Power Query Editor → New Source → Other
Sources → Blank Query → Home → Advanced Editor**, paste, **Done**, then rename the
query (top-right) to the bold name. Repeat for all three.

> `PSD_Oilseeds` and `fnTop5_Oil` are helpers — right-click each → untick
> **Enable load**. Only the `Soybean_Rankings` query (step 3) loads to a sheet.

### Query 1 — name it: `PSD_Oilseeds`
```m
(sourceFile as text) as table =>
let
    IncludeEU  = true,   // true = include "European Union" as an entity
    WBFolder   = try Excel.CurrentWorkbook(){[Name="WorkbookPath"]}[Content]{0}[Column1] otherwise null,
    UseWB      = WBFolder <> null and WBFolder <> "" and Text.StartsWith(WBFolder, "http") = false,
    Fallback   = "C:\Users\ecoutinh\Desktop\Final Python Project\data\",
    FolderPath = if UseWB then WBFolder & "data\" else Fallback,

    Source   = Csv.Document(File.Contents(FolderPath & sourceFile),
                  [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Promoted = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Kept     = Table.SelectColumns(Promoted,
                  {"Commodity_Description","Country_Name","Market_Year",
                   "Attribute_Description","Value"}),
    Typed    = Table.TransformColumnTypes(Kept,
                  {{"Market_Year", Int64.Type}, {"Value", type number}}),

    YearsF   = Table.SelectRows(Typed, each List.Contains({2023,2024,2025,2026}, [Market_Year])),
    Attrs    = {"Production","Exports","Imports","Domestic Consumption"},
    AttrF    = Table.SelectRows(YearsF, each List.Contains(Attrs, [Attribute_Description])),
    CountryF = if IncludeEU then AttrF
               else Table.SelectRows(AttrF, each [Country_Name] <> "European Union"),

    AddMY    = Table.AddColumn(CountryF, "MY", each
                  if [Market_Year] = 2026 then "26/27"
                  else if [Market_Year] = 2025 then "25/26"
                  else if [Market_Year] = 2024 then "24/25"
                  else if [Market_Year] = 2023 then "23/24"
                  else null, type text),

    Final    = Table.SelectColumns(AddMY,
                  {"Commodity_Description","Attribute_Description","Country_Name","MY","Value"})
in
    Final
```

### Query 2 — name it: `fnTop5_Oil`
```m
(base as table, commodity as text, attribute as text) as table =>
let
    Periods  = {"26/27","25/26","24/25","23/24"},
    Filtered = Table.SelectRows(base, each [Commodity_Description] = commodity
                                       and [Attribute_Description] = attribute),
    Pivoted  = Table.Pivot(Filtered, Periods, "MY", "Value", List.Sum),
    Sorted   = Table.Sort(Pivoted, {{"26/27", Order.Descending}}),
    Top5     = Table.FirstN(Sorted, 5),
    Ranked   = Table.AddIndexColumn(Top5, "Rank", 1, 1, Int64.Type),
    Out      = Table.SelectColumns(Ranked, {"Rank","Country_Name"} & Periods)
in
    Out
```

---

## 3. The rankings query (this one loads to a sheet)

### Query 3 — name it: `Soybean_Rankings`
```m
let
    // ---- which month: point at the file you want --------------------------
    SourceFile  = "psd_oilseeds_latest.csv",   // <- newest release (e.g. June)
    // -----------------------------------------------------------------------
    Base = PSD_Oilseeds(SourceFile),

    Commodities = {
        [Raw="Oilseed, Soybean", Label="Soybeans"],
        [Raw="Meal, Soybean",    Label="Soybean Meal"],
        [Raw="Oil, Soybean",     Label="Soybean Oil"]
    },
    Metrics = {
        [Metric="Producers", Attribute="Production"],
        [Metric="Exporters", Attribute="Exports"],
        [Metric="Importers", Attribute="Imports"],
        [Metric="Consumers", Attribute="Domestic Consumption"]
    },
    Combos = List.Combine(List.Transform(Commodities, (c) =>
                List.Transform(Metrics, (m) =>
                    [Commodity=c[Label], RawCommodity=c[Raw],
                     Metric=m[Metric], Attribute=m[Attribute]]))),
    AsTable  = Table.FromRecords(Combos),
    WithData = Table.AddColumn(AsTable, "Data",
                  each fnTop5_Oil(Base, [RawCommodity], [Attribute])),
    Expanded = Table.ExpandTableColumn(WithData, "Data",
                  {"Rank","Country_Name","26/27","25/26","24/25","23/24"}),
    Cleaned  = Table.SelectColumns(Expanded,
                  {"Commodity","Metric","Rank","Country_Name",
                   "26/27","25/26","24/25","23/24"}),
    Sorted   = Table.Sort(Cleaned, {{"Commodity", Order.Ascending},
                                    {"Metric", Order.Ascending},
                                    {"Rank", Order.Ascending}}),
    Reordered = Table.ReorderColumns(Sorted,
                  {"Commodity","Metric","Rank","Country_Name",
                   "23/24","24/25","25/26","26/27"})
in
    Reordered
```

Output columns:

| Commodity | Metric | Rank | Country_Name | 23/24 | 24/25 | 25/26 | 26/27 |
|-----------|--------|------|--------------|-------|-------|-------|-------|

---

## 4. Get BOTH months (the whole point)

`Soybean_Rankings` above = the **latest** file (June). To also get the **prior**
file (May), duplicate it:

1. Right-click `Soybean_Rankings` → **Duplicate** → rename the copy
   `Soybean_Rankings_Prior`.
2. In the copy, change the one line:
   ```m
   SourceFile  = "psd_oilseeds_prior.csv",   // <- prior release (e.g. May)
   ```

Now you have two tables with identical `23/24 … 26/27` columns — one per month —
and you compare the `26/27` column between them to see the month-over-month
revision.

---

## 5. Build the slides

The fastest route, since `Soybean_Rankings` matches your grains layout:

1. Right-click an existing sheet (e.g. **`Corn S&D`**) → **Move or Copy → create a
   copy**. Rename it `Soybean S&D`.
2. Repoint its tables at `Soybean_Rankings`, filtered to `Commodity = "Soybeans"`
   (and the relevant Metric per table block).
3. Repeat for `Soybean Meal S&D` (`Commodity = "Soybean Meal"`) and
   `Soybean Oil S&D` (`Commodity = "Soybean Oil"`).

The sheet names `Soybean S&D`, `Soybean Meal S&D`, `Soybean Oil S&D` are already in
`SLIDE_ORDER` in `export_pdf_deck.py`, so the PDF deck picks them up automatically
once they exist.

---

## Notes
- **Crush** is available too — to use it (soybeans are mostly crushed, not
  "consumed"), add `[Metric="Crush", Attribute="Crush"]` to the `Metrics` list.
- `IncludeEU = true` matches your grains queries. Set it `false` in `PSD_Oilseeds`
  to drop the European Union as an entity.
- Values are in **1,000 metric tons** (USDA PSD unit).
- All three soybean commodities have Production / Exports / Imports / Domestic
  Consumption / Crush, so any of those work as a Metric.
