## Power BI DAX & M 言語クックブック（実装パターン集）

この資料は、導入説明なしで即使える「実装パターン」と「ベストプラクティス」を日本語で体系化したクックブックです。DAX と M（Power Query）双方の実戦的レシピ、パフォーマンス最適化、トラブルシュート、テンプレートをまとめています。

— 目次 —
- DAX 基本パターン（実務コア）
- DAX 応用・モデリング（計算グループ、関係管理、RLS）
- DAX パフォーマンス最適化
- M（Power Query）高度変換パターン
- M パフォーマンス・フォールディング・インクリメンタルリフレッシュ
- M と DAX のエンドツーエンドシナリオ
- トラブルシュート／アンチパターン／テスト
- チェックリストと再利用テンプレート

---

## DAX 基本パターン（実務コア）

### ベースメジャー（基礎を分離する）
ベースメジャーを分離して、派生メジャー（率や比較系）はそれらを参照します。

```DAX
[Total Sales] = SUM('Fact Sales'[SalesAmount])

[Total Cost] = SUM('Fact Sales'[CostAmount])

[Profit] = [Total Sales] - [Total Cost]

[Profit Margin %] = 
VAR sales = [Total Sales]
RETURN 
    DIVIDE([Profit], sales)
```

ポイント:
- 派生メジャーでは `DIVIDE` を使用（ゼロ割り対策）。
- すべての下位集計はベースメジャーで行い、`X` 系イテレータは必要最低限に。

### フィルターコンテキスト制御（CALCULATE/ALL/KEEPFILTERS）
```DAX
-- 全体（フィルター解除）での売上
[Total Sales (All)] = CALCULATE([Total Sales], ALL('Dim Date'))

-- スライサー条件は残しつつ年のフィルターだけを解除
[Total Sales (All Years)] = CALCULATE([Total Sales], ALLEXCEPT('Dim Date', 'Dim Date'[Month]))

-- 既存フィルターを保ったまま追加条件を狭める
[Sales for Product A] = 
CALCULATE(
    [Total Sales],
    KEEPFILTERS('Dim Product'[ProductName] = "A")
)

-- 完全にフィルターを外す（ALL より意図が明確）
[Total Sales (Remove Filters)] = CALCULATE([Total Sales], REMOVEFILTERS('Dim Date'))
```

補足:
- `ALL` は列・テーブルを指定した部分のフィルターを外します。
- `REMOVEFILTERS` は同様の意図をより明確化。
- `KEEPFILTERS` は既存のフィルターを尊重しつつ条件を絞り込むのに有効。

### 時系列（Time Intelligence）
```DAX
[Sales PY] = CALCULATE([Total Sales], DATEADD('Dim Date'[Date], -1, YEAR))

[Sales YoY %] = 
VAR py = [Sales PY]
RETURN DIVIDE([Total Sales] - py, py)

-- 会計年度（例：7月開始）の YTD
[Sales FYTD] = 
TOTALYTD(
    [Total Sales],
    'Dim Date'[Date],
    "06/30"  -- 米国文化圏の M/D 表記。環境に応じて適切に。
)

-- ローリング12か月
[Sales Rolling 12M] = 
CALCULATE(
    [Total Sales],
    DATESINPERIOD('Dim Date'[Date], MAX('Dim Date'[Date]), -12, MONTH)
)
```

注意:
- Time Intelligence 関数は適切な`日付ディメンション`と連続する日付が必須。
- Auto Date/Time は本番では無効化し、明示的な `Dim Date` を使用。

### 半加法メジャー（残高/在庫など）
```DAX
[Ending Balance] = 
CALCULATE(
    LASTNONBLANKVALUE('Dim Date'[Date], [Total Sales]),
    VALUES('Dim Account'[AccountKey])
)

[Opening Balance] = 
CALCULATE([Ending Balance], DATEADD('Dim Date'[Date], -1, DAY))
```

別解（在庫スナップショット）:
```DAX
[Stock (Last Non Blank)] = 
CALCULATE(
    LASTNONBLANKVALUE('Dim Date'[Date], MAX('Fact Inventory'[StockLevel])),
    MAX('Fact Inventory'[StockLevel])
)
```

### 非アクティブなリレーションを使う（USERELATIONSHIP）
```DAX
[Sales by Ship Date] = 
CALCULATE(
    [Total Sales],
    USERELATIONSHIP('Dim Date'[Date], 'Fact Sales'[ShipDate])
)
```

### 仮想関係・ブリッジ（TREATAS/CROSSFILTER）
```DAX
-- マッピングテーブル（仮想）を使ってフィルター伝播
[Sales by Custom Group] = 
VAR vMap = 
    SUMMARIZE(
        'Dim Product',
        'Dim Product'[ProductKey],
        "Group", 'Dim Product'[CustomGroup]
    )
RETURN
    CALCULATE(
        [Total Sales],
        TREATAS(VALUES(vMap[Group]), 'Dim Group'[Group])
    )

-- 多対多の方向を一時的に変更
[Sales BiDir] = 
CALCULATE(
    [Total Sales],
    CROSSFILTER('Dim A'[Key], 'Dim B'[Key], BOTH)
)
```

### セグメンテーション（RFM/ABC）
```DAX
[Customer Sales Rank] = 
RANKX(
    ALL('Dim Customer'[CustomerKey]),
    [Total Sales],
    ,
    DESC,
    DENSE
)

[ABC Segment] = 
VAR v = [Customer Sales Rank]
RETURN 
    SWITCH(TRUE(),
        v <= 10, "A",
        v <= 100, "B",
        "C"
    )

-- ABC 構成比（仮想テーブル）
[ABC Share %] = 
VAR vt = 
    ADDCOLUMNS(
        SUMMARIZE('Dim Customer', 'Dim Customer'[CustomerKey]),
        "Sales", [Total Sales],
        "Rank", RANKX(ALL('Dim Customer'[CustomerKey]), [Total Sales], , DESC)
    )
VAR totalTop10 = 
    CALCULATE(
        SUMX(FILTER(vt, [Rank] <= 10), [Sales])
    )
RETURN DIVIDE(totalTop10, [Total Sales (All)])
```

### 動的タイトル/注釈
```DAX
[Dynamic Title] = 
VAR year = SELECTEDVALUE('Dim Date'[Year], "全期間")
VAR prod = SELECTEDVALUE('Dim Product'[ProductName], "全商品")
RETURN prod & " の " & year & " 年 売上"
```

### RLS（行レベルセキュリティ）
```DAX
-- ユーザーのメールドメインで地域を制限
'Dim Region'[Region] = 
VAR email = USERPRINCIPALNAME()
RETURN 
    CONTAINSSTRING(email, "@east.example.com") && 'Dim Region'[Region] = "East"

-- 階層（PATH）を使ったツリー制御
'Dim Org'[Path] = PATH('Dim Org'[ManagerId], 'Dim Org'[EmployeeId])
-- RLS フィルター：PATHCONTAINS([Path], USERID())
```

---

## DAX 応用・モデリング

### 計算グループ（Tabular Editor 使用）
- 時系列グループ: Current, PY, YoY%, YTD, Rolling
- 通貨換算グループ: Local, USD, JPY（FormatString: "#,0;(#,0)" 等）

例（式の概念）:
```DAX
-- Calc Item: PY
CALCULATE(SELECTEDMEASURE(), DATEADD('Dim Date'[Date], -1, YEAR))

-- Calc Item: YoY %
VAR curr = SELECTEDMEASURE()
VAR py = CALCULATE(SELECTEDMEASURE(), DATEADD('Dim Date'[Date], -1, YEAR))
RETURN DIVIDE(curr - py, py)
```

注意: 計算グループは強力だが副作用（フォーマット/並び順/ALLSELECTED の効き等）を理解して設計する。

### 多対多と関係方向
- 可能な限り単方向。両方向は最小限・意図的に。
- 仮想関係（TREATAS）による代替。

### ALLSELECTED と相対比率
```DAX
[Share within Selection %] = 
DIVIDE([Total Sales], CALCULATE([Total Sales], ALLSELECTED('Dim Product')))
```

### NATURAL* 系の回避
`NATURALINNERJOIN/NATURALLEFTOUTERJOIN` は列名一致依存で脆弱。キーを明示して `TREATAS` や `INTERSECT` 等の組合せで堅牢化。

---

## DAX パフォーマンス最適化

- 基本は「列演算＋集計」。行ごとのイテレーション（`SUMX` など）は最小に。
- ベースメジャーに集約ロジックを寄せ、派生は再集計のみ。
- `VAR` で重いメジャーの再計算を避ける。
- フィルター解除は必要最小範囲（`REMOVEFILTERS`/`ALLEXCEPT`）で。
- `ALLSELECTED` は重くなりやすい。必要性を再検討。
- 高カーディナリティ列の可視化粒度を下げる（バケツ化、ディメンション化）。
- モデル: 事実表は数値列のみ保持、文字列はディメンションへ。Star Schema 徹底。
- 計算列よりメジャー優先（行固定計算が本当に必要なときのみ計算列）。
- DAX Studio/Performance Analyzer/Server Timings で計測し、クエリプランを確認。

サンプル（`VAR` での最適化）:
```DAX
[Profit Margin % (Optimized)] = 
VAR sales = [Total Sales]
VAR profit = [Profit]
RETURN DIVIDE(profit, sales)
```

---

## M（Power Query）高度変換パターン

### 参照（Reference）と段階分離（Staging）
- 元クエリを「ステージング」（変換最小）として保持し、そこから `参照` でファクト/ディム等の派生を作る。
- ロードは最終テーブルのみ有効化。ステージングはロード無効。

### クエリフォールディング（折り畳み）最優先
- フィルタ、列選択、結合、集計は可能な限りデータソース側で実行。
- 非フォールディングの関数（Row-wise カスタム、正規表現、Web API など）は末尾に寄せる。
- `Table.Buffer` は最後の手段。フォールディングを破壊し得るため注意。

### 代表的な変換
```powerquery
let
  Source = Sql.Database("srv01", "SalesDW"),
  Fact = Source{[Schema="dbo",Item="FactSales"]}[Data],
  Filtered = Table.SelectRows(Fact, each [OrderDate] >= #date(2023,1,1)),
  Selected = Table.SelectColumns(Filtered, {"OrderDate","ProductKey","CustomerKey","SalesAmount","CostAmount"}),
  Typed = Table.TransformColumnTypes(Selected,{{"OrderDate", type date},{"SalesAmount", type number},{"CostAmount", type number}})
in
  Typed
```

### マージ・展開（Join）
```powerquery
let
  Sales = ...,
  Product = ...,
  Merged = Table.NestedJoin(Sales, {"ProductKey"}, Product, {"ProductKey"}, "Product", JoinKind.LeftOuter),
  Expanded = Table.ExpandTableColumn(Merged, "Product", {"Category","Name"}, {"Category","ProductName"})
in
  Expanded
```

### スキーマドリフト対策（列名変化に耐える）
```powerquery
let
  Source = Excel.Workbook(File.Contents("Customers.xlsx"), null, true),
  Table0 = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
  Promoted = Table.PromoteHeaders(Table0, [PromoteAllScalars=true]),
  Columns = Table.ColumnNames(Promoted),
  Safe = Table.TransformColumnTypes(Promoted, List.Transform(Columns, each {_, type text}), "ja-JP")
in
  Safe
```

### カスタム関数＋ページネーション（REST API）
```powerquery
// 関数: 1ページ取得
(page as number) as table =>
let
  Response = Web.Contents(
    "https://api.example.com", 
    [ RelativePath = "v1/items", Query = [ page = Text.From(page), size = "100" ] ]
  ),
  Json = Json.Document(Response),
  Items = Json[items],
  Table = Table.FromList(Items, Splitter.SplitByNothing(), null, null, ExtraValues.Error)
in
  Table
```

```powerquery
// 呼び出し: 総ページ数不明を List.Generate で
let
  Pages = List.Generate(
    () => [p=1, t=null],
    each t <> 0 or [p] = 1,
    each [ p = [p] + 1, t = if List.Count( try Function.Invoke("GetPage", {[p]}) otherwise {} ) < 100 then 0 else null ],
    each try Function.Invoke("GetPage", {[p]}) otherwise #table({}, {})
  ),
  Combined = Table.Combine(Pages)
in
  Combined
```

（実装時は関数名・エラー処理を適宜調整。`RelativePath` と `Query` を使うとキャッシュヒット率・認証安定が向上。）

### バイナリ結合（フォルダ取り込み）
```powerquery
let
  Source = Folder.Files("C:\\Data\\Sales"),
  Filtered = Table.SelectRows(Source, each Text.EndsWith([Extension], ".csv")),
  AddSource = Table.AddColumn(Filtered, "Csv", each Csv.Document(File.Contents([Folder Path] & [Name]), [Delimiter=",", Encoding=65001])),
  Expanded = Table.ExpandTableColumn(AddSource, "Csv", {"Column1","Column2","Column3"}),
  Typed = Table.TransformColumnTypes(Expanded, {{"Column1", type text}, {"Column2", type number}, {"Column3", type date}})
in
  Typed
```

### 例外処理
```powerquery
let
  SafeNumber = (t as any) as nullable number => try Number.From(t) otherwise null,
  Added = Table.AddColumn(Source, "Amount", each SafeNumber([AmountText]))
in
  Added
```

---

## M パフォーマンス／クエリフォールディング／インクリメンタルリフレッシュ

### パフォーマンス原則
- 早い段階で列選択・行フィルター（フォールディング維持）。
- 非フォールディング処理はできるだけ最後に。
- `Table.Buffer` はフォールディング阻害。必要性を計測で確認。
- 型指定は早期に実施（特に数値/日付）。
- 参照設計で共通前処理を集約し、重複計算を削減。

### インクリメンタルリフレッシュ（RangeStart/RangeEnd）
1) 日付/日時型のパラメータ `RangeStart`, `RangeEnd` を定義（データ型 DateTime/Date）。
2) 事実テーブルの `OrderDate` などに対し、`>= RangeStart` かつ `< RangeEnd` でフィルター。
3) 設定画面で保持期間・増分期間を設定。パーティションが自動生成される。

例（フィルター部分）:
```powerquery
let
  Source = Sql.Database("srv01", "SalesDW"),
  Fact = Source{[Schema="dbo",Item="FactSales"]}[Data],
  Filtered = Table.SelectRows(Fact, each [OrderDate] >= RangeStart and [OrderDate] < RangeEnd)
in
  Filtered
```

注意:
- フォールディング必須。データソースがフォールディング非対応（例: 一部のファイル/REST）の場合は代替設計を検討。
- パラメータ値は公開時にサービス側が置き換え。

---

## M と DAX のエンドツーエンドシナリオ

### 1) 通貨換算（M: 為替取込 → DAX: 計算グループ）
- M: 為替 API を日次で取り込み、`Dim CurrencyRate` を作成（Date, From, To, Rate）。
- DAX: 計算グループで `SELECTEDMEASURE()` にレートを適用。FormatString を通貨ごとに設定。

換算の概念式:
```DAX
VAR base = SELECTEDMEASURE()
VAR rate = CALCULATE( AVERAGE('Dim CurrencyRate'[Rate]), TREATAS(VALUES('Dim Date'[Date]), 'Dim CurrencyRate'[Date]) )
RETURN base * rate
```

### 2) SCD Type 2（履歴属性）
- M: `EffectiveFrom/EffectiveTo` を持つディムを整備。
- DAX: 参照日時に有効な行を `USERELATIONSHIP` またはブリッジで選択。

```DAX
[Attr at OrderDate] = 
CALCULATE(
    SELECTEDVALUE('Dim Customer SCD2'[Attribute]),
    USERELATIONSHIP('Dim Customer SCD2'[EffectiveFrom], 'Fact Sales'[OrderDate])
)
```

### 3) 在庫スナップショット（M: 粒度整形 → DAX: Last Non Blank）
- M: 日次スナップショットに正規化。
- DAX: `LASTNONBLANKVALUE` で日別残高を選択し、ロールアップ。

---

## トラブルシュート／アンチパターン／テスト

### よくあるアンチパターン
- 事実テーブル間の直接リレーション（ブリッジ/ディムを介す）。
- 全面双方向リレーション（必要な箇所のみ一時的に）。
- 計算列での集計（メジャーにする）。
- 無目的な `ALLSELECTED` 使用。
- 広いテーブルをそのまま DirectQuery（列選択・集約で削減）。

### テストと検証
- DAX Studio: サーバータイミング、クエリプラン、VertiPaq Analyzer。
- Performance Analyzer（デスクトップ）: 可視化別の負荷を把握。
- パラメータ化して What-If 検証。

### デバッグの定番メジャー
```DAX
[Row Count] = COUNTROWS(VALUES('Fact Sales'[SalesOrderID]))

[Current Filters] = 
CONCATENATEX(
    SUMMARIZECOLUMNS(
        'Dim Product'[Category], 'Dim Date'[Year]
    ),
    'Dim Product'[Category] & ":" & 'Dim Date'[Year],
    ", "
)
```

---

## チェックリストと再利用テンプレート

### DAX チェックリスト
- Star Schema か（スノーフレークは必要最小限）。
- Auto Date/Time 無効。`Dim Date` で管理。
- すべての派生はベースメジャー参照に。
- 時系列は `Date` 列へ正しくリレーション。
- `VAR` で重計算を抑制。`DIVIDE` でゼロ割回避。
- フィルター解除は最小範囲に限定。

### M チェックリスト
- 早期フィルター・列選択でフォールディング維持。
- 参照設計で前処理共通化。不要クエリはロード無効。
- 型は明示。ロケール（"ja-JP"）注意。
- API は `RelativePath`/`Query` でキャッシュ安定化。
- インクリメンタルは `RangeStart/RangeEnd` + フォールディング確認。

### DAX テンプレート（貼り替え用）
```DAX
-- Base Measures
[Total Amount] = SUM('Fact'[Amount])
[Total Qty]    = SUM('Fact'[Quantity])

-- Time Intelligence
[Amount PY]    = CALCULATE([Total Amount], DATEADD('Dim Date'[Date], -1, YEAR))
[Amount YoY %] = VAR py = [Amount PY] RETURN DIVIDE([Total Amount]-py, py)

-- Share within Selection
[Share %] = DIVIDE([Total Amount], CALCULATE([Total Amount], ALLSELECTED('Dim X')))

-- Inactive Relationship
[By Alt Date] = CALCULATE([Total Amount], USERELATIONSHIP('Dim Date'[Date], 'Fact'[AltDate]))
```

### M テンプレート（REST ページネーション）
```powerquery
let
  GetPage = (page as number) as table =>
    let
      Response = Web.Contents("https://api.example.com", [RelativePath="v1/items", Query=[page=Text.From(page), size="100"]]),
      Json = Json.Document(Response),
      Items = Json[items],
      Tbl = Table.FromList(Items, Splitter.SplitByNothing(), null, null, ExtraValues.Error)
    in
      Tbl,

  Pages = List.Generate(
    () => [p=1, lastCount=100],
    each [lastCount] = 100,
    each [ p = [p] + 1, lastCount = Table.RowCount(GetPage([p])) ],
    each GetPage([p])
  ),

  Combined = Table.Combine(Pages)
in
  Combined
```

---

必要に応じて、具体的なデータモデルや可視化要件に合わせて各レシピをカスタマイズしてください。実装時は常に計測（DAX Studio/Performance Analyzer）を併用し、根拠を持った最適化を行うことを推奨します。

