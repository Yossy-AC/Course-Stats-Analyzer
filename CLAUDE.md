Role: Dedicated engineer and assistant for a university prep school English teacher.
Style: Conclusion first, concise, direct, no token waste.
Prohibited: greetings, prefaces, apologies, emojis/kaomoji.

---

## Project Overview
月次受講人数集計CLI。lists フォルダ内の複数 Excel ファイル（生徒受講情報）を自動処理し、
講座別・月別の受講人数を集計して Pivot 形式 Excel で出力する。

**Status**: ✅ Production Ready (CLI/Skill版, 2025年度 4月～2月 稼働確認済み)

## Tech Stack
- CLI Framework: Python 3.14+
- Package manager: uv
- Excel処理: pandas + openpyxl
- スキル化: Claude Code Skill

## Project Structure
```
Course-Stats-Analyzer/
├── scripts/
│   └── aggregate.py       # CLI メイン実行スクリプト
├── services/
│   └── aggregator.py      # 集計コアロジック
├── .claude/skills/
│   └── aggregate-enrollment/
│       └── SKILL.md       # Claude Code Skill定義
├── outputs/
│   ├── results/           # 月別 CSV 保存（{YYYY-MM}.csv）
│   └── monthly_stats.xlsx # 最終出力 Excel（Pivot形式）
├── lists/                 # 入力 Excel ファイル（*_YYMM.xlsx）
├── pyproject.toml         # 依存関係定義
└── CLAUDE.md              # このファイル
```

## Run
```bash
# CLIで直接実行
python scripts/aggregate.py

# または Claude Code Skill として
/aggregate-enrollment
```

---

## Input File Format
**ファイル名パターン**: `*_YYMM.xlsx`
- YY: 年度下2桁（例：25 = 2025年度）
- MM: 月番号（01～12）
- 例：`〔定例報告〕2025AC受講者ﾘｽﾄ_2504.xlsx` → 2025年4月

**Fiscal Year 2025 (2025年度) Mapping**:
- _2504～_2512 → 2025-04 ～ 2025-12
- _2501～_2503 → 2026-01 ～ 2026-03
（詳細は `parse_target_month()` を参照）

## Excel Column Mapping (確定版)
Excelファイルの列構造（Row 4がヘッダー、0-indexed）

```python
COLUMN_INDICES = {
    "add_date": 2,      # Column C: 受講追加日付
    "cancel_date": 6,   # Column G: 受講取消日付
    "course": 9,        # Column J: 講座名
    "class_type": 10,   # Column K: 【マスター】【コア】
    "classroom": 11,    # Column L: 受講教室
    "grade": 15,        # Column P: 学年コード (31=高1, 32=高2, 33=高3)
    "teacher": 26,      # Column AA: 担当
    "gender": 17,       # Column R: 性別コード (1=男, 2=女)
    "school": 18,       # Column S: 在籍校
    "department": 28,   # Column AC: 学科
}
```

## 集計仕様 (Confirmed & Production)
- **集計対象**: 高1/高2/高3のみ（学年コード31/32/33）
- **基準日**: `target_month` の前月末日
  - 例：2025-05 → 2025-04-30
  - 判定式：`add_date <= cutoff AND (cancel_date IS NULL OR cancel_date > cutoff)`
- **出力形式**: Pivot（固定列5 + 月列11）
- **固定列順序**: 学年 | 教室 | 講座名 | マスター/コア | 担当
- **月列順序**: 4月～3月（会計年度順、実装値は `MONTH_ORDER` 参照）
- **出力場所**:
  - CSV: `outputs/results/{YYYY-MM}.csv`（1ファイル/月）
  - Excel: `outputs/monthly_stats.xlsx`（全月をマージ）

---

## Core Functions (services/aggregator.py)

### `parse_target_month(filename: str) -> pd.Period | None`
ファイル名から対象月を抽出。**任意年度対応の年号変換ロジック。**

```python
# 例
parse_target_month("file_2504.xlsx")  # → 2025-04（Period）
parse_target_month("file_2501.xlsx")  # → 2026-01（Period）
parse_target_month("file_2602.xlsx")  # → 2027-02（Period）
```

ロジック: `year = fiscal_year if month >= 4 else fiscal_year + 1` で、会計年度(4月開始)から暦年に変換。

### `load_excel(file: bytes | Path) -> pd.DataFrame`
Excel ファイルを読み込み。Row 4（header=3）をヘッダーとして使用。

### `aggregate(df: pd.DataFrame, target_month: pd.Period | None = None) -> pd.DataFrame`
対象月1ヶ月分の受講人数を集計。Pivot形式で返す。

- グループ化軸（5次元）: 学年, 教室, 講座名, マスター/コア, 担当
- 返り値: `KEY_COLS + [月名]` の DataFrame
- 講座名解決: `str.contains()` + `where()` で【マスター】【コア】をサフィックスとして追加

### `save_monthly_result(df: pd.DataFrame, target_month: pd.Period, results_dir: Path = RESULTS_DIR) -> None`
集計結果を CSV で保存。同月を再アップロードすると上書きされる。

### `build_pivot(results_dir: Path = RESULTS_DIR) -> pd.DataFrame`
`outputs/results/` に保存されている全月の CSV をマージして Pivot形式 DataFrame を生成。
月の並び順は `MONTH_ORDER` に従う。

### `to_excel_bytes(result: pd.DataFrame) -> bytes`
DataFrame を Excel bytes に変換。シート名「月次受講人数」で出力。

---

## Process Flow (CLI実行フロー)

`python scripts/aggregate.py` または `/aggregate-enrollment` スキル実行時：

1. **File Discovery** (`lists/` ディレクトリ走査)
   - glob で `*.xlsx` をリストアップ
   - `parse_target_month()` で各ファイルから _YYMM を抽出

2. **Clean Results Directory**
   - `outputs/results/` 配下の既存 CSV を全削除（再実行時のクリーンアップ）

3. **Per-File Processing**（各ファイルループ）
   - `load_excel()` で DataFrame に読み込み
   - `aggregate(df, target_month)` で月別人数を集計
   - `save_monthly_result()` で `outputs/results/{YYYY-MM}.csv` に保存

4. **Pivot Generation & Output**
   - `build_pivot()` が全月 CSV をマージして Pivot DataFrame を生成
   - `to_excel_bytes()` で `outputs/monthly_stats.xlsx` に出力

5. **Console Report**
   - 月別統計をコンソール出力
   - 処理完了を通知

---

## Testing

テストを実行する場合：

```bash
# テスト実行（pytest 必要）
uv run pytest tests/

# テストカバレッジ
uv run pytest tests/ --cov=services
```

テストスイート (`tests/test_aggregator.py`) は以下を検証：
- `parse_target_month()`: ファイル名パース、年度変換、エッジケース
- `aggregate()`: アクティブ行フィルタ、学年フィルタ、担当フィルタ、講座名解決
- `build_pivot()`: 月のマージ、月順序の保持

---

## 実装の最適化履歴

### v2 (2026年2月)
- **パフォーマンス最適化**: `df.apply(axis=1)` の廃止 → 完全ベクトル化
  - 集計時間 13s → 5.12s に短縮
  - 講座名解決: `str.contains()` + `where()` で置換
  - Pivot生成: 逐次 merge → `pd.concat()` + `groupby().sum()` に変更
- **`parse_target_month()` 汎用化**: year_suffix=25 のハードコード廃止 → 任意年度対応
  - 会計年度ロジック: `year = fiscal_year if month >= 4 else fiscal_year + 1`
  - 2026年度以降のファイル対応可能
- **分析機能**: 削除（要件変更により月次集計のみに特化）
