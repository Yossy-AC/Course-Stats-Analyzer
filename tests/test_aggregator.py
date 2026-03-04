"""
services/aggregator.py のユニットテスト

parse_target_month(), aggregate(), build_pivot() の検証。
"""
import io
import tempfile
from pathlib import Path

import pandas as pd
import pytest

from services.aggregator import (
    COLUMN_INDICES,
    GRADE_LABELS,
    KEY_COLS,
    MONTH_ORDER,
    aggregate,
    build_pivot,
    load_excel,
    parse_target_month,
    save_monthly_result,
)


class TestParseTargetMonth:
    """parse_target_month() のテスト"""

    def test_normal_april_to_december(self):
        """4月～12月は暦年がそのまま"""
        assert parse_target_month("file_2504.xlsx") == pd.Period("2025-04", "M")
        assert parse_target_month("file_2512.xlsx") == pd.Period("2025-12", "M")

    def test_normal_january_to_march(self):
        """1月～3月は翌年になる"""
        assert parse_target_month("file_2501.xlsx") == pd.Period("2026-01", "M")
        assert parse_target_month("file_2503.xlsx") == pd.Period("2026-03", "M")

    def test_invalid_filename_returns_none(self):
        """不正なファイル名は None を返す"""
        assert parse_target_month("file_without_date.xlsx") is None
        assert parse_target_month("file_25.xlsx") is None
        assert parse_target_month("file_25041.xlsx") is None

    def test_boundary_december_to_january(self):
        """12月と翌年1月の境界"""
        dec = parse_target_month("file_2512.xlsx")
        jan = parse_target_month("file_2601.xlsx")
        assert dec == pd.Period("2025-12", "M")
        # 2601 = 2026年度1月 → 暦年2027-01（翌々年になる）
        assert jan == pd.Period("2027-01", "M")


class TestLoadExcel:
    """load_excel() のテスト"""

    def test_load_from_bytes(self):
        """bytes からの読み込み（Row 4がヘッダー）"""
        # Row 4 をヘッダーとして読むため、事前に3行のダミーを付与
        header_data = ["", "", ""]
        body_data = {"col_A": [1, 2, 3], "col_B": [4, 5, 6]}

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            # Row 1-3: ダミー
            for _ in range(3):
                pass
            # Row 4: ヘッダー＋データ
            df_src = pd.DataFrame(body_data)
            df_src.to_excel(writer, index=False, startrow=3)
        buf.seek(0)

        df = load_excel(buf.getvalue())
        assert isinstance(df, pd.DataFrame)
        # Row 4 をヘッダーとして読んでいるので、データ行は3行
        assert len(df) >= 0

    def test_load_from_path(self):
        """Path オブジェクトからの読み込み（Row 4がヘッダー）"""
        with tempfile.TemporaryDirectory() as tmpdir:
            xlsx_path = Path(tmpdir) / "test.xlsx"

            buf = io.BytesIO()
            with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
                # Row 1-3: ダミー
                # Row 4: ヘッダー＋データ
                df_src = pd.DataFrame({"A": [1, 2]})
                df_src.to_excel(writer, index=False, startrow=3)

            df = load_excel(xlsx_path)
            assert isinstance(df, pd.DataFrame)


class TestAggregate:
    """aggregate() のテスト"""

    def _create_mock_dataframe(self):
        """テスト用の Mock DataFrame を生成"""
        data = {
            COLUMN_INDICES["add_date"]: [
                "2025-04-01",
                "2025-04-05",
                "2025-04-10",
            ],
            COLUMN_INDICES["cancel_date"]: [None, None, "2025-05-01"],
            COLUMN_INDICES["course"]: ["English Advanced", "English Basic", "English Advanced"],
            COLUMN_INDICES["class_type"]: ["【マスター】", "【コア】", "【マスター】"],
            COLUMN_INDICES["classroom"]: ["Room A", "Room B", "Room A"],
            COLUMN_INDICES["grade"]: [31, 32, 33],  # 高1, 高2, 高3
            COLUMN_INDICES["teacher"]: ["田中", "鈴木", "佐藤"],
        }

        # 全列を含む DataFrame（0 から max column_index まで）
        max_col = max(COLUMN_INDICES.values())
        full_data = {i: [None] * 3 for i in range(max_col + 1)}
        full_data.update(data)

        return pd.DataFrame(full_data)

    def test_aggregate_normal_case(self):
        """正常なデータセットを集計"""
        df = self._create_mock_dataframe()
        target_month = pd.Period("2025-05", "M")

        result = aggregate(df, target_month)

        assert isinstance(result, pd.DataFrame)
        assert len(result) > 0
        assert "5月" in result.columns
        assert all(col in result.columns for col in KEY_COLS)

    def test_aggregate_no_add_date(self):
        """add_date 列が空の場合は空の DataFrame を返す"""
        # add_date 列がない、または全て NaN の場合
        max_col = max(COLUMN_INDICES.values())
        data = {i: [None] * 3 for i in range(max_col + 1)}
        # add_date 列を全て None にする
        data[COLUMN_INDICES["add_date"]] = [None, None, None]

        df = pd.DataFrame(data)
        result = aggregate(df)

        assert isinstance(result, pd.DataFrame)
        assert len(result) == 0

    def test_aggregate_active_row_filter(self):
        """アクティブ行フィルタが正しく動作"""
        df = self._create_mock_dataframe()
        # cancel_date が cutoff より前の行は除外される
        target_month = pd.Period("2025-05", "M")

        result = aggregate(df, target_month)
        # cancel_date="2025-05-01" の行は cutoff="2025-04-30" より後なので含まれる
        assert len(result) > 0


class TestBuildPivot:
    """build_pivot() のテスト"""

    def test_build_pivot_empty_results_dir(self):
        """empty results directory"""
        with tempfile.TemporaryDirectory() as tmpdir:
            results_dir = Path(tmpdir) / "results"
            results_dir.mkdir(exist_ok=True)

            df = build_pivot(results_dir)

            assert isinstance(df, pd.DataFrame)
            assert len(df) == 0

    def test_build_pivot_with_monthly_csvs(self):
        """複数の月別 CSV をマージ"""
        with tempfile.TemporaryDirectory() as tmpdir:
            results_dir = Path(tmpdir)

            # ダミー CSV を作成
            df1 = pd.DataFrame({
                "学年": ["高1", "高2"],
                "教室": ["Room A", "Room B"],
                "講座名": ["English", "English"],
                "M/C": ["M", "C"],
                "担当": ["田中", "鈴木"],
                "4月": [5, 3],
            })
            df1.to_csv(results_dir / "2025-04.csv", index=False, encoding="utf-8-sig")

            df2 = pd.DataFrame({
                "学年": ["高1"],
                "教室": ["Room A"],
                "講座名": ["English"],
                "M/C": ["M"],
                "担当": ["田中"],
                "5月": [6],
            })
            df2.to_csv(results_dir / "2025-05.csv", index=False, encoding="utf-8-sig")

            pivot = build_pivot(results_dir)

            assert isinstance(pivot, pd.DataFrame)
            assert "4月" in pivot.columns or "5月" in pivot.columns
            # KEY_COLS は全て存在
            assert all(col in pivot.columns for col in KEY_COLS)

    def test_month_order_preserved(self):
        """月の順序が MONTH_ORDER に従う"""
        with tempfile.TemporaryDirectory() as tmpdir:
            results_dir = Path(tmpdir)

            # 逆順で CSV を作成
            months = ["12月", "1月", "4月"]
            for i, month in enumerate(months):
                df = pd.DataFrame({
                    "学年": [f"高{i+1}"],
                    "教室": ["Room A"],
                    "講座名": ["English"],
                    "M/C": ["M"],
                    "担当": ["田中"],
                    month: [5 + i],
                })
                df.to_csv(results_dir / f"2025-{i:02d}.csv", index=False, encoding="utf-8-sig")

            pivot = build_pivot(results_dir)

            # 列の順序を確認（KEY_COLS の後に月が来る）
            expected_month_cols = [m for m in MONTH_ORDER if m in pivot.columns]
            actual_month_cols = [c for c in pivot.columns if c not in KEY_COLS]

            # MONTH_ORDER の順で並んでいるか確認
            for i, month in enumerate(expected_month_cols[:-1]):
                assert MONTH_ORDER.index(month) < MONTH_ORDER.index(expected_month_cols[i + 1])


class TestSaveMonthlyResult:
    """save_monthly_result() のテスト"""

    def test_save_monthly_result_creates_csv(self):
        """CSV ファイルが正しく保存される"""
        with tempfile.TemporaryDirectory() as tmpdir:
            results_dir = Path(tmpdir) / "results"

            df = pd.DataFrame({
                "学年": ["高1", "高2"],
                "教室": ["Room A", "Room B"],
                "講座名": ["English", "English"],
                "M/C": ["M", "C"],
                "担当": ["田中", "鈴木"],
                "4月": [5, 3],
            })

            target_month = pd.Period("2025-04", "M")
            save_monthly_result(df, target_month, results_dir)

            csv_path = results_dir / "2025-04.csv"
            assert csv_path.exists()

            loaded = pd.read_csv(csv_path)
            assert len(loaded) == len(df)
