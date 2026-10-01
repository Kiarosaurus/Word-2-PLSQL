from __future__ import annotations

import unittest

from corporate_report_compiler.diagnostics import Diagnostics
from corporate_report_compiler.sql_reader import MAX_SQL_LENGTH, validate_sql


class SqlReaderTests(unittest.TestCase):
    def validate(self, sql: str) -> tuple[Diagnostics, tuple[str, ...]]:
        diagnostics = Diagnostics()
        binds = validate_sql(sql, diagnostics)
        return diagnostics, binds

    def test_select_and_with_are_accepted(self) -> None:
        for sql in (
            "select :dni as dni from dual",
            "-- comment\nwith x as (select 1 n from dual) select n from x",
        ):
            with self.subTest(sql=sql):
                diagnostics, _ = self.validate(sql)
                self.assertFalse(diagnostics.has_errors)

    def test_binds_are_unique_normalized_and_keep_first_order(self) -> None:
        diagnostics, binds = self.validate(
            "select * from t where a=:dni or b=:AREA or c=:dni"
        )
        self.assertFalse(diagnostics.has_errors)
        self.assertEqual(binds, ("DNI", "AREA"))

    def test_literals_comments_quoted_identifiers_and_double_colon_are_ignored(self) -> None:
        sql = """
            select ':NOT_A_BIND' as text_value,
                   "A:QUOTED" as quoted_value,
                   value::varchar2 as cast_value
              from report_source
             where id = :REAL_BIND
               -- :LINE_COMMENT
               /* :BLOCK_COMMENT */
               and note = 'it''s :STILL_TEXT'
        """
        diagnostics, binds = self.validate(sql)
        self.assertFalse(diagnostics.has_errors)
        self.assertEqual(binds, ("REAL_BIND",))

    def test_semicolon_in_literal_or_comment_is_not_a_second_statement(self) -> None:
        diagnostics, _ = self.validate("select ';' v from dual -- ;\n")
        self.assertFalse(diagnostics.has_errors)

    def test_oracle_q_quoted_literals_are_ignored_completely(self) -> None:
        cases = (
            "select q'[text ; :FALSE_BIND '' braces ]' value from dual where id=:REAL_BIND",
            "select q'{text ; :FALSE_BIND }' value from dual where id=:REAL_BIND",
            "select q'<text ; :FALSE_BIND >' value from dual where id=:REAL_BIND",
            "select q'!text ; :FALSE_BIND !' value from dual where id=:REAL_BIND",
        )
        for sql in cases:
            with self.subTest(sql=sql):
                diagnostics, binds = self.validate(sql)
                self.assertFalse(diagnostics.has_errors, diagnostics.to_dict())
                self.assertEqual(binds, ("REAL_BIND",))

    def test_non_query_and_multiple_statement_are_rejected(self) -> None:
        cases = (
            "delete from t",
            "begin null; end;",
            "select 1 from dual; delete from t",
            "",
        )
        for sql in cases:
            with self.subTest(sql=sql):
                diagnostics, _ = self.validate(sql)
                self.assertTrue(diagnostics.has_errors)

    def test_query_keyword_must_be_the_first_token(self) -> None:
        diagnostics, _ = self.validate(") select 1 from dual")
        self.assertIn("SQL-006", {item.code for item in diagnostics.items})

    def test_overlong_bind_is_rejected_instead_of_truncated(self) -> None:
        diagnostics, binds = self.validate(
            "select 1 from dual where x = :" + "A" * 31
        )
        self.assertEqual(binds, ())
        self.assertIn("SQL-011", {item.code for item in diagnostics.items})

    def test_positional_and_quoted_bind_syntax_is_rejected(self) -> None:
        for sql in ('select :1 as value from dual', 'select :"X" as value from dual'):
            with self.subTest(sql=sql):
                diagnostics, binds = self.validate(sql)
                self.assertEqual(binds, ())
                self.assertIn("SQL-012", {item.code for item in diagnostics.items})

    def test_unclosed_literal_identifier_and_comment_are_rejected(self) -> None:
        for sql in ("select 'oops from dual", 'select "oops from dual', "select 1 /* oops"):
            with self.subTest(sql=sql):
                diagnostics, _ = self.validate(sql)
                self.assertTrue(diagnostics.has_errors)

    def test_sql_length_limit(self) -> None:
        diagnostics, _ = self.validate("select 1 from dual " + " " * MAX_SQL_LENGTH)
        self.assertTrue(diagnostics.has_errors)
        self.assertIn("SQL-002", {item.code for item in diagnostics.items})

    def test_sql_length_limit_counts_utf8_bytes_not_characters(self) -> None:
        sql = "select '" + ("ñ" * 16_380) + "' as texto from dual"
        self.assertLess(len(sql), 32_767)
        self.assertGreater(len(sql.encode("utf-8")), 32_767)
        diagnostics, _binds = self.validate(sql)
        self.assertIn("SQL-002", {item.code for item in diagnostics.items})

    def test_sqlplus_slash_is_rejected(self) -> None:
        diagnostics, _ = self.validate("select 1 from dual\n/\n")
        self.assertTrue(
            diagnostics.has_errors,
            "El contrato excluye terminadores SQL*Plus; una línea '/' debe rechazarse.",
        )

    def test_nul_is_rejected(self) -> None:
        diagnostics, _ = self.validate("select 1 from dual\x00")
        self.assertTrue(diagnostics.has_errors, "El SQL con NUL no debe llegar a Oracle.")


if __name__ == "__main__":
    unittest.main()
