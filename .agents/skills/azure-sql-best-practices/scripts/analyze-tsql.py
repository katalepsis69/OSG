#!/usr/bin/env python3
"""
Azure SQL Best Practices - T-SQL Analyzer

This script analyzes T-SQL files for common best practice violations,
including all Microsoft SSDT Code Analysis rules (SR0001-SR0016).

Usage:
    python analyze-tsql.py <file_or_directory>
    python analyze-tsql.py --all  # Analyze all .sql files in current directory
"""

import re
import sys
import os
from pathlib import Path
from dataclasses import dataclass
from typing import List, Tuple
from enum import Enum

class Severity(Enum):
    CRITICAL = "CRITICAL"
    HIGH = "HIGH"
    MEDIUM = "MEDIUM"
    LOW = "LOW"

@dataclass
class Violation:
    rule_id: str
    severity: Severity
    line_number: int
    line_content: str
    message: str
    suggestion: str

class TSQLAnalyzer:
    """Analyzes T-SQL code for best practice violations including SSDT rules."""
    
    def __init__(self):
        self.violations: List[Violation] = []
        
    def analyze_file(self, filepath: str) -> List[Violation]:
        """Analyze a single T-SQL file."""
        self.violations = []
        
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
            lines = content.split('\n')
        
        # Run all checks (organized by SSDT rule ID where applicable)
        
        # SSDT Design Issues
        self._check_select_star(lines)                    # SR0001
        self._check_scope_identity(lines)                  # SR0008
        self._check_small_varchar(lines)                   # SR0009
        self._check_deprecated_join_syntax(lines)          # SR0010
        self._check_output_params(content, lines)          # SR0013
        self._check_implicit_conversion(lines)             # SR0014
        
        # SSDT Performance Issues
        self._check_in_predicate_indexes(lines)            # SR0004
        self._check_leading_wildcard(lines)                # SR0005
        self._check_column_both_sides(lines)               # SR0006
        self._check_isnull_on_nullable(lines)              # SR0007
        self._check_deterministic_functions(lines)         # SR0015
        
        # SSDT Naming Issues
        self._check_special_characters(lines)              # SR0011
        self._check_reserved_words(lines)                  # SR0012
        self._check_sp_prefix(lines)                       # SR0016
        
        # Additional best practice checks
        self._check_nocount(content, lines)
        self._check_dynamic_sql_injection(lines)
        self._check_cursor_usage(lines)
        self._check_schema_qualification(lines)
        self._check_functions_on_columns(lines)
        self._check_nolock_hints(lines)
        self._check_deprecated_syntax(lines)
        self._check_transaction_handling(content, lines)
        
        return self.violations
    
    # =========================================================================
    # SSDT Design Issues (SR0001, SR0008-SR0014)
    # =========================================================================
    
    def _check_select_star(self, lines: List[str]):
        """SR0001: Check for SELECT * usage (except in EXISTS clauses)."""
        pattern = re.compile(r'\bSELECT\s+\*\s+FROM\b', re.IGNORECASE)
        exists_pattern = re.compile(r'\bEXISTS\s*\(\s*SELECT\s+\*', re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            if pattern.search(line) and not exists_pattern.search(line):
                self.violations.append(Violation(
                    rule_id="SR0001",
                    severity=Severity.HIGH,
                    line_number=i,
                    line_content=line.strip(),
                    message="SELECT * found - explicitly list required columns",
                    suggestion="Replace SELECT * with specific column names to improve performance and maintainability"
                ))
    
    def _check_scope_identity(self, lines: List[str]):
        """SR0008: Check for @@IDENTITY usage instead of SCOPE_IDENTITY()."""
        pattern = re.compile(r'@@IDENTITY\b', re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            if pattern.search(line):
                self.violations.append(Violation(
                    rule_id="SR0008",
                    severity=Severity.MEDIUM,
                    line_number=i,
                    line_content=line.strip(),
                    message="@@IDENTITY can return wrong value if triggers insert into other tables",
                    suggestion="Use SCOPE_IDENTITY() instead of @@IDENTITY for reliable identity retrieval"
                ))
    
    def _check_small_varchar(self, lines: List[str]):
        """SR0009: Check for VARCHAR(1) or VARCHAR(2) - should use CHAR."""
        patterns = [
            (r'\bVARCHAR\s*\(\s*1\s*\)', "VARCHAR(1)"),
            (r'\bVARCHAR\s*\(\s*2\s*\)', "VARCHAR(2)"),
            (r'\bNVARCHAR\s*\(\s*1\s*\)', "NVARCHAR(1)"),
            (r'\bNVARCHAR\s*\(\s*2\s*\)', "NVARCHAR(2)"),
        ]
        
        for i, line in enumerate(lines, 1):
            for pattern, type_name in patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    self.violations.append(Violation(
                        rule_id="SR0009",
                        severity=Severity.LOW,
                        line_number=i,
                        line_content=line.strip(),
                        message=f"{type_name} found - variable-length type with tiny size is inefficient",
                        suggestion=f"Use CHAR or NCHAR instead of {type_name} for fixed small sizes"
                    ))
    
    def _check_deprecated_join_syntax(self, lines: List[str]):
        """SR0010: Check for deprecated *= and =* join syntax."""
        pattern = re.compile(r'\*=|=\*')
        
        for i, line in enumerate(lines, 1):
            if pattern.search(line):
                self.violations.append(Violation(
                    rule_id="SR0010",
                    severity=Severity.MEDIUM,
                    line_number=i,
                    line_content=line.strip(),
                    message="Deprecated outer join syntax (*= or =*) found",
                    suggestion="Use ANSI SQL JOIN syntax: LEFT OUTER JOIN or RIGHT OUTER JOIN"
                ))
    
    def _check_output_params(self, content: str, lines: List[str]):
        """SR0013: Check for output parameters that may not be set in all code paths."""
        # Find procedures with OUTPUT parameters
        proc_pattern = re.compile(r'@(\w+)\s+\w+.*\bOUTPUT\b', re.IGNORECASE)
        if_pattern = re.compile(r'\bIF\b', re.IGNORECASE)
        else_pattern = re.compile(r'\bELSE\b', re.IGNORECASE)
        
        output_params = proc_pattern.findall(content)
        has_if = if_pattern.search(content)
        has_else = else_pattern.search(content)
        
        # Simple heuristic: if there's IF without ELSE and OUTPUT params, flag it
        if output_params and has_if and not has_else:
            for i, line in enumerate(lines, 1):
                if proc_pattern.search(line):
                    self.violations.append(Violation(
                        rule_id="SR0013",
                        severity=Severity.MEDIUM,
                        line_number=i,
                        line_content=line.strip(),
                        message=f"Output parameter may not be set in all code paths",
                        suggestion="Ensure OUTPUT parameters are set in all branches (IF/ELSE) or initialize with default"
                    ))
                    break
    
    def _check_implicit_conversion(self, lines: List[str]):
        """SR0014: Check for potential implicit type conversions."""
        # Check for common patterns that suggest type mismatch
        patterns = [
            (r'@\w+\s+VARCHAR\s*\([^)]+\).*--.*NVARCHAR', "VARCHAR parameter for NVARCHAR column"),
            (r'\bCAST\s*\([^)]+\s+AS\s+(TINY|SMALL)INT\s*\)', "Cast to smaller integer type"),
            (r'\bCONVERT\s*\(\s*(TINY|SMALL)INT\s*,', "Convert to smaller integer type"),
            (r'SET\s+@\w+\s*=\s*@\w+\s*;?\s*$', "Direct assignment between potentially different types"),
        ]
        
        for i, line in enumerate(lines, 1):
            # Look for BIGINT to SMALLINT or similar patterns
            if re.search(r'\b(BIGINT|INT)\b.*\b(SMALLINT|TINYINT)\b', line, re.IGNORECASE):
                self.violations.append(Violation(
                    rule_id="SR0014",
                    severity=Severity.HIGH,
                    line_number=i,
                    line_content=line.strip(),
                    message="Potential data loss from implicit conversion to smaller type",
                    suggestion="Use explicit CAST/CONVERT with range validation to prevent data loss"
                ))
            # Look for FLOAT to INT conversions
            if re.search(r'\bFLOAT\b.*\b(INT|DECIMAL)\b', line, re.IGNORECASE):
                self.violations.append(Violation(
                    rule_id="SR0014",
                    severity=Severity.HIGH,
                    line_number=i,
                    line_content=line.strip(),
                    message="Potential data loss from FLOAT to integer/decimal conversion",
                    suggestion="Use ROUND(), FLOOR(), or CEILING() before casting FLOAT to integer types"
                ))
    
    # =========================================================================
    # SSDT Performance Issues (SR0004-SR0007, SR0015)
    # =========================================================================
    
    def _check_in_predicate_indexes(self, lines: List[str]):
        """SR0004: Flag IN predicates that may use non-indexed columns."""
        pattern = re.compile(r'\bIN\s*\([^)]+\)', re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            if pattern.search(line) and 'WHERE' in line.upper():
                self.violations.append(Violation(
                    rule_id="SR0004",
                    severity=Severity.MEDIUM,
                    line_number=i,
                    line_content=line.strip(),
                    message="IN predicate found - verify column is indexed",
                    suggestion="Ensure columns used in IN predicates have appropriate indexes"
                ))
    
    def _check_leading_wildcard(self, lines: List[str]):
        """SR0005: Check for LIKE patterns starting with '%'."""
        pattern = re.compile(r"LIKE\s+[N]?'%", re.IGNORECASE)
        pattern2 = re.compile(r"LIKE\s+[N]?'%[^']*'", re.IGNORECASE)
        pattern3 = re.compile(r"LIKE\s+'%'\s*\+", re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            if pattern.search(line) or pattern3.search(line):
                self.violations.append(Violation(
                    rule_id="SR0005",
                    severity=Severity.HIGH,
                    line_number=i,
                    line_content=line.strip(),
                    message="LIKE pattern starting with '%' prevents index usage",
                    suggestion="Use full-text search or redesign query to avoid leading wildcards"
                ))
    
    def _check_column_both_sides(self, lines: List[str]):
        """SR0006: Check for column references on both sides of comparison."""
        # Look for arithmetic operations on columns in WHERE
        patterns = [
            r'\bWHERE\b.*\w+\s*[\+\-\*/]\s*\w+\s*[><=]',
            r'\bWHERE\b.*\w+\s*[><=]\s*\w+\s*[\+\-\*/]',
            r'\bAND\b.*\w+\s*[\+\-\*/]\s*\w+\s*[><=]',
        ]
        
        for i, line in enumerate(lines, 1):
            for pattern in patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    # Exclude obvious variable references
                    if not re.search(r'@\w+\s*[\+\-\*/]', line):
                        self.violations.append(Violation(
                            rule_id="SR0006",
                            severity=Severity.MEDIUM,
                            line_number=i,
                            line_content=line.strip(),
                            message="Column arithmetic in WHERE clause prevents index usage",
                            suggestion="Move column to one side: 'col / x > y' becomes 'col > y * x'"
                        ))
                        break
    
    def _check_isnull_on_nullable(self, lines: List[str]):
        """SR0007: Check for nullable columns used in expressions without ISNULL."""
        # Look for arithmetic with columns that might be nullable
        arithmetic_pattern = re.compile(r'\b\w+\s*[\+\-\*/]\s*\w+\s+AS\b', re.IGNORECASE)
        isnull_pattern = re.compile(r'\bISNULL\s*\(', re.IGNORECASE)
        coalesce_pattern = re.compile(r'\bCOALESCE\s*\(', re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            if arithmetic_pattern.search(line):
                if not isnull_pattern.search(line) and not coalesce_pattern.search(line):
                    self.violations.append(Violation(
                        rule_id="SR0007",
                        severity=Severity.MEDIUM,
                        line_number=i,
                        line_content=line.strip(),
                        message="Column arithmetic without ISNULL - NULL propagates through expression",
                        suggestion="Use ISNULL(column, default_value) or COALESCE() for nullable columns"
                    ))
    
    def _check_deterministic_functions(self, lines: List[str]):
        """SR0015: Check for deterministic function calls in WHERE that could be extracted."""
        pattern = re.compile(r'\bWHERE\b.*\bGETDATE\s*\(\s*\)', re.IGNORECASE)
        pattern2 = re.compile(r'\bAND\b.*\bGETDATE\s*\(\s*\)', re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            if pattern.search(line) or pattern2.search(line):
                self.violations.append(Violation(
                    rule_id="SR0015",
                    severity=Severity.MEDIUM,
                    line_number=i,
                    line_content=line.strip(),
                    message="GETDATE() in WHERE clause is evaluated per row",
                    suggestion="Extract GETDATE() to a variable before the query for better performance"
                ))
    
    # =========================================================================
    # SSDT Naming Issues (SR0011, SR0012, SR0016)
    # =========================================================================
    
    def _check_special_characters(self, lines: List[str]):
        """SR0011: Check for special characters in object names."""
        # Look for bracketed identifiers with special chars
        pattern = re.compile(r'\[[\w]*[-$#@!%^&*()+={}\[\]|\\:;"\'<>,?/][\w]*\]')
        
        for i, line in enumerate(lines, 1):
            if pattern.search(line):
                # Skip if it's a comment
                if line.strip().startswith('--'):
                    continue
                self.violations.append(Violation(
                    rule_id="SR0011",
                    severity=Severity.LOW,
                    line_number=i,
                    line_content=line.strip(),
                    message="Object name contains special characters requiring brackets",
                    suggestion="Use only letters, numbers, and underscores in object names"
                ))
    
    def _check_reserved_words(self, lines: List[str]):
        """SR0012: Check for reserved words used as identifiers."""
        reserved_words = [
            'SELECT', 'INSERT', 'UPDATE', 'DELETE', 'FROM', 'WHERE', 'ORDER',
            'TABLE', 'INDEX', 'VIEW', 'PROCEDURE', 'FUNCTION', 'TRIGGER',
            'DATABASE', 'SCHEMA', 'USER', 'ROLE', 'GRANT', 'REVOKE',
            'PRIMARY', 'FOREIGN', 'KEY', 'CONSTRAINT', 'DEFAULT', 'NULL',
            'AND', 'OR', 'NOT', 'IN', 'EXISTS', 'BETWEEN', 'LIKE', 'AS',
            'JOIN', 'LEFT', 'RIGHT', 'INNER', 'OUTER', 'CROSS', 'ON',
            'GROUP', 'HAVING', 'DISTINCT', 'TOP', 'PERCENT', 'WITH',
            'UNION', 'EXCEPT', 'INTERSECT', 'ALL', 'ANY', 'SOME',
        ]
        
        # Look for reserved words used as column aliases
        alias_pattern = re.compile(r'\bAS\s+\[(' + '|'.join(reserved_words) + r')\]', re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            match = alias_pattern.search(line)
            if match:
                self.violations.append(Violation(
                    rule_id="SR0012",
                    severity=Severity.MEDIUM,
                    line_number=i,
                    line_content=line.strip(),
                    message=f"Reserved word '{match.group(1)}' used as identifier",
                    suggestion="Avoid using SQL reserved words as column or object names"
                ))
    
    def _check_sp_prefix(self, lines: List[str]):
        """SR0016: Check for sp_ prefix on stored procedures."""
        pattern = re.compile(r'\bCREATE\s+(OR\s+ALTER\s+)?PROCEDURE\s+sp_', re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            if pattern.search(line):
                self.violations.append(Violation(
                    rule_id="SR0016",
                    severity=Severity.MEDIUM,
                    line_number=i,
                    line_content=line.strip(),
                    message="Stored procedure uses sp_ prefix which is reserved for system procedures",
                    suggestion="Use a different prefix (e.g., usp_, proc_) or no prefix at all"
                ))
    
    # =========================================================================
    # Additional Best Practice Checks
    # =========================================================================
    
    def _check_nocount(self, content: str, lines: List[str]):
        """Check for SET NOCOUNT ON in stored procedures."""
        proc_pattern = re.compile(r'\bCREATE\s+(OR\s+ALTER\s+)?PROCEDURE\b', re.IGNORECASE)
        nocount_pattern = re.compile(r'\bSET\s+NOCOUNT\s+ON\b', re.IGNORECASE)
        
        if proc_pattern.search(content) and not nocount_pattern.search(content):
            for i, line in enumerate(lines, 1):
                if proc_pattern.search(line):
                    self.violations.append(Violation(
                        rule_id="tsql-set-nocount",
                        severity=Severity.MEDIUM,
                        line_number=i,
                        line_content=line.strip(),
                        message="Stored procedure missing SET NOCOUNT ON",
                        suggestion="Add 'SET NOCOUNT ON;' at the beginning of the procedure body"
                    ))
                    break
    
    def _check_dynamic_sql_injection(self, lines: List[str]):
        """Check for potential SQL injection in dynamic SQL."""
        patterns = [
            (r"['\"].*\+\s*@\w+\s*\+.*['\"]", "String concatenation with variable"),
            (r"EXEC\s*\(\s*@?\w+\s*\)", "EXEC with variable (use sp_executesql)"),
            (r"'''\s*\+\s*@\w+\s*\+\s*'''", "Quotes around concatenated variable"),
        ]
        
        for i, line in enumerate(lines, 1):
            for pattern, desc in patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    self.violations.append(Violation(
                        rule_id="query-parameterize",
                        severity=Severity.CRITICAL,
                        line_number=i,
                        line_content=line.strip(),
                        message=f"Potential SQL injection vulnerability: {desc}",
                        suggestion="Use sp_executesql with parameters instead of string concatenation"
                    ))
    
    def _check_cursor_usage(self, lines: List[str]):
        """Check for cursor usage that could be replaced with set-based operations."""
        cursor_pattern = re.compile(r'\bDECLARE\s+\w+\s+CURSOR\b', re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            if cursor_pattern.search(line):
                self.violations.append(Violation(
                    rule_id="query-avoid-cursors",
                    severity=Severity.HIGH,
                    line_number=i,
                    line_content=line.strip(),
                    message="Cursor declaration found - consider set-based alternative",
                    suggestion="Replace cursor logic with set-based UPDATE/INSERT/DELETE operations"
                ))
    
    def _check_schema_qualification(self, lines: List[str]):
        """Check for unqualified object references."""
        unqualified = re.compile(r'\b(FROM|JOIN)\s+(?!\w+\.)\w+\b', re.IGNORECASE)
        qualified = re.compile(r'\b(FROM|JOIN)\s+\w+\.\w+', re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            if re.search(r'(#|@)\w+', line):
                continue
            if unqualified.search(line) and not qualified.search(line):
                if not re.search(r'(#|@|\.)\w+', line.split('FROM')[-1] if 'FROM' in line.upper() else line):
                    self.violations.append(Violation(
                        rule_id="tsql-schema-qualify",
                        severity=Severity.LOW,
                        line_number=i,
                        line_content=line.strip(),
                        message="Object reference may not be schema-qualified",
                        suggestion="Use schema-qualified names: dbo.TableName or schema.TableName"
                    ))
    
    def _check_functions_on_columns(self, lines: List[str]):
        """Check for functions applied to columns in WHERE clauses (non-SARGable)."""
        patterns = [
            (r'\bWHERE\b.*\bYEAR\s*\(\s*\w+\s*\)', "YEAR() function on column"),
            (r'\bWHERE\b.*\bMONTH\s*\(\s*\w+\s*\)', "MONTH() function on column"),
            (r'\bWHERE\b.*\bDATENAME\s*\(', "DATENAME() function on column"),
            (r'\bWHERE\b.*\bCONVERT\s*\(\s*\w+\s*,\s*\w+', "CONVERT() on column"),
            (r'\bWHERE\b.*\bCAST\s*\(\s*\w+\s+AS\b', "CAST() on column"),
            (r'\bWHERE\b.*\bLEFT\s*\(\s*\w+', "LEFT() on column"),
            (r'\bWHERE\b.*\bRIGHT\s*\(\s*\w+', "RIGHT() on column"),
            (r'\bWHERE\b.*\bSUBSTRING\s*\(\s*\w+', "SUBSTRING() on column"),
            (r'\bWHERE\b.*\bUPPER\s*\(\s*\w+', "UPPER() on column"),
            (r'\bWHERE\b.*\bLOWER\s*\(\s*\w+', "LOWER() on column"),
        ]
        
        for i, line in enumerate(lines, 1):
            for pattern, desc in patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    self.violations.append(Violation(
                        rule_id="query-sargable",
                        severity=Severity.HIGH,
                        line_number=i,
                        line_content=line.strip(),
                        message=f"Non-SARGable predicate: {desc} prevents index usage",
                        suggestion="Rewrite to avoid applying functions to columns in WHERE clause"
                    ))
    
    def _check_nolock_hints(self, lines: List[str]):
        """Check for NOLOCK hints."""
        nolock_pattern = re.compile(r'\bWITH\s*\(\s*NOLOCK\s*\)', re.IGNORECASE)
        
        for i, line in enumerate(lines, 1):
            if nolock_pattern.search(line):
                self.violations.append(Violation(
                    rule_id="tsql-avoid-hints",
                    severity=Severity.MEDIUM,
                    line_number=i,
                    line_content=line.strip(),
                    message="NOLOCK hint found - can cause dirty reads",
                    suggestion="Consider Read Committed Snapshot Isolation (RCSI) or proper transaction isolation levels"
                ))
    
    def _check_deprecated_syntax(self, lines: List[str]):
        """Check for deprecated T-SQL syntax."""
        patterns = [
            (r'\bSET\s+ROWCOUNT\s+\d+', "SET ROWCOUNT is deprecated", "Use TOP clause instead"),
            (r'\bTEXT\b\s+(?!NOT\s+NULL|NULL)', "TEXT data type is deprecated", "Use VARCHAR(MAX)"),
            (r'\bNTEXT\b', "NTEXT data type is deprecated", "Use NVARCHAR(MAX)"),
            (r'\bIMAGE\b\s+(?!NOT\s+NULL|NULL)', "IMAGE data type is deprecated", "Use VARBINARY(MAX)"),
        ]
        
        for i, line in enumerate(lines, 1):
            for pattern, msg, suggestion in patterns:
                if re.search(pattern, line, re.IGNORECASE):
                    self.violations.append(Violation(
                        rule_id="tsql-deprecated",
                        severity=Severity.MEDIUM,
                        line_number=i,
                        line_content=line.strip(),
                        message=msg,
                        suggestion=suggestion
                    ))
    
    def _check_transaction_handling(self, content: str, lines: List[str]):
        """Check for proper transaction handling."""
        has_begin_tran = re.search(r'\bBEGIN\s+TRAN(SACTION)?\b', content, re.IGNORECASE)
        has_try_catch = re.search(r'\bBEGIN\s+TRY\b', content, re.IGNORECASE)
        has_xact_abort = re.search(r'\bSET\s+XACT_ABORT\s+ON\b', content, re.IGNORECASE)
        
        if has_begin_tran and not has_try_catch:
            for i, line in enumerate(lines, 1):
                if re.search(r'\bBEGIN\s+TRAN', line, re.IGNORECASE):
                    self.violations.append(Violation(
                        rule_id="tsql-error-handling",
                        severity=Severity.HIGH,
                        line_number=i,
                        line_content=line.strip(),
                        message="Transaction without TRY-CATCH error handling",
                        suggestion="Wrap transaction in TRY-CATCH with ROLLBACK in CATCH block"
                    ))
                    break
        
        if has_begin_tran and not has_xact_abort:
            self.violations.append(Violation(
                rule_id="tsql-xact-abort",
                severity=Severity.MEDIUM,
                line_number=1,
                line_content="",
                message="Transaction without SET XACT_ABORT ON",
                suggestion="Add SET XACT_ABORT ON for consistent transaction behavior on errors"
            ))


def print_violations(filepath: str, violations: List[Violation]):
    """Print violations in a formatted way."""
    if not violations:
        print(f"✅ {filepath}: No violations found")
        return
    
    print(f"\n❌ {filepath}: {len(violations)} violation(s) found\n")
    print("-" * 80)
    
    # Sort by severity then line number
    severity_order = {Severity.CRITICAL: 0, Severity.HIGH: 1, Severity.MEDIUM: 2, Severity.LOW: 3}
    sorted_violations = sorted(violations, key=lambda x: (severity_order[x.severity], x.line_number))
    
    for v in sorted_violations:
        severity_color = {
            Severity.CRITICAL: "\033[91m",  # Red
            Severity.HIGH: "\033[93m",      # Yellow
            Severity.MEDIUM: "\033[94m",    # Blue
            Severity.LOW: "\033[90m",       # Gray
        }
        reset = "\033[0m"
        
        print(f"{severity_color[v.severity]}[{v.severity.value}]{reset} {v.rule_id}")
        line_preview = v.line_content[:60] + "..." if len(v.line_content) > 60 else v.line_content
        if v.line_number > 0:
            print(f"  Line {v.line_number}: {line_preview}")
        print(f"  ⚠️  {v.message}")
        print(f"  💡 {v.suggestion}")
        print()


def print_summary(violations: List[Violation]):
    """Print a summary of violations by rule ID."""
    if not violations:
        return
    
    print("\n" + "=" * 80)
    print("SUMMARY BY RULE")
    print("=" * 80)
    
    rule_counts = {}
    for v in violations:
        if v.rule_id not in rule_counts:
            rule_counts[v.rule_id] = {"count": 0, "severity": v.severity}
        rule_counts[v.rule_id]["count"] += 1
    
    # Sort by count descending
    sorted_rules = sorted(rule_counts.items(), key=lambda x: -x[1]["count"])
    
    for rule_id, data in sorted_rules:
        print(f"  {rule_id}: {data['count']} violation(s) [{data['severity'].value}]")


def main():
    if len(sys.argv) < 2:
        print("Usage: python analyze-tsql.py <file_or_directory>")
        print("       python analyze-tsql.py --all")
        sys.exit(1)
    
    analyzer = TSQLAnalyzer()
    files_to_analyze = []
    
    if sys.argv[1] == "--all":
        files_to_analyze = list(Path(".").rglob("*.sql"))
    elif os.path.isdir(sys.argv[1]):
        files_to_analyze = list(Path(sys.argv[1]).rglob("*.sql"))
    else:
        files_to_analyze = [Path(sys.argv[1])]
    
    if not files_to_analyze:
        print("No .sql files found")
        sys.exit(0)
    
    all_violations = []
    total_violations = 0
    critical_count = 0
    
    for filepath in files_to_analyze:
        violations = analyzer.analyze_file(str(filepath))
        print_violations(str(filepath), violations)
        all_violations.extend(violations)
        total_violations += len(violations)
        critical_count += sum(1 for v in violations if v.severity == Severity.CRITICAL)
    
    print_summary(all_violations)
    
    print("\n" + "=" * 80)
    print(f"Total: {len(files_to_analyze)} file(s) analyzed, {total_violations} violation(s) found")
    print(f"Critical: {critical_count}, High: {sum(1 for v in all_violations if v.severity == Severity.HIGH)}")
    print(f"Medium: {sum(1 for v in all_violations if v.severity == Severity.MEDIUM)}, Low: {sum(1 for v in all_violations if v.severity == Severity.LOW)}")
    
    # Exit with error if critical violations found
    if critical_count > 0:
        sys.exit(1)


if __name__ == "__main__":
    main()
