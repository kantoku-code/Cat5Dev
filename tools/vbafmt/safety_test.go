package main

import (
	"strings"
	"testing"
)

func TestContinuationTokenSafety(t *testing.T) {
	for _, tc := range []struct {
		input, output string
		continuation  bool
	}{
		{"Debug.Print value_", "Debug.Print value_", false},
		{"Show(value_", "Show(value_", false},
		{"Debug.Print 日本語_", "Debug.Print 日本語_", false},
		{"Debug.Print value1_", "Debug.Print value1_", false},
		{"Show(_", "Show( _", true},
		{"x = 1_", "x = 1 _", true},
		{"x = 1 _", "x = 1 _", true},
		{`x = "abc_"`, `x = "abc_"`, false},
		{"' comment_", "' comment_", false},
	} {
		if got := ensureContinuationSpaceLine(tc.input); got != tc.output {
			t.Errorf("%q: got %q, want %q", tc.input, got, tc.output)
		}
		if got := isContinuationLine(tc.input); got != tc.continuation {
			t.Errorf("%q: continuation=%v", tc.input, got)
		}
	}
}

func TestNamedArgumentTokenSafety(t *testing.T) {
	for _, input := range []string{`MsgBox Prompt:="hello"`, `MsgBox Prompt := "hello"`, `f x:=1,y:=2`} {
		got := normalizeOperatorLine(input)
		if strings.Contains(got, ": =") || !strings.Contains(got, ":=") {
			t.Errorf("named argument corrupted: %q", got)
		}
	}
	parts := splitColonLine(`f x:=1: f y:=2`)
	if len(parts) != 2 || parts[0] != "f x:=1" || parts[1] != "f y:=2" {
		t.Fatalf("colon split corrupted named arguments: %q", parts)
	}
}

func TestSafetyFormatIdempotent(t *testing.T) {
	opts := DefaultOptions()
	opts.LineEndings = "LF"
	opts.NormalizeOperatorSpacing = true
	input := "Sub Test()\nDim value_ As Long\nDebug.Print value_\nDebug.Print 日本語_\nMsgBox Prompt:=\"hello\"\nEnd Sub\n"
	got := Format(input, opts)
	if strings.Contains(got, "value _") || strings.Contains(got, "日本語 _") || strings.Contains(got, ": =") || !strings.Contains(got, "\nEnd Sub\n") {
		t.Fatalf("unsafe format: %s", got)
	}
	if twice := Format(got, opts); twice != got {
		t.Fatalf("not idempotent:\n%s", twice)
	}
	if got := normalizeThenPlacement([]string{"If value_", "Then"}); len(got) != 2 {
		t.Fatalf("identifier joined with Then: %q", got)
	}
	// Opposite unmatched parentheses must not cancel across separate statements.
	diags := checkUnmatchedParens([]string{"Debug.Print (value_", "Debug.Print )"}, LintOptions{UnmatchedParens: true})
	if len(diags) != 2 {
		t.Fatalf("identifier treated as continuation in lint: %v", diags)
	}
}
