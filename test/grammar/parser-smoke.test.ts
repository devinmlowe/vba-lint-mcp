// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { CharStream, CommonTokenStream } from 'antlr4ng';
import { VBALexer } from '../../src/parser/generated/grammar/VBALexer.js';
import { VBAParser } from '../../src/parser/generated/grammar/VBAParser.js';
import { patchParserHelpers } from '../../src/parser/vba-parser-helpers.js';

function parseVBA(code: string) {
  const input = CharStream.fromString(code);
  const lexer = new VBALexer(input);
  const tokens = new CommonTokenStream(lexer);
  const parser = new VBAParser(tokens);
  patchParserHelpers(parser);
  parser.removeErrorListeners(); // Suppress console output
  return parser.startRule();
}

describe('VBA Parser Smoke Tests', () => {
  it('parses an empty Sub', () => {
    const tree = parseVBA('Sub Test()\nEnd Sub\n');
    expect(tree).toBeDefined();
    expect(tree.children).toBeDefined();
  });

  it('parses a Sub with a variable declaration', () => {
    const tree = parseVBA('Sub Test()\n    Dim x As Long\n    x = 42\nEnd Sub\n');
    expect(tree).toBeDefined();
  });

  it('parses a Function with return type', () => {
    const tree = parseVBA('Function Add(a As Long, b As Long) As Long\n    Add = a + b\nEnd Function\n');
    expect(tree).toBeDefined();
  });

  it('parses a Class module with Property', () => {
    const code = [
      'Option Explicit',
      '',
      'Private mName As String',
      '',
      'Public Property Get Name() As String',
      '    Name = mName',
      'End Property',
      '',
      'Public Property Let Name(ByVal value As String)',
      '    mName = value',
      'End Property',
      '',
    ].join('\n');
    const tree = parseVBA(code);
    expect(tree).toBeDefined();
  });

  it('parses If/Else blocks', () => {
    const code = [
      'Sub Test()',
      '    If x > 0 Then',
      '        MsgBox "Positive"',
      '    ElseIf x = 0 Then',
      '        MsgBox "Zero"',
      '    Else',
      '        MsgBox "Negative"',
      '    End If',
      'End Sub',
      '',
    ].join('\n');
    const tree = parseVBA(code);
    expect(tree).toBeDefined();
  });

  it('parses For/Next loop', () => {
    const code = [
      'Sub Test()',
      '    Dim i As Long',
      '    For i = 1 To 10',
      '        Debug.Print i',
      '    Next i',
      'End Sub',
      '',
    ].join('\n');
    const tree = parseVBA(code);
    expect(tree).toBeDefined();
  });

  it('parses Select Case', () => {
    const code = [
      'Sub Test()',
      '    Select Case x',
      '        Case 1',
      '            MsgBox "One"',
      '        Case 2, 3',
      '            MsgBox "Two or Three"',
      '        Case Else',
      '            MsgBox "Other"',
      '    End Select',
      'End Sub',
      '',
    ].join('\n');
    const tree = parseVBA(code);
    expect(tree).toBeDefined();
  });

  it('parses error handling', () => {
    const code = [
      'Sub Test()',
      '    On Error GoTo ErrHandler',
      '    Dim x As Long',
      '    x = 1 / 0',
      '    Exit Sub',
      'ErrHandler:',
      '    MsgBox Err.Description',
      'End Sub',
      '',
    ].join('\n');
    const tree = parseVBA(code);
    expect(tree).toBeDefined();
  });

  it('parses Enum declaration', () => {
    const code = [
      'Public Enum Color',
      '    Red = 1',
      '    Green = 2',
      '    Blue = 3',
      'End Enum',
      '',
    ].join('\n');
    const tree = parseVBA(code);
    expect(tree).toBeDefined();
  });

  it('parses Type (UDT) declaration', () => {
    const code = [
      'Private Type Point',
      '    X As Long',
      '    Y As Long',
      'End Type',
      '',
    ].join('\n');
    const tree = parseVBA(code);
    expect(tree).toBeDefined();
  });
});
