// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseIgnoreFile, filterIgnoredFiles } from '../vbalintignore.js';

describe('.vbalintignore', () => {
  describe('parseIgnoreFile', () => {
    it('parses patterns from file content', () => {
      const content = '*.frm\ntests/\n**/generated/**\n';
      expect(parseIgnoreFile(content)).toEqual(['*.frm', 'tests/', '**/generated/**']);
    });

    it('strips comments', () => {
      const content = '# This is a comment\n*.frm\n# Another comment\ntests/\n';
      expect(parseIgnoreFile(content)).toEqual(['*.frm', 'tests/']);
    });

    it('strips blank lines', () => {
      const content = '\n*.frm\n\n\ntests/\n\n';
      expect(parseIgnoreFile(content)).toEqual(['*.frm', 'tests/']);
    });

    it('trims whitespace', () => {
      const content = '  *.frm  \n  tests/  \n';
      expect(parseIgnoreFile(content)).toEqual(['*.frm', 'tests/']);
    });

    it('returns empty array for empty content', () => {
      expect(parseIgnoreFile('')).toEqual([]);
    });

    it('returns empty array for comments-only content', () => {
      expect(parseIgnoreFile('# comment\n# another\n')).toEqual([]);
    });
  });

  describe('filterIgnoredFiles', () => {
    const allFiles = [
      'Module1.bas',
      'Module2.bas',
      'Form1.frm',
      'Class1.cls',
      'tests/TestModule.bas',
      'tests/TestHelper.bas',
      'src/generated/Auto.bas',
      'src/generated/deep/nested.bas',
      'lib/Utils.bas',
    ];

    it('returns all files when no patterns', () => {
      expect(filterIgnoredFiles(allFiles, [])).toEqual(allFiles);
    });

    it('filters *.frm pattern', () => {
      const result = filterIgnoredFiles(allFiles, ['*.frm']);
      expect(result).not.toContain('Form1.frm');
      expect(result).toContain('Module1.bas');
      expect(result).toContain('Class1.cls');
    });

    it('filters tests/ directory pattern', () => {
      const result = filterIgnoredFiles(allFiles, ['tests/**']);
      expect(result).not.toContain('tests/TestModule.bas');
      expect(result).not.toContain('tests/TestHelper.bas');
      expect(result).toContain('Module1.bas');
    });

    it('filters **/generated/** pattern', () => {
      const result = filterIgnoredFiles(allFiles, ['**/generated/**']);
      expect(result).not.toContain('src/generated/Auto.bas');
      expect(result).not.toContain('src/generated/deep/nested.bas');
      expect(result).toContain('Module1.bas');
    });

    it('applies multiple patterns', () => {
      const result = filterIgnoredFiles(allFiles, ['*.frm', 'tests/**', '**/generated/**']);
      expect(result).toEqual([
        'Module1.bas',
        'Module2.bas',
        'Class1.cls',
        'lib/Utils.bas',
      ]);
    });

    it('handles pattern that matches nothing', () => {
      const result = filterIgnoredFiles(allFiles, ['*.xyz']);
      expect(result).toEqual(allFiles);
    });

    it('handles specific file pattern', () => {
      const result = filterIgnoredFiles(allFiles, ['Module1.bas']);
      expect(result).not.toContain('Module1.bas');
      expect(result).toContain('Module2.bas');
    });
  });
});
