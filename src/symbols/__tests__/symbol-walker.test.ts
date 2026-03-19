// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { describe, it, expect } from 'vitest';
import { parseCode } from '../../parser/index.js';
import { collectDeclarations } from '../symbol-walker.js';
import type { Declaration, DeclarationType } from '../declaration.js';

function getDeclarations(code: string, moduleName = 'Module1'): Declaration[] {
  const result = parseCode(code);
  return collectDeclarations(result, moduleName);
}

function findByType(decls: Declaration[], type: DeclarationType): Declaration[] {
  return decls.filter(d => d.declarationType === type);
}

function findByName(decls: Declaration[], name: string): Declaration | undefined {
  return decls.find(d => d.name.toLowerCase() === name.toLowerCase());
}

describe('Symbol Walker — Declaration Collection', () => {
  describe('Module declaration', () => {
    it('creates a Module declaration as root', () => {
      const decls = getDeclarations('Sub Test()\nEnd Sub\n');
      const mod = findByType(decls, 'Module');
      expect(mod).toHaveLength(1);
      expect(mod[0].name).toBe('Module1');
      expect(mod[0].accessibility).toBe('Public');
    });

    it('uses custom module name', () => {
      const decls = getDeclarations('Sub Test()\nEnd Sub\n', 'MyModule');
      const mod = findByType(decls, 'Module');
      expect(mod[0].name).toBe('MyModule');
    });
  });

  describe('Sub declarations', () => {
    it('extracts a Sub', () => {
      const decls = getDeclarations('Sub MySub()\nEnd Sub\n');
      const subs = findByType(decls, 'Sub');
      expect(subs).toHaveLength(1);
      expect(subs[0].name).toBe('MySub');
      expect(subs[0].accessibility).toBe('Public'); // implicit
    });

    it('extracts Private Sub', () => {
      const decls = getDeclarations('Private Sub MySub()\nEnd Sub\n');
      const subs = findByType(decls, 'Sub');
      expect(subs[0].accessibility).toBe('Private');
    });

    it('sets parentScope to module', () => {
      const decls = getDeclarations('Sub MySub()\nEnd Sub\n');
      const sub = findByType(decls, 'Sub')[0];
      expect(sub.parentScope?.declarationType).toBe('Module');
    });
  });

  describe('Function declarations', () => {
    it('extracts a Function with return type', () => {
      const decls = getDeclarations('Function GetValue() As Long\nEnd Function\n');
      const funcs = findByType(decls, 'Function');
      expect(funcs).toHaveLength(1);
      expect(funcs[0].name).toBe('GetValue');
      expect(funcs[0].asTypeName).toBe('Long');
      expect(funcs[0].isImplicitType).toBe(false);
    });

    it('marks implicit Variant return type', () => {
      const decls = getDeclarations('Function GetValue()\nEnd Function\n');
      const func = findByType(decls, 'Function')[0];
      expect(func.isImplicitType).toBe(true);
      expect(func.asTypeName).toBeUndefined();
    });
  });

  describe('Property declarations', () => {
    it('extracts Property Get', () => {
      const decls = getDeclarations('Property Get Value() As Long\nEnd Property\n');
      const props = findByType(decls, 'PropertyGet');
      expect(props).toHaveLength(1);
      expect(props[0].name).toBe('Value');
      expect(props[0].asTypeName).toBe('Long');
    });

    it('extracts Property Let', () => {
      const decls = getDeclarations('Property Let Value(ByVal v As Long)\nEnd Property\n');
      const props = findByType(decls, 'PropertyLet');
      expect(props).toHaveLength(1);
      expect(props[0].name).toBe('Value');
    });

    it('extracts Property Set', () => {
      const decls = getDeclarations('Property Set Obj(ByVal v As Object)\nEnd Property\n');
      const props = findByType(decls, 'PropertySet');
      expect(props).toHaveLength(1);
      expect(props[0].name).toBe('Obj');
    });
  });

  describe('Variable declarations', () => {
    it('extracts Dim variable with type', () => {
      const decls = getDeclarations('Sub Test()\n    Dim x As Long\nEnd Sub\n');
      const vars = findByType(decls, 'Variable');
      expect(vars).toHaveLength(1);
      expect(vars[0].name).toBe('x');
      expect(vars[0].asTypeName).toBe('Long');
      expect(vars[0].isImplicitType).toBe(false);
    });

    it('extracts Dim variable without type (implicit Variant)', () => {
      const decls = getDeclarations('Sub Test()\n    Dim x\nEnd Sub\n');
      const vars = findByType(decls, 'Variable');
      expect(vars[0].isImplicitType).toBe(true);
    });

    it('extracts multiple variables in one Dim', () => {
      const decls = getDeclarations('Sub Test()\n    Dim a As Long, b As String\nEnd Sub\n');
      const vars = findByType(decls, 'Variable');
      expect(vars).toHaveLength(2);
      expect(vars[0].name).toBe('a');
      expect(vars[1].name).toBe('b');
    });

    it('sets parentScope to procedure for locals', () => {
      const decls = getDeclarations('Sub MySub()\n    Dim x As Long\nEnd Sub\n');
      const v = findByType(decls, 'Variable')[0];
      expect(v.parentScope?.name).toBe('MySub');
      expect(v.parentScope?.declarationType).toBe('Sub');
    });

    it('sets parentScope to module for module-level', () => {
      const decls = getDeclarations('Dim x As Long\nSub Test()\nEnd Sub\n');
      const v = findByType(decls, 'Variable')[0];
      expect(v.parentScope?.declarationType).toBe('Module');
    });

    it('extracts Public module-level variable', () => {
      const decls = getDeclarations('Public x As Long\n');
      const v = findByType(decls, 'Variable')[0];
      expect(v.accessibility).toBe('Public');
    });

    it('extracts array variable', () => {
      const decls = getDeclarations('Sub Test()\n    Dim arr(10) As Long\nEnd Sub\n');
      const v = findByType(decls, 'Variable')[0];
      expect(v.isArray).toBe(true);
    });
  });

  describe('Constant declarations', () => {
    it('extracts Const with type', () => {
      const decls = getDeclarations('Sub Test()\n    Const PI As Double = 3.14\nEnd Sub\n');
      const consts = findByType(decls, 'Constant');
      expect(consts).toHaveLength(1);
      expect(consts[0].name).toBe('PI');
      expect(consts[0].asTypeName).toBe('Double');
    });

    it('extracts Public Const at module level', () => {
      const decls = getDeclarations('Public Const MAX_SIZE As Long = 100\n');
      const c = findByType(decls, 'Constant')[0];
      expect(c.accessibility).toBe('Public');
    });
  });

  describe('Enum declarations', () => {
    it('extracts Enum and its members', () => {
      const code = 'Public Enum Color\n    Red\n    Green\n    Blue\nEnd Enum\n';
      const decls = getDeclarations(code);
      const enums = findByType(decls, 'Enum');
      expect(enums).toHaveLength(1);
      expect(enums[0].name).toBe('Color');
      expect(enums[0].accessibility).toBe('Public');

      const members = findByType(decls, 'EnumMember');
      expect(members).toHaveLength(3);
      expect(members.map(m => m.name)).toEqual(['Red', 'Green', 'Blue']);
      expect(members[0].parentScope?.name).toBe('Color');
    });
  });

  describe('Type (UDT) declarations', () => {
    it('extracts Type and its members', () => {
      const code = 'Type Point\n    X As Long\n    Y As Long\nEnd Type\n';
      const decls = getDeclarations(code);
      const types = findByType(decls, 'Type');
      expect(types).toHaveLength(1);
      expect(types[0].name).toBe('Point');

      const members = findByType(decls, 'TypeMember');
      expect(members).toHaveLength(2);
      expect(members[0].parentScope?.name).toBe('Point');
    });
  });

  describe('Event declarations', () => {
    it('extracts Event', () => {
      const code = 'Public Event Click()\n';
      const decls = getDeclarations(code);
      const events = findByType(decls, 'Event');
      expect(events).toHaveLength(1);
      expect(events[0].name).toBe('Click');
    });
  });

  describe('Parameter declarations', () => {
    it('extracts parameters from Sub', () => {
      const code = 'Sub Test(x As Long, y As String)\nEnd Sub\n';
      const decls = getDeclarations(code);
      const params = findByType(decls, 'Parameter');
      expect(params).toHaveLength(2);
      expect(params[0].name).toBe('x');
      expect(params[0].asTypeName).toBe('Long');
      expect(params[1].name).toBe('y');
    });

    it('detects ByVal parameter', () => {
      const code = 'Sub Test(ByVal x As Long)\nEnd Sub\n';
      const decls = getDeclarations(code);
      const param = findByType(decls, 'Parameter')[0];
      expect(param.isByRef).toBe(false);
    });

    it('detects ByRef parameter (explicit)', () => {
      const code = 'Sub Test(ByRef x As Long)\nEnd Sub\n';
      const decls = getDeclarations(code);
      const param = findByType(decls, 'Parameter')[0];
      expect(param.isByRef).toBe(true);
    });

    it('detects implicit ByRef (default)', () => {
      const code = 'Sub Test(x As Long)\nEnd Sub\n';
      const decls = getDeclarations(code);
      const param = findByType(decls, 'Parameter')[0];
      expect(param.isByRef).toBe(true);
    });

    it('detects Optional parameter', () => {
      const code = 'Sub Test(Optional x As Long = 0)\nEnd Sub\n';
      const decls = getDeclarations(code);
      const param = findByType(decls, 'Parameter')[0];
      expect(param.isOptional).toBe(true);
    });

    it('sets parentScope to procedure', () => {
      const code = 'Sub MySub(x As Long)\nEnd Sub\n';
      const decls = getDeclarations(code);
      const param = findByType(decls, 'Parameter')[0];
      expect(param.parentScope?.name).toBe('MySub');
    });
  });

  describe('Line labels', () => {
    it('extracts identifier label', () => {
      const code = 'Sub Test()\nErrorHandler:\n    Resume Next\nEnd Sub\n';
      const decls = getDeclarations(code);
      const labels = findByType(decls, 'LineLabel');
      expect(labels).toHaveLength(1);
      expect(labels[0].name).toBe('ErrorHandler');
    });
  });

  describe('Complex module', () => {
    it('extracts all declaration types from a complex module', () => {
      const code = `
Option Explicit

Public Const MAX_SIZE As Long = 100
Private mCount As Long

Public Enum Status
    Active
    Inactive
End Enum

Type Point
    X As Long
    Y As Long
End Type

Public Sub Initialize(count As Long)
    Dim i As Long
    mCount = count
End Sub

Public Function GetCount() As Long
    GetCount = mCount
End Function

Property Get Value() As Long
    Value = mCount
End Property
`;
      const decls = getDeclarations(code);

      expect(findByType(decls, 'Module')).toHaveLength(1);
      expect(findByType(decls, 'Constant')).toHaveLength(1);
      expect(findByType(decls, 'Variable').length).toBeGreaterThanOrEqual(2); // mCount + i
      expect(findByType(decls, 'Enum')).toHaveLength(1);
      expect(findByType(decls, 'EnumMember')).toHaveLength(2);
      expect(findByType(decls, 'Type')).toHaveLength(1);
      expect(findByType(decls, 'TypeMember')).toHaveLength(2);
      expect(findByType(decls, 'Sub')).toHaveLength(1);
      expect(findByType(decls, 'Function')).toHaveLength(1);
      expect(findByType(decls, 'PropertyGet')).toHaveLength(1);
      expect(findByType(decls, 'Parameter')).toHaveLength(1);
    });
  });
});
