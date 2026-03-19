# vba-lint-mcp

An MCP (Model Context Protocol) server providing VBA code inspections, linting, and parse tree analysis.

## Overview

`vba-lint-mcp` exposes VBA static analysis capabilities through the Model Context Protocol, enabling AI assistants and development tools to perform code quality checks, syntax validation, and parse tree inspection on VBA source files.

## Attribution

This project is inspired by and derives significant design and grammar work from the [Rubberduck VBA](https://github.com/rubberduck-vba/Rubberduck) project — an open-source VBIDE add-in providing code inspections, refactoring, navigation, and unit testing for VBA developers.

The VBA grammar definitions used in this project are based on Rubberduck's ANTLR4 grammar work. Rubberduck is licensed under the GNU General Public License v3.0.

## Features

- VBA syntax parsing via ANTLR4 grammar
- Code inspections and lint rules
- Parse tree analysis and traversal
- MCP-compatible tool interface

## License

This project is licensed under the GNU General Public License v3.0. See [LICENSE](LICENSE) for details.

In accordance with GPL-3.0 requirements, derived grammar and inspection logic retain attribution to the Rubberduck VBA project and its contributors.

## Status

Early development. Grammar integration and MCP server scaffolding in progress.
