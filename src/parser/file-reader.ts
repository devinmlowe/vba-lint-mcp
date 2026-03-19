// Derived from Rubberduck VBA — Copyright (C) Rubberduck Contributors
// License: GPL-3.0 — https://github.com/rubberduck-vba/Rubberduck/blob/main/LICENSE

import { readFile } from 'node:fs/promises';

const MAX_FILE_SIZE = 512 * 1024; // 512KB default limit

/**
 * Read a VBA file with proper encoding detection, BOM stripping,
 * and CRLF normalization.
 *
 * VBA files exported from the VBE can be:
 * - UTF-8 (modern editors)
 * - UTF-8 with BOM (common on Windows)
 * - Windows-1252 / ANSI (legacy VBE default)
 */
export async function readVBAFile(
  filePath: string,
  maxSize: number = MAX_FILE_SIZE,
): Promise<string> {
  const buffer = await readFile(filePath);

  if (buffer.length > maxSize) {
    throw new Error(`File exceeds maximum size of ${maxSize} bytes: ${filePath}`);
  }

  let content: string;

  // Check for BOM
  if (buffer[0] === 0xEF && buffer[1] === 0xBB && buffer[2] === 0xBF) {
    // UTF-8 BOM — strip it
    content = buffer.subarray(3).toString('utf-8');
  } else if (buffer[0] === 0xFF && buffer[1] === 0xFE) {
    // UTF-16 LE BOM
    content = buffer.subarray(2).toString('utf16le');
  } else if (isLikelyWindows1252(buffer)) {
    // Contains bytes in 0x80-0x9F range that are valid in Windows-1252
    // but invalid in UTF-8. Decode as Windows-1252.
    content = decodeWindows1252(buffer);
  } else {
    // Default: UTF-8
    content = buffer.toString('utf-8');
  }

  // Normalize CRLF to LF for consistent line counting
  return content.replace(/\r\n/g, '\n');
}

/**
 * Check if a buffer likely contains Windows-1252 encoded text.
 * Windows-1252 uses bytes 0x80-0x9F for printable characters
 * (smart quotes, em-dash, etc.) that are invalid in UTF-8.
 */
function isLikelyWindows1252(buffer: Buffer): boolean {
  for (let i = 0; i < buffer.length; i++) {
    const byte = buffer[i];
    if (byte >= 0x80 && byte <= 0x9F) {
      return true;
    }
  }
  return false;
}

/**
 * Decode a Windows-1252 encoded buffer to string.
 * Maps the 0x80-0x9F range to Unicode codepoints.
 */
function decodeWindows1252(buffer: Buffer): string {
  // Windows-1252 to Unicode mapping for 0x80-0x9F
  const win1252Map: Record<number, number> = {
    0x80: 0x20AC, // €
    0x82: 0x201A, // ‚
    0x83: 0x0192, // ƒ
    0x84: 0x201E, // „
    0x85: 0x2026, // …
    0x86: 0x2020, // †
    0x87: 0x2021, // ‡
    0x88: 0x02C6, // ˆ
    0x89: 0x2030, // ‰
    0x8A: 0x0160, // Š
    0x8B: 0x2039, // ‹
    0x8C: 0x0152, // Œ
    0x8E: 0x017D, // Ž
    0x91: 0x2018, // '
    0x92: 0x2019, // '
    0x93: 0x201C, // "
    0x94: 0x201D, // "
    0x95: 0x2022, // •
    0x96: 0x2013, // –
    0x97: 0x2014, // —
    0x98: 0x02DC, // ˜
    0x99: 0x2122, // ™
    0x9A: 0x0161, // š
    0x9B: 0x203A, // ›
    0x9C: 0x0153, // œ
    0x9E: 0x017E, // ž
    0x9F: 0x0178, // Ÿ
  };

  const chars: string[] = [];
  for (let i = 0; i < buffer.length; i++) {
    const byte = buffer[i];
    if (byte >= 0x80 && byte <= 0x9F && win1252Map[byte] !== undefined) {
      chars.push(String.fromCodePoint(win1252Map[byte]));
    } else {
      chars.push(String.fromCodePoint(byte));
    }
  }
  return chars.join('');
}
