import { describe, expect, it } from 'vitest'

import { parsePageRanges } from '../src/renderer/page-ranges'

describe('parsePageRanges', () => {
  it('normalizes single pages and spans, dropping whitespace', () => {
    expect(parsePageRanges('1-3,5')).toBe('1-3,5')
    expect(parsePageRanges(' 1 - 3 , 5 ')).toBe('1-3,5')
    expect(parsePageRanges('2,7,11')).toBe('2,7,11')
    expect(parsePageRanges('4-4')).toBe('4')
  })

  it('rejects malformed input', () => {
    expect(parsePageRanges('')).toBeNull()
    expect(parsePageRanges('   ')).toBeNull()
    expect(parsePageRanges('abc')).toBeNull()
    expect(parsePageRanges('1-')).toBeNull()
    expect(parsePageRanges('-3')).toBeNull()
    expect(parsePageRanges('1,,2')).toBeNull()
    expect(parsePageRanges('1;2')).toBeNull()
    expect(parsePageRanges('1-3,')).toBeNull()
  })

  it('rejects out-of-range segments', () => {
    expect(parsePageRanges('0')).toBeNull()
    expect(parsePageRanges('5-2')).toBeNull()
    expect(parsePageRanges('0-3')).toBeNull()
  })
})
