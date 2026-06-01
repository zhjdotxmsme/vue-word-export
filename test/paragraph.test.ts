import { describe, it, expect } from 'vitest'
import { buildParagraph } from '../src/engine/paragraph'
import { defaultTheme } from '../src/engine/theme'

describe('buildParagraph', () => {
  const theme = defaultTheme()

  it('should handle simple text content', () => {
    const result = buildParagraph({
      type: 'paragraph',
      content: { type: 'text', text: 'Hello world' },
    }, theme)

    expect(result.length).toBe(1)
  })

  it('should handle mixed rich text fragments', () => {
    const result = buildParagraph({
      type: 'paragraph',
      content: [
        { type: 'text', text: 'Bold ', style: { bold: true } },
        { type: 'text', text: 'and normal' },
      ],
    }, theme)

    expect(result.length).toBe(1)
  })

  it('should handle break fragments', () => {
    const result = buildParagraph({
      type: 'paragraph',
      content: [
        { type: 'text', text: 'First line' },
        { type: 'break' },
        { type: 'text', text: 'After break' },
      ],
    }, theme)

    expect(result.length).toBeGreaterThanOrEqual(2)
  })

  it('should render section title when provided', () => {
    const result = buildParagraph({
      type: 'paragraph',
      title: 'Section Title',
      content: { type: 'text', text: 'Body content' },
    }, theme)

    expect(result.length).toBeGreaterThanOrEqual(2)
  })

  it('should handle dynamic text via function', () => {
    const result = buildParagraph({
      type: 'paragraph',
      content: { type: 'text', text: () => 'Dynamic Text' },
    }, theme)

    expect(result.length).toBe(1)
  })

  it('should apply alignment style', () => {
    const result = buildParagraph({
      type: 'paragraph',
      content: { type: 'text', text: 'Centered' },
      style: { align: 'center' },
    }, theme)

    expect(result.length).toBe(1)
  })

  it('should apply indent style', () => {
    const result = buildParagraph({
      type: 'paragraph',
      content: { type: 'text', text: 'Indented text' },
      style: { indentFirstLine: 480 },
    }, theme)

    expect(result.length).toBe(1)
  })

  it('should handle empty content array', () => {
    const result = buildParagraph({
      type: 'paragraph',
      content: [],
    }, theme)

    expect(result.length).toBe(0)
  })
})
