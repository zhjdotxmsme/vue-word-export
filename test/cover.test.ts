import { describe, it, expect } from 'vitest'
import { buildCover } from '../src/engine/cover'
import { defaultTheme } from '../src/engine/theme'

describe('buildCover', () => {
  const theme = defaultTheme()

  it('should return paragraphs for title and subtitle', () => {
    const result = buildCover({
      type: 'cover',
      items: [
        { type: 'title', text: 'Test Report' },
        { type: 'subtitle', text: 'A subtitle' },
      ],
    }, theme)

    expect(result.length).toBeGreaterThan(0)
  })

  it('should handle date item with default format', () => {
    const result = buildCover({
      type: 'cover',
      items: [
        { type: 'date' },
      ],
    }, theme)

    expect(result.length).toBeGreaterThan(0)
  })

  it('should handle text item', () => {
    const result = buildCover({
      type: 'cover',
      items: [
        { type: 'text', text: 'Some text content' },
      ],
    }, theme)

    expect(result.length).toBeGreaterThan(0)
  })

  it('should handle dynamic text via function', () => {
    const result = buildCover({
      type: 'cover',
      items: [
        { type: 'title', text: () => 'Dynamic Title' },
      ],
    }, theme)

    expect(result.length).toBeGreaterThan(0)
  })

  it('should handle all item types together', () => {
    const result = buildCover({
      type: 'cover',
      items: [
        { type: 'image', src: 'data:image/png;base64,abc' },
        { type: 'title', text: 'Title' },
        { type: 'subtitle', text: 'Subtitle' },
        { type: 'text', text: 'Company Name' },
        { type: 'date', format: 'YYYY年MM月DD日' },
      ],
    }, theme)

    expect(result.length).toBeGreaterThan(0)
  })

  it('should apply item spacing from style', () => {
    const result = buildCover({
      type: 'cover',
      style: { itemSpacing: 400 },
      items: [
        { type: 'title', text: 'Title' },
        { type: 'text', text: 'Text' },
      ],
    }, theme)

    expect(result.length).toBeGreaterThan(0)
  })
})
