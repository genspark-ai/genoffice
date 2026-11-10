import { describe, it, expect } from 'vitest'
import { isCopyrightHost } from '../src/shared'

describe('isCopyrightHost suffix match', () => {
  it('blocks exact stock domains', () => {
    expect(isCopyrightHost('https://shutterstock.com/x.jpg')).toBe(true)
    expect(isCopyrightHost('https://gettyimages.com/x.jpg')).toBe(true)
    expect(isCopyrightHost('https://istockphoto.com/x.jpg')).toBe(true)
    expect(isCopyrightHost('https://corbis.com/x.jpg')).toBe(true)
  })

  it('blocks subdomains of stock hosts', () => {
    expect(isCopyrightHost('https://sub.shutterstock.com/y.jpg')).toBe(true)
    expect(isCopyrightHost('https://media.shutterstock.com/y.jpg')).toBe(true)
    expect(isCopyrightHost('https://media.gettyimages.com/x.jpg')).toBe(true)
    expect(isCopyrightHost('https://www.shutterstock.co.uk/y.jpg')).toBe(true)
    expect(isCopyrightHost('https://gettyimages.com.au/x.jpg')).toBe(true)
  })

  it('allows lookalike domains that merely contain a stock name', () => {
    expect(isCopyrightHost('https://myshutterstock.com/z.jpg')).toBe(false)
    expect(isCopyrightHost('https://notgettyimages.com/a.jpg')).toBe(false)
    expect(isCopyrightHost('https://shutterstock.example.com/a.jpg')).toBe(false)
  })

  it('matches case-insensitively', () => {
    expect(isCopyrightHost('https://SUB.SHUTTERSTOCK.COM/a.jpg')).toBe(true)
    expect(isCopyrightHost('https://MyShutterStock.com/b.jpg')).toBe(false)
    expect(isCopyrightHost('https://MEDIA.GETTYIMAGES.COM/x.jpg')).toBe(true)
  })

  it('ignores stock names that appear only in the path', () => {
    expect(isCopyrightHost('https://cdn.example.com/shutterstock-review.png')).toBe(false)
  })

  it('returns false for invalid urls', () => {
    expect(isCopyrightHost('not a url')).toBe(false)
    expect(isCopyrightHost('')).toBe(false)
  })
})
