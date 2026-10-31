import { describe, expect, it, vi } from 'vitest'

import { parseLpOptions, queryPrinterOptions } from '../src/printer-options'
import { lpArgumentsFor } from '../src/print-cups'
import { printJobOptions } from '../src/print-job'

// the shape a real driver reports (an HP laser queue's -l output, trimmed)
const HP_QUEUE = `ColorModel/Color Mode: Gray *RGB
cupsPrintQuality/Quality: *Normal
PageSize/Media Size: 4x6 8.5x13.5 *A4 A5 B5 Letter Custom.WIDTHxHEIGHT
MediaType/MediaType: stationery stationery-heavyweight cardstock labels *any
HPCustomMediaTray/Custom Tray: *None Tray1 Tray2
`

describe('parseLpOptions', () => {
  it('reads name, translation, values and the starred default', () => {
    const groups = parseLpOptions(HP_QUEUE)
    expect(groups.map((g) => g.name)).toEqual([
      'ColorModel',
      'cupsPrintQuality',
      'PageSize',
      'MediaType',
      'HPCustomMediaTray',
    ])
    const color = groups[0]!
    expect(color.translation).toBe('Color Mode')
    expect(color.values).toEqual(['Gray', 'RGB'])
    expect(color.defaultValue).toBe('RGB')
    const quality = groups[1]!
    // a lone default is both the value list and the default
    expect(quality.values).toEqual(['Normal'])
    expect(quality.defaultValue).toBe('Normal')
  })

  it('keeps only charset-clean names and values', () => {
    const groups = parseLpOptions(
      ['Good/Good Name: alpha *beta gamma<x>', 'bad name/Injected: ok', 'Empty/: nogo'].join('\n'),
    )
    // the space-bearing name line is dropped whole; a bad value token inside a
    // good line is skipped without losing its siblings
    expect(groups.map((g) => g.name)).toEqual(['Good'])
    expect(groups[0]!.values).toEqual(['alpha', 'beta'])
    expect(groups[0]!.defaultValue).toBe('beta')
  })

  it('skips lines without a value list', () => {
    expect(parseLpOptions('OnlyAName/NoValues:\n')).toEqual([])
  })
})

describe('queryPrinterOptions', () => {
  it('resolves the parsed groups and swallows failures to empty', async () => {
    const execOk = ((_p: string, _a: string[], cb: (e: Error | null, o: string) => void) => {
      cb(null, HP_QUEUE)
    }) as never
    await expect(queryPrinterOptions('HP', execOk)).resolves.toHaveLength(5)
    const execFail = ((_p: string, _a: string[], cb: (e: Error, o: string) => void) => {
      cb(new Error('No such file or directory'), '')
    }) as never
    await expect(queryPrinterOptions('mangled', execFail)).resolves.toEqual([])
  })

  it('asks lpoptions for exactly the queue', async () => {
    const exec = vi.fn((_p: string, _a: string[], cb: (e: Error | null, o: string) => void) => {
      cb(null, '')
    }) as never
    await queryPrinterOptions('HP_Laser', exec)
    expect(exec).toHaveBeenCalledWith(
      '/usr/bin/lpoptions',
      ['-p', 'HP_Laser', '-l'],
      expect.any(Function),
    )
  })
})

describe('driver options through the job', () => {
  it('printJobOptions carries charset-clean picks and drops the rest', () => {
    const job = printJobOptions({
      deviceName: 'HP',
      driverOptions: [
        { name: 'cupsPrintQuality', value: 'High' },
        { name: 'bad name', value: 'x' },
        { name: 'ColorModel', value: 'Gray<x>' },
        'not-an-object',
      ],
    })
    expect(job.driverOptions).toEqual([{ name: 'cupsPrintQuality', value: 'High' }])
  })

  it('lpArgumentsFor appends the picks as -o name=value', () => {
    expect(
      lpArgumentsFor({
        deviceName: 'HP',
        driverOptions: [
          { name: 'cupsPrintQuality', value: 'High' },
          { name: 'ColorModel', value: 'Gray' },
        ],
      }),
    ).toEqual(['-d', 'HP', '-o', 'cupsPrintQuality=High', '-o', 'ColorModel=Gray'])
  })
})
