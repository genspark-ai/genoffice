export const READ_UNITS = ['emu', 'pt', 'px', 'cm', 'in'] as const
export type ReadUnit = (typeof READ_UNITS)[number]

const EMU_PER: Record<Exclude<ReadUnit, 'emu'>, number> = {
  pt: 12700,
  px: 9525,
  cm: 360000,
  in: 914400,
}

export function isReadUnit(value: string): value is ReadUnit {
  return (READ_UNITS as readonly string[]).includes(value)
}

export function fromEmu(emu: number, unit: ReadUnit): number {
  if (unit === 'emu') return emu
  const scale = unit === 'pt' || unit === 'px' ? 10 : 100
  return Math.round((emu / EMU_PER[unit]) * scale) / scale
}

export function emuPer(unit: Exclude<ReadUnit, 'emu'>): number {
  return EMU_PER[unit]
}
