import { describe, expect, it } from 'vitest';
import { WIRE_DATA } from '../constants';
import { AppState } from '../types';
import { calculateEverything, formatFeederSpec } from './electricalMath';

const createState = (overrides: Partial<AppState> = {}): AppState => ({
  voltage: 208,
  amperage: 20,
  distance: 100,
  phase: 3,
  material: 'Copper',
  maxVoltageDrop: 3,
  sets: 1,
  forceSets: false,
  forceWireSize: false,
  forcedWireSize: '12',
  powerFactor: 0.9,
  oversizeConduit: false,
  groundingTable: 'EGC',
  ...overrides,
});

const getWireIndex = (size: string) => WIRE_DATA.findIndex((wire) => wire.size === size);

const getOffsetWireSize = (size: string, offset: number) => {
  const currentIndex = getWireIndex(size);
  const nextIndex = Math.min(WIRE_DATA.length - 1, Math.max(0, currentIndex + offset));
  return WIRE_DATA[nextIndex].size;
};

const getDifferentWireSize = (size: string) => {
  const currentIndex = getWireIndex(size);
  if (currentIndex < WIRE_DATA.length - 1) {
    return WIRE_DATA[currentIndex + 1].size;
  }
  return WIRE_DATA[Math.max(0, currentIndex - 1)].size;
};

describe('calculateEverything', () => {
  it('keeps auto sizing unchanged when forced wire size is off', () => {
    const baselineState = createState({ amperage: 70, distance: 180, voltage: 240, phase: 1 });
    const baseline = calculateEverything(baselineState);
    const ignoredOverride = calculateEverything({
      ...baselineState,
      forcedWireSize: getDifferentWireSize(baseline.selectedSize),
    });

    expect(ignoredOverride.selectedSize).toBe(baseline.selectedSize);
    expect(ignoredOverride.recommendedSize).toBe(baseline.recommendedSize);
    expect(ignoredOverride.isWireSizeForced).toBe(false);
    expect(ignoredOverride.warnings).toEqual([]);
  });

  it('reduces voltage drop and increases conduit fill when a larger wire is forced', () => {
    const state = createState({ amperage: 50, distance: 200, voltage: 240, phase: 1 });
    const baseline = calculateEverything(state);
    const largerSize = getOffsetWireSize(baseline.selectedSize, 2);
    const forced = calculateEverything({
      ...state,
      forceWireSize: true,
      forcedWireSize: largerSize,
    });

    expect(forced.selectedSize).toBe(largerSize);
    expect(forced.recommendedSize).toBe(baseline.selectedSize);
    expect(forced.voltageDropPercentage).toBeLessThan(baseline.voltageDropPercentage);
    expect(forced.wireAreaTotal).toBeGreaterThan(baseline.wireAreaTotal);
    expect(forced.warnings).toEqual([]);
  });

  it('warns when a forced smaller wire fails ampacity and voltage-drop limits', () => {
    const state = createState({ amperage: 225, distance: 700, voltage: 480, phase: 3 });
    const forced = calculateEverything({
      ...state,
      forceWireSize: true,
      forcedWireSize: '12',
    });

    expect(forced.selectedSize).toBe('12');
    expect(forced.warnings.some((warning) => /ampacity/i.test(warning))).toBe(true);
    expect(forced.warnings.some((warning) => /voltage drop/i.test(warning))).toBe(true);
  });

  it('keeps forced parallel sets while exposing the recommended size for that set count', () => {
    const state = createState({
      amperage: 900,
      distance: 250,
      voltage: 480,
      phase: 3,
      sets: 3,
      forceSets: true,
    });
    const baseline = calculateEverything(state);
    const forcedSize = getDifferentWireSize(baseline.selectedSize);
    const forced = calculateEverything({
      ...state,
      forceWireSize: true,
      forcedWireSize: forcedSize,
    });

    expect(baseline.sets).toBe(3);
    expect(forced.sets).toBe(3);
    expect(forced.recommendedSize).toBe(baseline.selectedSize);
    expect(forced.selectedSize).toBe(forcedSize);
  });
});

describe('formatFeederSpec', () => {
  it('uses the selected wire size instead of the recommended size', () => {
    const state = createState({ amperage: 90, distance: 220, voltage: 208, phase: 3 });
    const baseline = calculateEverything(state);
    const forced = calculateEverything({
      ...state,
      forceWireSize: true,
      forcedWireSize: getDifferentWireSize(baseline.selectedSize),
    });

    expect(forced.selectedSize).not.toBe(forced.recommendedSize);

    const spec = formatFeederSpec(
      {
        ...state,
        forceWireSize: true,
        forcedWireSize: forced.selectedSize,
      },
      forced
    );
    const firstLine = spec.split('\n')[0];

    expect(firstLine).toContain(`#${forced.selectedSize} H, 1#${forced.selectedSize} N`);
    expect(firstLine).not.toContain(
      `#${forced.recommendedSize} H, 1#${forced.recommendedSize} N`
    );
  });
});
