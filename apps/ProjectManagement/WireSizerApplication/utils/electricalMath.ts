import {
  WIRE_DATA,
  GROUND_DATA,
  GEC_DATA,
  CONDUIT_EMT_40_PERCENT,
  K_FACTOR_CU,
  K_FACTOR_AL,
  MAX_WIRE_SIZE_CU,
  MAX_WIRE_SIZE_AL,
  MAX_WIRE_SIZE_FORCED,
  MIN_RECOMMENDED_CONDUCTOR_SIZE,
} from '../constants';
import { AppState, CalculationResult, WireSizeData, ConductorMaterial } from '../types';

export const isThreePhaseAllowed = (voltage: number) => voltage !== 120 && voltage !== 277;

export const getHotCount = (phase: number, voltage: number) => {
  if (phase === 3) return 3;
  if (phase === 1) {
    if (voltage === 120 || voltage === 277) return 1;
    if (voltage === 208 || voltage === 240 || voltage === 480) return 2;
  }
  return phase === 3 ? 3 : 2;
};

const indexOf1_0 = WIRE_DATA.findIndex((wire) => wire.size === '1/0');
const indexOfMinRecommendedConductor = WIRE_DATA.findIndex(
  (wire) => wire.size === MIN_RECOMMENDED_CONDUCTOR_SIZE
);

function getWireTableForMaterial(material: ConductorMaterial, maxSizeOverride?: string): WireSizeData[] {
  const maxSize = maxSizeOverride ?? (material === 'Copper' ? MAX_WIRE_SIZE_CU : MAX_WIRE_SIZE_AL);
  const maxIndex = WIRE_DATA.findIndex((wire) => wire.size === maxSize);
  return maxIndex === -1 ? WIRE_DATA : WIRE_DATA.slice(0, maxIndex + 1);
}

function getRecommendedWireTable(material: ConductorMaterial, maxSizeOverride?: string): WireSizeData[] {
  const wireTable = getWireTableForMaterial(material, maxSizeOverride);
  if (indexOfMinRecommendedConductor === -1) {
    return wireTable;
  }

  const minWire = WIRE_DATA[indexOfMinRecommendedConductor];
  const minWireIndexInTable = wireTable.findIndex((wire) => wire.size === minWire.size);

  if (minWireIndexInTable === -1) {
    return wireTable;
  }

  return wireTable.slice(minWireIndexInTable);
}

function getAmpacity(wire: WireSizeData, material: ConductorMaterial, wireIndex: number): number {
  const isSmallWire = wireIndex < indexOf1_0;
  if (material === 'Copper') {
    return isSmallWire ? wire.ampacity60Cu : wire.ampacity75Cu;
  }
  return isSmallWire ? wire.ampacity60Al : wire.ampacity75Al;
}

function getWireIndex(size: string) {
  return WIRE_DATA.findIndex((wire) => wire.size === size);
}

function getWireBySize(size: string) {
  return WIRE_DATA.find((wire) => wire.size === size);
}

function getMinimumRecommendedWire() {
  if (indexOfMinRecommendedConductor === -1) {
    return undefined;
  }
  return WIRE_DATA[indexOfMinRecommendedConductor];
}

function isBelowMinimumRecommendedWire(size: string) {
  const wireIndex = getWireIndex(size);
  return (
    wireIndex !== -1 &&
    indexOfMinRecommendedConductor !== -1 &&
    wireIndex < indexOfMinRecommendedConductor
  );
}

function getTempRatingForWire(wire: WireSizeData): 60 | 75 {
  const wireIndex = getWireIndex(wire.size);
  if (wireIndex === -1) {
    return 60;
  }
  return wireIndex < indexOf1_0 ? 60 : 75;
}

function getTotalAmpacityForWire(
  wire: WireSizeData,
  material: ConductorMaterial,
  sets: number
): number {
  const wireIndex = getWireIndex(wire.size);
  if (wireIndex === -1) {
    return 0;
  }
  return getAmpacity(wire, material, wireIndex) * sets;
}

type FeederSpecResult = Pick<
  CalculationResult,
  'conduitSize' | 'selectedSize' | 'groundWireSize' | 'voltageDropPercentage' | 'sets'
>;

export const formatFeederSpec = (state: AppState, results: FeederSpecResult) => {
  const conduit = `${results.conduitSize}"C.`;
  const effectivePhase = isThreePhaseAllowed(state.voltage) ? state.phase : 1;
  const numHots = getHotCount(effectivePhase, state.voltage);
  const gndLabel = state.groundingTable === 'GEC' ? 'GEC' : 'G';
  const setsPrefix = results.sets > 1 ? `(${results.sets} sets) ` : '';
  const line1 = `${setsPrefix}${conduit}, ${numHots}#${results.selectedSize} H, 1#${results.selectedSize} N, 1#${results.groundWireSize} ${gndLabel}`;
  const line2 = `LENGTH ~= ${state.distance}'`;
  const line3 = `VOLTAGE DROP ~= ${results.voltageDropPercentage.toFixed(2)}%`;

  return `${line1}\n${line2}\n${line3}`;
};

interface WireSelection {
  wire: WireSizeData;
  sets: number;
  tempRating: 60 | 75;
}

function findOptimalWireAndSets(
  amperage: number,
  distance: number,
  voltage: number,
  maxVoltageDrop: number,
  material: ConductorMaterial,
  phaseFactor: number,
  kFactor: number,
  maxSizeOverride?: string
): WireSelection {
  const wireTable = getRecommendedWireTable(material, maxSizeOverride);
  const maxDropVolts = voltage * (maxVoltageDrop / 100);
  const MAX_SETS = 10;

  for (let sets = 1; sets <= MAX_SETS; sets++) {
    const ampsPerSet = amperage / sets;

    let ampacityWire: WireSizeData | undefined;
    for (let i = 0; i < wireTable.length; i++) {
      const globalIndex = getWireIndex(wireTable[i].size);
      const ampacity = getAmpacity(wireTable[i], material, globalIndex);
      if (ampacity >= ampsPerSet) {
        ampacityWire = wireTable[i];
        break;
      }
    }

    if (!ampacityWire) {
      continue;
    }

    let selectedWire = ampacityWire;

    if (maxDropVolts > 0 && distance > 0) {
      const requiredCM = (kFactor * ampsPerSet * distance * phaseFactor) / maxDropVolts;

      if (selectedWire.circularMils < requiredCM) {
        const upsized = wireTable.find((wire) => wire.circularMils >= requiredCM);
        if (upsized) {
          selectedWire = upsized;
        } else {
          continue;
        }
      }
    }

    return { wire: selectedWire, sets, tempRating: getTempRatingForWire(selectedWire) };
  }

  const fallbackWire = wireTable[wireTable.length - 1];
  return {
    wire: fallbackWire,
    sets: MAX_SETS,
    tempRating: getTempRatingForWire(fallbackWire),
  };
}

function selectWireForSets(
  amperage: number,
  distance: number,
  voltage: number,
  maxVoltageDrop: number,
  material: ConductorMaterial,
  phaseFactor: number,
  kFactor: number,
  sets: number,
  maxSizeOverride?: string
): WireSelection {
  const wireTable = getRecommendedWireTable(material, maxSizeOverride);
  const ampsPerSet = amperage / sets;

  let foundWire: WireSizeData | undefined;
  for (let i = 0; i < wireTable.length; i++) {
    const globalIndex = getWireIndex(wireTable[i].size);
    const ampacity = getAmpacity(wireTable[i], material, globalIndex);
    if (ampacity >= ampsPerSet) {
      foundWire = wireTable[i];
      break;
    }
  }

  if (!foundWire) {
    foundWire = wireTable[wireTable.length - 1];
  }

  const maxDropVolts = voltage * (maxVoltageDrop / 100);
  if (maxDropVolts > 0 && distance > 0) {
    const requiredCM = (kFactor * ampsPerSet * distance * phaseFactor) / maxDropVolts;
    if (foundWire.circularMils < requiredCM) {
      const upsized = wireTable.find((wire) => wire.circularMils >= requiredCM);
      if (upsized) {
        foundWire = upsized;
      }
    }
  }

  return { wire: foundWire, sets, tempRating: getTempRatingForWire(foundWire) };
}

export const calculateEverything = (state: AppState): CalculationResult => {
  const {
    voltage,
    phase,
    material,
    oversizeConduit,
    forceWireSize,
    forcedWireSize,
  } = state;
  const effectivePhase = isThreePhaseAllowed(voltage) ? phase : 1;

  const amperage = state.amperage === '' ? 0 : state.amperage;
  const distance = state.distance === '' ? 0 : state.distance;
  const maxVoltageDrop = state.maxVoltageDrop === '' ? 3 : state.maxVoltageDrop;

  const userSets = state.sets === '' || state.sets === 0 ? 1 : state.sets;
  const forceSets = state.forceSets;

  const kFactor = material === 'Copper' ? K_FACTOR_CU : K_FACTOR_AL;
  const phaseFactor = effectivePhase === 3 ? 1.732 : 2;

  const autoResult = findOptimalWireAndSets(
    amperage,
    distance,
    voltage,
    maxVoltageDrop,
    material,
    phaseFactor,
    kFactor
  );

  const effectiveSets = forceSets ? userSets : Math.max(userSets, autoResult.sets);

  let recommendedWire: WireSizeData;
  let recommendedTempRating: 60 | 75;

  if (forceSets) {
    const forcedResult = selectWireForSets(
      amperage,
      distance,
      voltage,
      maxVoltageDrop,
      material,
      phaseFactor,
      kFactor,
      effectiveSets,
      MAX_WIRE_SIZE_FORCED
    );
    recommendedWire = forcedResult.wire;
    recommendedTempRating = forcedResult.tempRating;
  } else if (effectiveSets > autoResult.sets) {
    const overrideResult = selectWireForSets(
      amperage,
      distance,
      voltage,
      maxVoltageDrop,
      material,
      phaseFactor,
      kFactor,
      effectiveSets
    );
    recommendedWire = overrideResult.wire;
    recommendedTempRating = overrideResult.tempRating;
  } else {
    recommendedWire = autoResult.wire;
    recommendedTempRating = autoResult.tempRating;
  }

  let selectedWire = recommendedWire;
  let finalTempRating = recommendedTempRating;
  const warnings: string[] = [];
  const forcedWire = forceWireSize ? getWireBySize(forcedWireSize) : undefined;
  const minimumRecommendedWire = getMinimumRecommendedWire();
  const isWireSizeForced = Boolean(forceWireSize && forcedWire);

  if (forceWireSize) {
    if (forcedWire) {
      if (minimumRecommendedWire && isBelowMinimumRecommendedWire(forcedWire.size)) {
        selectedWire = minimumRecommendedWire;
        finalTempRating = getTempRatingForWire(minimumRecommendedWire);
        warnings.push(
          `Forced wire size #${forcedWire.size} is below the #${MIN_RECOMMENDED_CONDUCTOR_SIZE} minimum. Using #${minimumRecommendedWire.size} instead.`
        );
      } else {
        selectedWire = forcedWire;
        finalTempRating = getTempRatingForWire(forcedWire);
      }
    } else if (forcedWireSize) {
      warnings.push(
        `Forced wire size #${forcedWireSize} is not available in the conductor table. Showing the recommended size instead.`
      );
    }
  }

  const sets = effectiveSets;
  const ampsPerSet = amperage / sets;

  const actualCM = selectedWire.circularMils;
  const actualVD =
    actualCM > 0 ? (kFactor * ampsPerSet * distance * phaseFactor) / actualCM : 0;
  const actualVDPercent = voltage > 0 ? (actualVD / voltage) * 100 : 0;
  const voltsAtLoad = voltage - actualVD;
  const actualAmpacity = getTotalAmpacityForWire(selectedWire, material, sets);

  if (isWireSizeForced) {
    if (actualAmpacity < amperage) {
      warnings.push(
        `Forced wire size #${selectedWire.size} has only ${actualAmpacity}A total ampacity across ${sets} set${sets === 1 ? '' : 's'}, below the ${amperage}A load.`
      );
    }
    if (maxVoltageDrop > 0 && actualVDPercent > maxVoltageDrop) {
      warnings.push(
        `Forced wire size #${selectedWire.size} causes ${actualVDPercent.toFixed(2)}% voltage drop, above the ${maxVoltageDrop}% limit.`
      );
    }
  }

  let groundWireSize: string;

  if (state.groundingTable === 'GEC') {
    const conductorCM = selectedWire.circularMils * sets;
    const gecRow = GEC_DATA.find((ground) => {
      const threshold = material === 'Copper' ? ground.maxCuCM : ground.maxAlCM;
      return conductorCM <= threshold;
    });
    groundWireSize = gecRow
      ? material === 'Copper'
        ? gecRow.gecCuSize
        : gecRow.gecAlSize
      : 'See Eng.';
  } else {
    const groundRow = GROUND_DATA.find((ground) => ground.rating >= amperage);
    groundWireSize = groundRow
      ? material === 'Copper'
        ? groundRow.cuSize
        : groundRow.alSize
      : 'See Eng.';
  }

  const hotCount = getHotCount(effectivePhase, voltage);
  const numCurrentCarrying = hotCount + 1;
  const totalWiresPerSet = numCurrentCarrying;

  const groundWireObj = getWireBySize(groundWireSize);
  const groundArea = groundWireObj ? groundWireObj.areaSqIn : 0.02;
  const totalFillArea = selectedWire.areaSqIn * totalWiresPerSet + groundArea;

  const minConduitIndex = CONDUIT_EMT_40_PERCENT.findIndex((conduit) => conduit.size === '3/4');
  const availableConduits = CONDUIT_EMT_40_PERCENT.slice(minConduitIndex);

  let recommendedConduitIndex = availableConduits.findIndex(
    (conduit) => conduit.area >= totalFillArea
  );
  if (recommendedConduitIndex === -1) {
    recommendedConduitIndex = availableConduits.length - 1;
  }
  if (oversizeConduit && recommendedConduitIndex < availableConduits.length - 1) {
    recommendedConduitIndex++;
  }

  const recommendedConduit = availableConduits[recommendedConduitIndex];

  return {
    recommendedSize: recommendedWire.size,
    selectedSize: selectedWire.size,
    isWireSizeForced,
    warnings,
    actualAmpacity,
    voltageDrop: actualVD,
    voltageDropPercentage: actualVDPercent,
    voltageAtLoad: voltsAtLoad,
    groundWireSize,
    conduitSize: recommendedConduit ? recommendedConduit.size : '4"+',
    conduitType: 'EMT',
    maxDistanceFor3Percent:
      actualCM > 0
        ? (actualCM * (voltage * 0.03)) / (kFactor * ampsPerSet * phaseFactor)
        : 0,
    wireAreaTotal: totalFillArea,
    conduitFillPercentage: recommendedConduit
      ? (totalFillArea / (recommendedConduit.area / 0.4)) * 100
      : 100,
    tempRatingUsed: finalTempRating,
    sets,
    recommendedSets: autoResult.sets,
  };
};
