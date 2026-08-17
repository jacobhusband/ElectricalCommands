export type ConductorMaterial = 'Copper' | 'Aluminum';
export type Phase = 1 | 3;
export type GroundingTableType = 'EGC' | 'GEC';

export interface WireSizeData {
  size: string;
  circularMils: number;
  ampacity60Cu: number;
  ampacity75Cu: number;
  ampacity60Al: number;
  ampacity75Al: number;
  areaSqIn: number; // For conduit fill
}

export interface GroundingData {
  rating: number; // Rating of overcurrent device
  cuSize: string;
  alSize: string;
}

export interface GECData {
  maxCuCM: number;   // Max circular mils for copper service conductor
  maxAlCM: number;   // Max circular mils for aluminum service conductor
  gecCuSize: string;  // GEC size in copper
  gecAlSize: string;  // GEC size in aluminum
}

export interface CalculationResult {
  recommendedSize: string; // Baseline recommended conductor size for the effective set count
  selectedSize: string; // Actual conductor size used for displayed results
  isWireSizeForced: boolean;
  warnings: string[];
  actualAmpacity: number;
  voltageDrop: number;
  voltageDropPercentage: number;
  voltageAtLoad: number;
  groundWireSize: string;
  conduitSize: string;
  conduitType: string;
  maxDistanceFor3Percent: number;
  wireAreaTotal: number;
  conduitFillPercentage: number;
  tempRatingUsed: 60 | 75;
  sets: number; // Effective number of parallel sets (forced or recommended)
  recommendedSets: number; // Auto-calculated number of parallel sets
}

export interface AppState {
  voltage: number;
  amperage: number | '';
  distance: number | ''; // in feet
  phase: Phase;
  material: ConductorMaterial;
  maxVoltageDrop: number | ''; // percentage
  sets: number | ''; // Parallel runs
  forceSets: boolean; // Force user-defined sets to override recommendation
  forceWireSize: boolean; // Force a specific conductor size
  forcedWireSize: string; // Exact conductor size to use when forcing wire size
  powerFactor: number;
  oversizeConduit: boolean;
  groundingTable: GroundingTableType;
}
