import { WireSizeData, GroundingData, GECData } from './types';

// Simplified Table 310.16 (60°C and 75°C Columns) + Physical Properties
export const WIRE_DATA: WireSizeData[] = [
  { size: "14", circularMils: 4110, ampacity60Cu: 15, ampacity75Cu: 20, ampacity60Al: 0, ampacity75Al: 0, areaSqIn: 0.0097 },
  { size: "12", circularMils: 6530, ampacity60Cu: 20, ampacity75Cu: 25, ampacity60Al: 15, ampacity75Al: 20, areaSqIn: 0.0133 },
  { size: "10", circularMils: 10380, ampacity60Cu: 30, ampacity75Cu: 35, ampacity60Al: 25, ampacity75Al: 30, areaSqIn: 0.0211 },
  { size: "8", circularMils: 16510, ampacity60Cu: 40, ampacity75Cu: 50, ampacity60Al: 30, ampacity75Al: 40, areaSqIn: 0.0366 },
  { size: "6", circularMils: 26240, ampacity60Cu: 55, ampacity75Cu: 65, ampacity60Al: 40, ampacity75Al: 50, areaSqIn: 0.0507 },
  { size: "4", circularMils: 41740, ampacity60Cu: 70, ampacity75Cu: 85, ampacity60Al: 55, ampacity75Al: 65, areaSqIn: 0.0824 },
  { size: "3", circularMils: 52620, ampacity60Cu: 85, ampacity75Cu: 100, ampacity60Al: 65, ampacity75Al: 75, areaSqIn: 0.0973 },
  { size: "2", circularMils: 66360, ampacity60Cu: 95, ampacity75Cu: 115, ampacity60Al: 75, ampacity75Al: 90, areaSqIn: 0.1158 },
  { size: "1", circularMils: 83690, ampacity60Cu: 110, ampacity75Cu: 130, ampacity60Al: 85, ampacity75Al: 100, areaSqIn: 0.1562 },
  { size: "1/0", circularMils: 105600, ampacity60Cu: 125, ampacity75Cu: 150, ampacity60Al: 100, ampacity75Al: 120, areaSqIn: 0.1855 },
  { size: "2/0", circularMils: 133100, ampacity60Cu: 145, ampacity75Cu: 175, ampacity60Al: 115, ampacity75Al: 135, areaSqIn: 0.2223 },
  { size: "3/0", circularMils: 167800, ampacity60Cu: 165, ampacity75Cu: 200, ampacity60Al: 130, ampacity75Al: 155, areaSqIn: 0.2679 },
  { size: "4/0", circularMils: 211600, ampacity60Cu: 195, ampacity75Cu: 230, ampacity60Al: 150, ampacity75Al: 180, areaSqIn: 0.3237 },
  { size: "250", circularMils: 250000, ampacity60Cu: 215, ampacity75Cu: 255, ampacity60Al: 170, ampacity75Al: 205, areaSqIn: 0.3970 },
  { size: "300", circularMils: 300000, ampacity60Cu: 240, ampacity75Cu: 285, ampacity60Al: 190, ampacity75Al: 230, areaSqIn: 0.4608 },
  { size: "350", circularMils: 350000, ampacity60Cu: 260, ampacity75Cu: 310, ampacity60Al: 210, ampacity75Al: 250, areaSqIn: 0.5242 },
  { size: "400", circularMils: 400000, ampacity60Cu: 280, ampacity75Cu: 335, ampacity60Al: 225, ampacity75Al: 270, areaSqIn: 0.5863 },
  { size: "500", circularMils: 500000, ampacity60Cu: 320, ampacity75Cu: 380, ampacity60Al: 260, ampacity75Al: 310, areaSqIn: 0.7073 },
  { size: "600", circularMils: 600000, ampacity60Cu: 350, ampacity75Cu: 420, ampacity60Al: 285, ampacity75Al: 340, areaSqIn: 0.8676 },
  { size: "700", circularMils: 700000, ampacity60Cu: 385, ampacity75Cu: 460, ampacity60Al: 310, ampacity75Al: 375, areaSqIn: 0.9887 },
  { size: "750", circularMils: 750000, ampacity60Cu: 400, ampacity75Cu: 475, ampacity60Al: 320, ampacity75Al: 385, areaSqIn: 1.0496 },
  { size: "800", circularMils: 800000, ampacity60Cu: 410, ampacity75Cu: 490, ampacity60Al: 330, ampacity75Al: 395, areaSqIn: 1.1085 },
  { size: "900", circularMils: 900000, ampacity60Cu: 435, ampacity75Cu: 520, ampacity60Al: 355, ampacity75Al: 425, areaSqIn: 1.2311 },
  { size: "1000", circularMils: 1000000, ampacity60Cu: 455, ampacity75Cu: 545, ampacity60Al: 375, ampacity75Al: 445, areaSqIn: 1.3478 },
  { size: "1250", circularMils: 1250000, ampacity60Cu: 495, ampacity75Cu: 590, ampacity60Al: 405, ampacity75Al: 485, areaSqIn: 1.718 },
  { size: "1500", circularMils: 1500000, ampacity60Cu: 525, ampacity75Cu: 625, ampacity60Al: 435, ampacity75Al: 520, areaSqIn: 2.0157 },
  { size: "1750", circularMils: 1750000, ampacity60Cu: 545, ampacity75Cu: 650, ampacity60Al: 455, ampacity75Al: 545, areaSqIn: 2.3127 },
  { size: "2000", circularMils: 2000000, ampacity60Cu: 555, ampacity75Cu: 665, ampacity60Al: 470, ampacity75Al: 560, areaSqIn: 2.6073 },
];

// Table 250.122 Minimum Size Equipment Grounding Conductors
export const GROUND_DATA: GroundingData[] = [
  { rating: 15, cuSize: "14", alSize: "12" },
  { rating: 20, cuSize: "12", alSize: "10" },
  { rating: 60, cuSize: "10", alSize: "8" },
  { rating: 100, cuSize: "8", alSize: "6" },
  { rating: 200, cuSize: "6", alSize: "4" },
  { rating: 300, cuSize: "4", alSize: "2" },
  { rating: 400, cuSize: "3", alSize: "1" },
  { rating: 500, cuSize: "2", alSize: "1/0" },
  { rating: 600, cuSize: "1", alSize: "2/0" },
  { rating: 800, cuSize: "1/0", alSize: "3/0" },
  { rating: 1000, cuSize: "2/0", alSize: "4/0" },
  { rating: 1200, cuSize: "3/0", alSize: "250" },
  { rating: 1600, cuSize: "4/0", alSize: "350" },
];

// Table 250.66 Grounding Electrode Conductor for AC Systems
export const GEC_DATA: GECData[] = [
  { maxCuCM: 66360,    maxAlCM: 105600,   gecCuSize: "8",   gecAlSize: "6" },
  { maxCuCM: 105600,   maxAlCM: 167800,   gecCuSize: "6",   gecAlSize: "4" },
  { maxCuCM: 167800,   maxAlCM: 250000,   gecCuSize: "4",   gecAlSize: "2" },
  { maxCuCM: 350000,   maxAlCM: 500000,   gecCuSize: "2",   gecAlSize: "1/0" },
  { maxCuCM: 600000,   maxAlCM: 900000,   gecCuSize: "1/0", gecAlSize: "3/0" },
  { maxCuCM: 1100000,  maxAlCM: 1750000,  gecCuSize: "2/0", gecAlSize: "4/0" },
  { maxCuCM: Infinity, maxAlCM: Infinity,  gecCuSize: "3/0", gecAlSize: "250" },
];

// Simplified internal area of EMT Conduit (Trade Size -> Sq Inches @ 40% fill)
// Using Chapter 9 Table 4 (40% fill column for over 2 wires)
export const CONDUIT_EMT_40_PERCENT: { size: string; area: number }[] = [
  { size: "1/2", area: 0.122 },
  { size: "3/4", area: 0.213 },
  { size: "1", area: 0.346 },
  { size: "1-1/4", area: 0.598 },
  { size: "1-1/2", area: 0.814 },
  { size: "2", area: 1.342 },
  { size: "2-1/2", area: 1.924 },
  { size: "3", area: 2.907 },
  { size: "3-1/2", area: 3.848 },
  { size: "4", area: 4.947 },
];

export const K_FACTOR_CU = 12.9;
export const K_FACTOR_AL = 21.2;
export const MIN_RECOMMENDED_CONDUCTOR_SIZE = "12";

// Maximum wire sizes by material for parallel set calculations
export const MAX_WIRE_SIZE_CU = "600"; // 600 kcmil max for copper
export const MAX_WIRE_SIZE_AL = "750"; // 750 kcmil max for aluminum
export const MAX_WIRE_SIZE_FORCED = "2000"; // Max allowed when forcing sets

// NEC Table 310.16, 90°C column, for THHN/THWN-2 correction and adjustment ONLY.
// Final allowable ampacity is capped by the selected terminal rating and 240.4(D).
// Reference: Southwire MC Installation Guide (2024), ampacity tables.
export const AMPACITY_90: Record<string, { Copper: number; Aluminum: number }> = {
  '14': { Copper: 25, Aluminum: 0 }, '12': { Copper: 30, Aluminum: 25 },
  '10': { Copper: 40, Aluminum: 35 }, '8': { Copper: 55, Aluminum: 45 },
  '6': { Copper: 75, Aluminum: 55 }, '4': { Copper: 95, Aluminum: 75 },
  '3': { Copper: 115, Aluminum: 85 }, '2': { Copper: 130, Aluminum: 100 },
  '1': { Copper: 150, Aluminum: 115 }, '1/0': { Copper: 170, Aluminum: 135 },
  '2/0': { Copper: 195, Aluminum: 150 }, '3/0': { Copper: 225, Aluminum: 175 },
  '4/0': { Copper: 260, Aluminum: 205 }, '250': { Copper: 290, Aluminum: 230 },
  '300': { Copper: 320, Aluminum: 260 }, '350': { Copper: 350, Aluminum: 280 },
  '400': { Copper: 380, Aluminum: 305 }, '500': { Copper: 430, Aluminum: 350 },
  '600': { Copper: 475, Aluminum: 385 }, '700': { Copper: 520, Aluminum: 425 },
  '750': { Copper: 535, Aluminum: 435 }, '800': { Copper: 555, Aluminum: 445 },
  '900': { Copper: 585, Aluminum: 480 }, '1000': { Copper: 615, Aluminum: 500 },
  '1250': { Copper: 665, Aluminum: 545 }, '1500': { Copper: 705, Aluminum: 585 },
  '1750': { Copper: 735, Aluminum: 615 }, '2000': { Copper: 750, Aluminum: 630 },
};

// Table 310.15(B)(1), 90°C insulation, ambient range 10–80°C.
export const AMBIENT_FACTORS_90 = [
  [10, 1.15], [15, 1.12], [20, 1.08], [25, 1.04], [30, 1],
  [35, 0.96], [40, 0.91], [45, 0.87], [50, 0.82], [55, 0.76],
  [60, 0.71], [65, 0.65], [70, 0.58], [75, 0.50], [80, 0.41],
];
