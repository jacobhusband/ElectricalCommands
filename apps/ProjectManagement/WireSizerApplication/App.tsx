import React, { useState, useEffect } from 'react';
import {
  Settings,
  Zap,
  ArrowRight,
  Cable,
  AlertTriangle,
  Info,
  Ruler,
  Activity,
  Copy,
  CheckCircle2,
} from 'lucide-react';
import { AppState, CalculationResult } from './types';
import { calculateEverything, formatFeederSpec, isThreePhaseAllowed } from './utils/electricalMath';
import { InputGroup } from './components/InputGroup';
import { PieChart, Pie, Cell, ResponsiveContainer, Tooltip as RechartsTooltip } from 'recharts';
import { MIN_RECOMMENDED_CONDUCTOR_SIZE, WIRE_DATA } from './constants';

const inputClass =
  'w-full p-3 border border-gray-300 rounded-lg shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white transition-all text-gray-800 font-medium';
const selectClass =
  'w-full p-3 border border-gray-300 rounded-lg shadow-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 bg-white transition-all text-gray-800 font-medium';
const minRecommendedWireIndex = WIRE_DATA.findIndex(
  ({ size }) => size === MIN_RECOMMENDED_CONDUCTOR_SIZE
);
const selectableWireSizes =
  minRecommendedWireIndex === -1 ? WIRE_DATA : WIRE_DATA.slice(minRecommendedWireIndex);

function App() {
  const [state, setState] = useState<AppState>({
    voltage: 208,
    amperage: 20,
    distance: 100,
    phase: 3,
    material: 'Copper',
    maxVoltageDrop: 3,
    sets: 1,
    forceSets: false,
    forceWireSize: false,
    forcedWireSize: MIN_RECOMMENDED_CONDUCTOR_SIZE,
    powerFactor: 0.9,
    oversizeConduit: false,
    groundingTable: 'EGC',
  });

  const [results, setResults] = useState<CalculationResult | null>(null);
  const [copyFeedback, setCopyFeedback] = useState(false);

  useEffect(() => {
    const res = calculateEverything(state);
    setResults(res);
  }, [state]);

  const handleChange = (field: keyof AppState, value: AppState[keyof AppState]) => {
    setState((prev) => {
      if (field === 'voltage' && typeof value === 'number') {
        const nextPhase = isThreePhaseAllowed(value) ? prev.phase : 1;
        return { ...prev, voltage: value, phase: nextPhase };
      }
      if (field === 'phase' && typeof value === 'number') {
        if (value === 3 && !isThreePhaseAllowed(prev.voltage)) {
          return { ...prev, phase: 1 };
        }
        return { ...prev, phase: value as 1 | 3 };
      }
      return { ...prev, [field]: value };
    });
  };

  const handleNumChange = (field: keyof AppState, value: string) => {
    if (value === '') {
      handleChange(field, '');
      return;
    }
    handleChange(field, Number(value));
  };

  const handleForceWireSizeToggle = (checked: boolean) => {
    setState((prev) => ({
      ...prev,
      forceWireSize: checked,
      forcedWireSize: checked ? results?.selectedSize ?? prev.forcedWireSize : prev.forcedWireSize,
    }));
  };

  const copySpecToClipboard = () => {
    if (!results) return;

    const spec = formatFeederSpec(state, results);
    navigator.clipboard.writeText(spec).then(() => {
      setCopyFeedback(true);
      setTimeout(() => setCopyFeedback(false), 2000);
    });
  };

  const getVoltageStatusColor = (percent: number) => {
    if (percent <= 3) return 'text-green-600';
    if (percent <= 5) return 'text-yellow-600';
    return 'text-red-600';
  };

  const displayedWireSizeOption = state.forceWireSize
    ? selectableWireSizes.some(({ size }) => size === state.forcedWireSize)
      ? state.forcedWireSize
      : MIN_RECOMMENDED_CONDUCTOR_SIZE
    : results?.selectedSize ?? MIN_RECOMMENDED_CONDUCTOR_SIZE;

  return (
    <div className="min-h-screen bg-slate-50 font-sans pb-12 text-slate-900">
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10 shadow-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-blue-600 p-2 rounded-lg">
              <Zap className="text-white w-5 h-5" />
            </div>
            <div>
              <h1 className="text-lg font-bold text-slate-900 leading-tight">ACIES Wire Sizer</h1>
              <p className="text-[10px] text-slate-500 uppercase tracking-widest font-bold">
                Engineering Utility
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={copySpecToClipboard}
              className="flex items-center gap-2 px-4 py-2 text-sm font-bold text-gray-700 bg-white border border-gray-300 rounded-md hover:bg-gray-50 transition-all shadow-sm active:scale-95"
            >
              {copyFeedback ? <CheckCircle2 size={16} className="text-green-600" /> : <Copy size={16} />}
              {copyFeedback ? 'Copied!' : 'Copy Spec'}
            </button>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
        <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
          <div className="lg:col-span-4 space-y-6">
            <div className="bg-white rounded-xl shadow-sm border border-gray-200">
              <div className="bg-slate-50 px-5 py-3 border-b border-gray-200 flex items-center gap-2">
                <Settings className="w-4 h-4 text-gray-500" />
                <h2 className="font-bold text-sm text-gray-700">Circuit Configuration</h2>
              </div>

              <div className="p-5 space-y-4">
                <div className="grid grid-cols-2 gap-3">
                  <InputGroup label="Phase">
                    <select
                      value={state.phase}
                      onChange={(e) => handleChange('phase', Number(e.target.value))}
                      className={selectClass}
                    >
                      <option value={1}>1 Phase</option>
                      <option value={3} disabled={!isThreePhaseAllowed(state.voltage)}>
                        3 Phase
                      </option>
                    </select>
                  </InputGroup>
                  <InputGroup label="Voltage">
                    <select
                      value={state.voltage}
                      onChange={(e) => handleChange('voltage', Number(e.target.value))}
                      className={selectClass}
                    >
                      {[120, 208, 240, 277, 480].map((voltage) => (
                        <option key={voltage} value={voltage}>
                          {voltage}V
                        </option>
                      ))}
                    </select>
                  </InputGroup>
                </div>

                <div className="grid grid-cols-2 gap-3">
                  <InputGroup label="Load (Amps)">
                    <input
                      type="number"
                      value={state.amperage}
                      onChange={(e) => handleNumChange('amperage', e.target.value)}
                      className={inputClass}
                    />
                  </InputGroup>
                  <InputGroup label="Dist. (ft)">
                    <input
                      type="number"
                      value={state.distance}
                      onChange={(e) => handleNumChange('distance', e.target.value)}
                      className={inputClass}
                    />
                  </InputGroup>
                </div>

                <InputGroup label="Conductor">
                  <div className="grid grid-cols-2 bg-gray-100 p-1 rounded-lg">
                    {['Copper', 'Aluminum'].map((material) => (
                      <button
                        key={material}
                        onClick={() => handleChange('material', material as AppState['material'])}
                        className={`py-2 text-sm font-bold rounded-md transition-all ${
                          state.material === material
                            ? 'bg-white shadow-sm text-blue-600'
                            : 'text-gray-500 hover:text-gray-700'
                        }`}
                      >
                        {material}
                      </button>
                    ))}
                  </div>
                </InputGroup>

                <div className="grid grid-cols-2 gap-3">
                  <InputGroup label="V.D. Max %">
                    <input
                      type="number"
                      step="0.1"
                      value={state.maxVoltageDrop}
                      onChange={(e) => handleNumChange('maxVoltageDrop', e.target.value)}
                      className={inputClass}
                    />
                  </InputGroup>
                  <InputGroup label="Number of Parallel Sets">
                    <input
                      type="number"
                      min="1"
                      value={state.sets}
                      onChange={(e) => handleNumChange('sets', e.target.value)}
                      className={inputClass}
                    />
                    {results && state.sets !== '' && state.sets !== results.recommendedSets && (
                      <p className="mt-2 text-xs font-semibold text-red-600">
                        Note: Recommended Number of Parallel Sets is {results.recommendedSets}
                      </p>
                    )}
                    <label className="mt-3 flex items-center gap-2 text-xs font-semibold text-gray-600">
                      <input
                        type="checkbox"
                        checked={state.forceSets}
                        onChange={(e) => handleChange('forceSets', e.target.checked)}
                        className="w-4 h-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                      />
                      Force parallel sets (override recommended)
                    </label>
                  </InputGroup>
                </div>

                <div className="pt-2">
                  <label className="flex items-center gap-3 cursor-pointer group">
                    <input
                      type="checkbox"
                      checked={state.oversizeConduit}
                      onChange={(e) => handleChange('oversizeConduit', e.target.checked)}
                      className="w-5 h-5 rounded border-gray-300 text-blue-600 focus:ring-blue-500 cursor-pointer"
                    />
                    <span className="text-sm font-semibold text-gray-700 group-hover:text-blue-600 transition-colors">
                      Oversize conduit for safety
                    </span>
                  </label>
                </div>

                <InputGroup
                  label="Wire Size Override"
                  subLabel="Force a specific conductor size to see the actual voltage drop and conduit fill for that wire."
                >
                  <div className="space-y-3">
                    <label className="flex items-center gap-2 text-xs font-semibold text-gray-600">
                      <input
                        type="checkbox"
                        checked={state.forceWireSize}
                        onChange={(e) => handleForceWireSizeToggle(e.target.checked)}
                        className="w-4 h-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
                      />
                      Force wire size (override recommended)
                    </label>

                    <select
                      value={displayedWireSizeOption}
                      onChange={(e) => handleChange('forcedWireSize', e.target.value)}
                      disabled={!state.forceWireSize}
                      className={`${selectClass} ${
                        !state.forceWireSize
                          ? 'opacity-60 cursor-not-allowed bg-gray-50 text-gray-400'
                          : ''
                      }`}
                    >
                      {selectableWireSizes.map(({ size }) => (
                        <option key={size} value={size}>
                          #{size}
                        </option>
                      ))}
                    </select>

                    {results &&
                      state.forceWireSize &&
                      results.selectedSize !== results.recommendedSize && (
                        <p className="text-xs font-semibold text-blue-600">
                          Recommended for {results.sets} set{results.sets === 1 ? '' : 's'}: #
                          {results.recommendedSize}
                        </p>
                      )}
                  </div>
                </InputGroup>
              </div>
            </div>
          </div>

          <div className="lg:col-span-8 space-y-6">
            {results && (
              <>
                <div className="bg-white rounded-xl shadow-md border border-gray-200 overflow-hidden">
                  <div className="p-6">
                    <div className="flex flex-col md:flex-row justify-between gap-6 mb-6">
                      <div>
                        <span className="text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-1 block">
                          Feeder Specification
                        </span>
                        <div className="flex items-baseline gap-2">
                          <span className="text-4xl font-black text-slate-900">
                            {results.sets > 1 ? `${results.sets}x ` : ''}
                            {results.selectedSize}
                          </span>
                          <span className="text-lg text-slate-500 font-bold">AWG</span>
                        </div>
                        <p className="text-sm text-slate-500 font-medium mt-1">
                          {state.material} THHN - {state.phase === 3 ? '3PH+N' : '1PH+N'} -{' '}
                          {results.tempRatingUsed}C
                        </p>
                        {results.isWireSizeForced && (
                          <div className="mt-2 flex flex-wrap items-center gap-2">
                            <span className="inline-flex items-center rounded-full bg-amber-100 px-2 py-1 text-[10px] font-bold uppercase tracking-wide text-amber-800">
                              Forced Size
                            </span>
                            {results.selectedSize !== results.recommendedSize && (
                              <span className="text-xs font-semibold text-blue-600">
                                Recommended: #{results.recommendedSize}
                              </span>
                            )}
                          </div>
                        )}
                      </div>

                      <div className="flex gap-4">
                        <div className="bg-slate-50 rounded-lg p-3 border border-slate-100 text-center min-w-[100px]">
                          <span className="text-[10px] font-bold text-slate-400 block mb-1 uppercase">
                            Ground
                          </span>
                          <span className="text-lg font-bold text-slate-700">
                            #{results.groundWireSize}
                          </span>
                          <span className="text-[9px] text-slate-400 block">
                            {state.groundingTable === 'EGC' ? 'T. 250.122' : 'T. 250.66'}
                          </span>
                          <div className="grid grid-cols-2 bg-gray-100 p-0.5 rounded-md mt-2">
                            {(['EGC', 'GEC'] as const).map((groundingTable) => (
                              <button
                                key={groundingTable}
                                onClick={() => handleChange('groundingTable', groundingTable)}
                                className={`py-1 text-[10px] font-bold rounded transition-all ${
                                  state.groundingTable === groundingTable
                                    ? 'bg-white shadow-sm text-blue-600'
                                    : 'text-gray-400 hover:text-gray-600'
                                }`}
                              >
                                {groundingTable}
                              </button>
                            ))}
                          </div>
                        </div>
                        <div className="bg-blue-50 rounded-lg p-3 border border-blue-100 text-center min-w-[100px]">
                          <span className="text-[10px] font-bold text-blue-400 block mb-1 uppercase">
                            Conduit
                          </span>
                          <span className="text-lg font-bold text-blue-700">
                            {results.conduitSize}"
                          </span>
                          {results.sets > 1 && (
                            <span className="text-[9px] text-blue-400 block">per set</span>
                          )}
                        </div>
                      </div>
                    </div>

                    {results.warnings.length > 0 && (
                      <div className="mb-4 rounded-xl border border-amber-200 bg-amber-50 p-4">
                        <div className="flex gap-3">
                          <AlertTriangle className="mt-0.5 h-5 w-5 flex-shrink-0 text-amber-600" />
                          <div className="space-y-1">
                            <p className="text-xs font-bold uppercase tracking-wide text-amber-900">
                              Check forced wire size
                            </p>
                            {results.warnings.map((warning) => (
                              <p key={warning} className="text-sm font-medium text-amber-900">
                                {warning}
                              </p>
                            ))}
                          </div>
                        </div>
                      </div>
                    )}

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div className="border border-slate-100 rounded-xl p-4 bg-slate-50/50">
                        <div className="flex justify-between items-center mb-3">
                          <h4 className="text-xs font-bold text-slate-500 uppercase">
                            Voltage Drop
                          </h4>
                          <span
                            className={`text-xs font-bold ${getVoltageStatusColor(
                              results.voltageDropPercentage
                            )}`}
                          >
                            {results.voltageDropPercentage.toFixed(2)}%
                          </span>
                        </div>
                        <div className="w-full bg-slate-200 rounded-full h-1.5 mb-2">
                          <div
                            className={`h-1.5 rounded-full transition-all duration-500 ${
                              results.voltageDropPercentage > (state.maxVoltageDrop || 3)
                                ? 'bg-red-500'
                                : 'bg-green-500'
                            }`}
                            style={{
                              width: `${Math.min(
                                100,
                                (results.voltageDropPercentage / (state.maxVoltageDrop || 3)) * 100
                              )}%`,
                            }}
                          />
                        </div>
                        <p className="text-[10px] text-slate-400">
                          Limit: {state.maxVoltageDrop || 3}% - At Load:{' '}
                          {results.voltageAtLoad.toFixed(1)}V
                        </p>
                      </div>

                      <div className="border border-slate-100 rounded-xl p-4 bg-slate-50/50">
                        <div className="flex justify-between items-center mb-3">
                          <h4 className="text-xs font-bold text-slate-500 uppercase">
                            Conduit Fill
                          </h4>
                          <span className="text-xs font-bold text-blue-600">
                            {results.conduitFillPercentage.toFixed(1)}% Total
                          </span>
                        </div>
                        <div className="w-full bg-slate-200 rounded-full h-1.5 mb-2">
                          <div
                            className="h-1.5 bg-blue-500 rounded-full transition-all duration-500"
                            style={{
                              width: `${Math.min(100, (results.conduitFillPercentage / 40) * 100)}%`,
                            }}
                          />
                        </div>
                        <p className="text-[10px] text-slate-400">
                          NEC Limit: 40% - Fill Area: {results.wireAreaTotal.toFixed(3)} sq.in
                        </p>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-4">
                  <div className="flex gap-3 text-xs text-gray-500">
                    <Info size={14} className="text-blue-500 flex-shrink-0" />
                    <p>
                      Calculations adhere to <strong>NEC 2023</strong> standards. Terminals assumed
                      at 60C for loads &lt;=100A or wires &lt;=1 AWG, and 75C for loads &gt;100A or wires
                      &gt;=1/0 AWG. Bars show 100% at code-defined limits (3% V.D. and 40% Fill).
                    </p>
                  </div>
                </div>
              </>
            )}
          </div>
        </div>
      </main>
    </div>
  );
}

export default App;
