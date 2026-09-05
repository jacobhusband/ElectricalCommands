# ACIES Wire Sizer

Embedded React / TypeScript feeder calculator. The desktop app loads `dist/index.html`
through `main.py:get_wire_sizer_url`. No AutoCAD command or JSON contract is changed.
The copied specification is plain text for human use.

## Development and delivery

```powershell
npm ci
npm test
npx tsc --noEmit
npm run build
npm run dev
```

Reopen Wire Sizer in the desktop app after rebuilding. Build output is ignored by Git
but is required locally and by the desktop packaging workflow. No API key is needed.
The existing HTML still loads Tailwind and the Inter font from their CDNs.

## Calculation model

- Enter total operating current, the continuous portion of that current, and the actual
  breaker/fuse rating separately. Design current = operating current + 25% of the
  continuous portion. Voltage drop uses operating current.
- Choose the lowest marked terminal rating at either end, 60°C or 75°C. Wire size does
  not select the terminal rating. The default is 60°C.
- General-purpose sizing conservatively requires usable ampacity to cover both the
  design current and the breaker rating. No next-standard-OCPD, 100%-rated assembly,
  motor, transformer, service, tap, or other special exceptions are applied.
- THHN/THWN-2 ampacity starts from Table 310.16. Correct the 90°C column for ambient
  temperature (10–80°C), then adjust for current-carrying conductors. Cap at the
  terminal-column ampacity and the small-conductor protection limit in 240.4(D).
- Each parallel set occupies a SEPARATE EMT raceway containing a complete circuit and
  its own insulated EGC. Other circuits, common raceways, cable bundles, underground
  duct-bank thermal effects, and rooftop temperature adders are not modeled.
- A neutral can be absent, full size, or specified with a user-calculated design load.
  L-N circuits always require a neutral and check it against the circuit requirement.
  A neutral counts for adjustment by default. Exclusion is allowed only where the user
  identifies an eligible imbalance-only neutral; L-N and two-phase wye neutrals are
  always counted. Neutral imbalance and harmonic load calculations remain external.
- General parallel phase and neutral minimum is 1/0 AWG. ACIES single-conductor minimum
  is #12 AWG. Automatic selection minimizes set count first, then conductor size,
  up to 600 kcmil Cu / 750 kcmil Al and 10 sets. Exact-set mode permits 2000 kcmil.
- EGC material is independent of phase/neutral material. Table 250.122 uses the actual
  OCPD rating, up to the currently supported 1600 A table row. Upsizing beyond the
  ampacity minimum proportionally increases EGC circular-mil area under 250.122(B),
  rounding UP to the next available size. The baseline includes the same terminal,
  derating, OCPD, and parallel constraints. No engineering exception is assumed.
- Specified feeder neutrals are also checked against a grounding-based size floor in
  their own material, including the upsizing multiplier. This floor is conservatively
  applied per set; reduced parallel-neutral exceptions are not implemented.
- Table 250.66 GEC sizing is a SEPARATE reference based on equivalent phase conductor
  area. It never replaces the feeder EGC or enters the feeder's conduit fill / copied
  specification. Electrode-specific limits and service bonding are not modeled.
- Voltage drop uses K=12.9 Cu / 21.2 Al and sqrt(3) for balanced three-phase L-L or 2
  for single-phase L-L. A two-wire L-N circuit uses the actual phase PLUS neutral
  resistance. These are resistance approximations, not AC impedance / power-factor
  calculations. The 3% default is a design target, not a universal code requirement.
- Fill uses the existing generic insulated-conductor areas and EMT 40% areas, with a
  3/4-inch ACIES minimum. All supported circuits have at least three conductors with
  the included insulated EGC. Confirm actual construction/insulation areas for the
  selected product, especially compact aluminum.

## Failure behavior

Invalid or incomplete inputs suppress calculations. Unsatisfied capacity, voltage drop,
parallel minimum, neutral sizing, grounding-table range, or conduit fill limits generate
explicit errors and disable Copy Spec. The largest candidate may remain visible for
troubleshooting, labeled as a candidate. Unknown ground area produces unresolved fill,
never a guessed area. GEC reference values on an invalid candidate are not an approved
service design.

## References

The model targets the existing NEC 2023 edition; select the locally adopted code and
project requirements separately. References used to cross-check the implementation:

- [Schneider terminal-rating guidance, 110.14(C)](https://www.se.com/us/en/faqs/FA144510/)
- [Schneider correction and adjustment tables](https://productinfo.se.com/na-std-ref/viewer/5c0fee4b347bdf0001de4d55/5c0fee5b347bdf0001de4d6b/r/CorrectionAndAdjustmentFactors-F20BC8E5)
- [Southwire copper ampacities, SPEC 45002](https://cabletechsupport.southwire.com/en/cablespec/download_spec/?spec=45002)
- [Eaton EGC sizing based on OCPD trip rating](https://www.eaton.com/content/dam/eaton/products/design-guides---consultant-audience/canada/cag/eaton-power-distribution-systems-consulting-application-guide-tb08104003e-tab-1-ca08104001e-ca.pdf)
- NEC 310.16, 310.15(B)(1), 310.15(C)(1), 310.15(E), 310.10(G), 215.2(B),
  240.4(D), 250.122, 250.66, and Chapter 9 Tables 4 and 5.

Regression tests cover the original screenshot, search exhaustion, overfill, forced
undersizing, EGC upsize rounding, separate grounding material, continuous loads,
terminal selection, correction factors, parallel and neutral constraints, invalid input,
zero length, and copy formatting / refusal.
