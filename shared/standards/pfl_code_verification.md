# PFL Audit Catalog Verification Record

Verified: 2026-08-18  
Catalog: `master_audit_catalog.json` version 3.0.0

## Verification basis

This review checks whether each PFL item is a code requirement, a conditional code review, or an ACIES/client coordination standard. It uses the adopted-code text and official agency material available on the verification date. PFL is a scoping and QA/QC aid: the permit application date, enforcing agency, local amendments, project facts, listed equipment, and AHJ interpretations remain controlling.

Primary sources:

- [California Building Standards Commission code editions](https://www.dgs.ca.gov/BSC/Codes)
- [2025 California Electrical Code errata/history (2023 NEC basis; effective January 1, 2026)](https://www.dgs.ca.gov/-/media/Divisions/BSC/02-Codes/Errta-01012026-Pt3-CEC-locked.pdf)
- [2025 local ordinances filed with CBSC](https://www.dgs.ca.gov/BSC/Codes/2025-Ordinances)
- [2022 California Energy Code, Subchapter 4: Sections 130.0-130.5](https://codes.iccsafe.org/content/CAEC2022P2/subchapter-4-nonresidential-and-hotel-motel-occupancies-mandatory-requirements-for-lighting-systems-and-equipment-and-electrical-power-distribution-systems)
- [2022 CBC Chapter 9: fire protection and duct-smoke interfaces](https://codes.iccsafe.org/content/CABC2022P3/chapter-9-fire-protection-and-life-safety-systems)
- [2022 CBC Chapter 10: egress illumination and exit signs](https://codes.iccsafe.org/content/CABC2022P1/chapter-10-means-of-egress)
- [NFPA 70/NEC development and public records](https://www.nfpa.org/codes-and-standards/nfpa-70-standard-development/70)
- [NFPA 96 development and public records](https://www.nfpa.org/codes-and-standards/nfpa-96-standard-development/96)
- [Bank of America prototype baseline used by ACIES](bank_of_america_prototype_specs.md)

## Original 43-rule disposition

| Original rule ID | Result | Verification / change |
|---|---|---|
| `gen_setup_reference_manager` | ACIES QA retained | Not represented as a code mandate. Relative/clean xrefs remain a drawing-delivery check. |
| `gen_setup_clean_titleblock` | ACIES QA retained | Not represented as a code mandate. |
| `gen_setup_spec_sheet_state` | Corrected | Code cycle now depends on permit basis and AHJ. Notes distinguish 2022 CEC/2020 NEC from 2025 CEC/2023 NEC. |
| `gen_setup_symbols_legend_match` | ACIES QA retained | Not represented as a code mandate. |
| `elec_sld_switchboard_kaic` | Corrected | Existing equipment is not categorically exempt from adequate ratings; 110.24 calculation/marking scope and exceptions must be evaluated separately. |
| `elec_sld_high_fault_current_coordination` | Corrected | Removed the false implication that 10 kA is a code trigger. Compare calculated fault current with each device rating; document tested series combinations under 240.86 when used. |
| `elec_sld_panel_aic_ratings` | Verified and scoped | Applies only when service/distribution work makes fault-current review relevant. Transformer impedance is a calculation input, not an automatic safe harbor. |
| `elec_sld_voltage_drop_5percent` | Corrected | Title 24 §130.5(c) uses a combined 5% maximum for covered work. The familiar 2% feeder/3% branch split is a design target, while NEC language is generally informational unless adopted elsewhere. |
| `elec_sld_grounding_electrode` | Verified and scoped | Service and separately derived system work trigger the review; tenant subpanels do not automatically receive new electrodes. |
| `elec_sld_spd_service_entrance` | Corrected | Article 242 is not a blanket commercial-service SPD mandate. PFL now triggers dwelling/sleeping-occupancy service requirements by cycle and the separate Article 700 SPD requirement. |
| `elec_sld_feeder_conduit_fill` | Corrected | Forty-percent fill applies to more than two conductors; 60% nipple rule retained. A wire EGC is not universally required when a recognized metal raceway supplies the grounding path. |
| `elec_pwr_panel_clearance_space` | Corrected | Working depth depends on voltage to ground and Table 110.26 conditions, not voltage alone. Added headroom, egress, illumination, and dedicated-space nuances. |
| `elec_pwr_service_receptacle_25ft` | Corrected | Restricted to covered service-equipment work; no longer applies to every panelboard. A qualifying existing receptacle can comply and a dedicated circuit is not universally required. |
| `elec_pwr_controlled_receptacles_t24` | Corrected | Removed the blanket commercial “50%” claim. Section 130.5(d) primarily uses proximity/split-outlet/workstation rules; hotel guest rooms have a separate half-controlled rule. Added equipment, patient-care, and alteration exceptions. |
| `elec_pwr_tamper_resistant_receptacles` | Corrected | Public access, retail, or business occupancy alone is not the trigger. PFL now uses the occupancies/spaces listed in 406.12 and preserves its height, appliance, and dedicated-equipment exceptions. |
| `elec_pwr_gfci_protection_locations` | Corrected | Uses the adopted 210.8 location, voltage, current, phase, and appliance scope. Dedicated status is not a general exception; 2020- and 2023-NEC appliance coverage differs. |
| `elec_pwr_modular_furniture_poc` | Verified and scoped | Article 605 coordination retained. Title 24 controlled-outlet requirement is applied only when its separate predicate is met. |
| `elec_pwr_meeting_room_floor_boxes` | Corrected | The original threshold was reversed. Section 210.65 covers meeting rooms **not more than** 1,000 sq ft; the central floor-outlet geometry is now a separate question/check. |
| `elec_pwr_hand_dryer_circuit` | Corrected | Removed the universal dedicated 20 A requirement. Circuiting, connection, GFCI, and disconnect follow nameplate, load, listing, location, and adopted cycle. |
| `elec_pwr_food_disposer_switch_outlet` | Corrected | Removed mandatory wall-toggle/under-counter topology. PFL now accepts listed cord-and-plug, hardwired, air-switch, and other compliant arrangements. |
| `elec_pwr_commercial_kitchen_shunt_trip` | Corrected | NFPA 96 requires automatic shutdown of covered fuel and electric heat-producing equipment; shunt trip is one implementation, not the mandate itself. |
| `elec_pwr_grease_trap_alarm` | Corrected | A grease interceptor alone does not prove an electrical load. Review is now equipment-data and local-health-requirement coordination. |
| `elec_pwr_water_heater_disconnect` | Corrected | Section 422.13's 125% rule is limited to qualifying storage water heaters of 120 gal or less. Removed the arbitrary 1,500 W point-of-use exception. |
| `mech_coord_rtu_receptacle_disconnect` | Corrected | Section 210.63(A)/Article 440 applies to HVAC/R equipment. Non-HVAC roof motors are now routed to Article 430 instead. |
| `mech_coord_rtu_rcp_dashed_representation` | ACIES QA retained | Clearly identified as drafting coordination, not a code mandate. |
| `mech_coord_duct_smoke_detectors` | Corrected | General threshold is **over** 2,000 CFM. Replaced the erroneous CBC 907.2.12.1.2 citation with CMC 609.1/609.1.1 and CBC 907.3.1; alternative detection now requires its own documentation check. |
| `mech_coord_mca_mocp_cross_check` | Corrected | Uses nameplate MCA/MOCP and Article 440 rather than generic Article 430 percentages; conductor ampacity must meet MCA and OCPD must not exceed MOCP. |
| `elec_ev_charging_infrastructure` | Corrected | Article 625 installed-EVSE design is separated from CALGreen future-parking infrastructure. Listed energy management can set the permitted maximum load. |
| `elec_solar_pv_backfed_breaker` | Corrected | PV and BESS are separated. The 120% busbar method is one Article 705 load-side option, not a universal PV/storage formula; supply-side and listed PCS paths remain possible. |
| `ltg_t24_demand_response_4000w` | Corrected | Trigger is general lighting **over** 4,000 W, not 4,000 W or more. Health/life-safety exclusions now create a documentation check. |
| `ltg_t24_daylight_harvesting_zones` | Corrected | Removed the false 100-sq-ft trigger. The central threshold is 120 W in an applicable daylit zone, with documented obstruction, overhang, glazing, parking, retail, and other cycle-specific exceptions. |
| `ltg_t24_occupancy_sensor_mandatory_spaces` | Corrected | Removed the claim that every enclosed commercial space uses one sensor behavior. PFL now asks whether regulated spaces exist and directs review to the cycle-specific control mode and exceptions. |
| `ltg_t24_restroom_ceiling_sensor` | Verified as mixed basis | Automatic shutoff is code-based; ceiling/dual-tech placement in multi-stall rooms remains explicitly identified as ACIES coverage QA. |
| `ltg_t24_multilevel_dimming_controls` | Corrected | Trigger is room area at least 100 sq ft and LPD over 0.5 W/sq ft. Restroom and single-luminaire exceptions retained; 0-10 V is not the only compliant method. |
| `ltg_t24_timeclock_2hr_override` | Corrected | Recast as automatic shutoff/scheduling. Occupant sensing can be the compliance method; a timeclock is not mandatory in every interior space. |
| `ltg_t24_exterior_photocell_astronomical` | Corrected | Exterior control method now depends on application. Emergency/egress and other exceptions generate a documentation check rather than suppressing all exterior review. |
| `ltg_em_inverter_circuit_unswitched` | Corrected | Tied to local normal-branch failure and the listed UL 924 sequence, not a generic unswitched-wire prescription for every topology. |
| `ltg_em_inverter_cross_coordination` | Verified and expanded | Retains ACIES cross-sheet coordination; added Article 700 capacity, separation, transfer, fault-current, listing, and environmental considerations. |
| `ltg_em_battery_units_90min` | Corrected | Unit equipment is 700.12(I) in the 2020 NEC basis and 700.12(H) in the 2023 NEC basis. Retains 90-minute and local normal-lighting circuit requirements with permitted arrangements. |
| `ltg_em_exit_sign_single_vs_double_face` | Corrected | CBC 1013.1 governs sign location/visibility. Single/double face is derived from sightlines and product coordination, not a fixed code face-count rule. |
| `ltg_em_egress_illumination_1fc` | Corrected | Adds the complete CBC 1008.3.5 criteria: initial 1.0 fc average/0.1 fc minimum; 90-minute 0.6 fc average/0.06 fc minimum; 40:1 maximum uniformity. |
| `bofa_std_fixture_catalog_3500k` | Client standard retained | Verified against the local ACIES BofA prototype source and kept separate from code mandates. |
| `bofa_std_wattstopper_dlm_controls` | Client standard retained | Verified against the local ACIES BofA prototype source and kept separate from code mandates. |

## Added coverage

Version 3.0 adds conditional checks for:

- permit date, AHJ, local amendments, and California lighting-alteration path;
- Title 24 metering/load separation and exception documentation;
- arc-flash marking, large-service labeling, arc-energy reduction, transformers, and high-leg delta systems;
- separate meeting-room floor-outlet geometry, roof motor circuits, and duct-smoke alternatives;
- CALGreen EV infrastructure independent of installed EVSE;
- BESS electrical/fire-code review independent of PV;
- demand-response, daylight, controlled-receptacle, and exterior-control exception documentation;
- Title 24 lighting-control acceptance testing;
- separate exit-sign and emergency-illumination facts;
- Article 700, 701, and 702 classification, generators, and selective coordination;
- fire alarm, fire pump, elevators/conveyances, Article 517 health care, hazardous locations, pools/spas/fountains, and AFCI coverage.

## Operational limits

- PFL selects review topics; it does not calculate fault current, voltage drop, photometrics, feeder ampacity, EV counts, or battery fire-code thresholds.
- A checked exception means “verify and document the exception,” not “ignore the subject.”
- Local ordinances, utility standards, fire-department criteria, HCAI/DSA jurisdiction, equipment listings, and delegated-design documents must be entered or reviewed by the engineer.
- Rule citations are starting points. The adopted edition and all referenced subsections, tables, definitions, exceptions, and California/local amendments control.
