using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ElectricalCommands
{
  public static class AuditEngine
  {
    public const string AuditStateFileName = "project_audit_state.json";
    private static List<AuditRuleDefinition> _cachedCatalog = null;

    public static List<AuditRuleDefinition> GetMasterCatalog()
    {
      if (_cachedCatalog != null && _cachedCatalog.Count > 0)
      {
        return _cachedCatalog;
      }

      var rules = new List<AuditRuleDefinition>();

      // 1. Try loading from Embedded Resource
      try
      {
        var assembly = Assembly.GetExecutingAssembly();
        string resourceName = assembly.GetManifestResourceNames()
          .FirstOrDefault(r => r.EndsWith("master_audit_catalog.json", StringComparison.OrdinalIgnoreCase));

        if (!string.IsNullOrWhiteSpace(resourceName))
        {
          using (var stream = assembly.GetManifestResourceStream(resourceName))
          {
            if (stream != null)
            {
              using (var reader = new StreamReader(stream, Encoding.UTF8))
              {
                string json = reader.ReadToEnd();
                var catalog = JsonConvert.DeserializeObject<MasterAuditCatalogFile>(json);
                if (catalog?.Rules != null && catalog.Rules.Count > 0)
                {
                  rules.AddRange(catalog.Rules);
                }
              }
            }
          }
        }
      }
      catch
      {
      }

      // 2. Fallback to adjacent file if running during development
      if (rules.Count == 0)
      {
        try
        {
          string loc = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
          string directPath = Path.Combine(loc ?? string.Empty, "master_audit_catalog.json");
          if (File.Exists(directPath))
          {
            string json = File.ReadAllText(directPath, Encoding.UTF8);
            var catalog = JsonConvert.DeserializeObject<MasterAuditCatalogFile>(json);
            if (catalog?.Rules != null && catalog.Rules.Count > 0)
            {
              rules.AddRange(catalog.Rules);
            }
          }
        }
        catch
        {
        }
      }

      _cachedCatalog = rules;
      return rules;
    }

    public static List<AuditRuleDefinition> FilterRules(
      IEnumerable<AuditRuleDefinition> allRules,
      ProjectScopeConfig scope
    )
    {
      if (allRules == null)
      {
        return new List<AuditRuleDefinition>();
      }

      var active = new List<AuditRuleDefinition>();
      var derivedFlags = DeriveInferredFlags(scope);

      foreach (var rule in allRules)
      {
        if (IsRuleApplicable(rule, scope, derivedFlags))
        {
          active.Add(rule);
        }
      }

      return active.OrderBy(r => r.Order).ThenBy(r => r.Title).ToList();
    }

    public static Dictionary<string, bool> DeriveInferredFlags(ProjectScopeConfig scope)
    {
      var flags = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
      if (scope == null) return flags;

      bool isCalifornia = (scope.CodeJurisdiction ?? string.Empty)
        .StartsWith("California", StringComparison.OrdinalIgnoreCase);
      bool uses2023Nec = (scope.CodeJurisdiction ?? string.Empty)
        .IndexOf("2023", StringComparison.OrdinalIgnoreCase) >= 0 ||
        (scope.CodeJurisdiction ?? string.Empty)
        .IndexOf("CEC 2025", StringComparison.OrdinalIgnoreCase) >= 0;
      bool isNewBuildingOrAddition =
        string.Equals(scope.ConstructionScopeNature, "New Building / Addition", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(scope.ConstructionScopeNature, "First-Time Tenant Build-Out / Shell Completion", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(scope.ConstructionScopeNature, "New Construction / 1st Time Tenant Space", StringComparison.OrdinalIgnoreCase);
      bool hasServiceWork = !string.Equals(
        scope.ElectricalServiceScope,
        "Existing Service - No Service Work",
        StringComparison.OrdinalIgnoreCase
      ) && !string.Equals(
        scope.ElectricalServiceScope,
        "No Electrical Service in Scope",
        StringComparison.OrdinalIgnoreCase
      );
      bool completeDistribution = string.Equals(
        scope.ElectricalDistributionScope,
        "Entirely New / Complete Replacement",
        StringComparison.OrdinalIgnoreCase
      );
      bool hasDistributionWork = completeDistribution || string.Equals(
        scope.ElectricalDistributionScope,
        "Partial Alteration / Tenant Distribution",
        StringComparison.OrdinalIgnoreCase
      );
      bool fullLightingControls = string.Equals(
        scope.LightingComplianceScope,
        "New / Complete Lighting System - Full Controls",
        StringComparison.OrdinalIgnoreCase
      ) || string.Equals(
        scope.LightingComplianceScope,
        "Alteration - Full Controls per Table 141.0-F",
        StringComparison.OrdinalIgnoreCase
      );
      bool hasRegulatedLightingWork = !string.Equals(
        scope.LightingComplianceScope,
        "No Regulated Lighting Work",
        StringComparison.OrdinalIgnoreCase
      );
      bool lightingScopeUnknown = string.Equals(
        scope.LightingComplianceScope,
        "Unknown - Verify Alteration Path",
        StringComparison.OrdinalIgnoreCase
      );

      bool occupancyDwelling = scope.HasDwellingUnits ||
        (scope.BuildingOccupancy ?? string.Empty).IndexOf("Dwelling", StringComparison.OrdinalIgnoreCase) >= 0;
      bool occupancyHotelDorm = scope.HasHotelDormitoryOrAssistedLiving ||
        (scope.BuildingOccupancy ?? string.Empty).IndexOf("Hotel", StringComparison.OrdinalIgnoreCase) >= 0 ||
        (scope.BuildingOccupancy ?? string.Empty).IndexOf("Dormitory", StringComparison.OrdinalIgnoreCase) >= 0;
      bool occupancyHealthCare = scope.HasHealthCareAreas ||
        (scope.BuildingOccupancy ?? string.Empty).IndexOf("Health", StringComparison.OrdinalIgnoreCase) >= 0;
      bool occupancyEducation = scope.HasChildCareOrEducationAreas ||
        (scope.BuildingOccupancy ?? string.Empty).IndexOf("Education", StringComparison.OrdinalIgnoreCase) >= 0;
      bool occupancyAssembly = scope.HasAssemblyAreas ||
        (scope.BuildingOccupancy ?? string.Empty).IndexOf("Assembly", StringComparison.OrdinalIgnoreCase) >= 0;
      bool occupancyHazardous = scope.HasHazardousClassifiedLocations ||
        (scope.BuildingOccupancy ?? string.Empty).IndexOf("Hazardous", StringComparison.OrdinalIgnoreCase) >= 0;

      flags["IsCaliforniaProject"] = isCalifornia;
      flags["Uses2023Nec"] = uses2023Nec;
      bool codeStatusConfirmed = string.Equals(
          scope.CodeBasisStatus,
          "Confirmed - No Local Amendments",
          StringComparison.OrdinalIgnoreCase
        ) || string.Equals(
          scope.CodeBasisStatus,
          "Confirmed - Local Amendments Apply",
          StringComparison.OrdinalIgnoreCase
        );
      flags["NeedsCodeBasisConfirmation"] = !codeStatusConfirmed ||
        string.IsNullOrWhiteSpace(scope.AuthorityHavingJurisdiction) ||
        string.IsNullOrWhiteSpace(scope.PermitApplicationDate);
      flags["NeedsLocalAmendmentReview"] = scope.HasLocalAmendments || string.Equals(
        scope.CodeBasisStatus,
        "Confirmed - Local Amendments Apply",
        StringComparison.OrdinalIgnoreCase
      );
      flags["NeedsLightingScopeResolution"] = isCalifornia && lightingScopeUnknown;

      // Service, distribution, and overcurrent protection.
      flags["NeedsElectricalServiceReview"] = hasServiceWork;
      flags["NeedsDistributionEquipmentReview"] = hasServiceWork || hasDistributionWork;
      flags["NeedsDistributionFaultCurrentReview"] = hasServiceWork || hasDistributionWork;
      flags["NeedsFeederReview"] = hasDistributionWork;
      flags["NeedsGroundingElectrodeReview"] = hasServiceWork || scope.HasNewTransformers;
      flags["NeedsHighFaultCurrentReview"] = string.Equals(
        scope.FaultCurrentLevel,
        "> 10kAIC / High Fault Current (22k-65kAIC)",
        StringComparison.OrdinalIgnoreCase
      ) && (hasServiceWork || hasDistributionWork);
      bool expanded2023SpdOccupancy = uses2023Nec && scope.Has2023NecSpdSleepingOccupancy;
      flags["NeedsDwellingServiceSpd"] = hasServiceWork && (occupancyDwelling || expanded2023SpdOccupancy);
      flags["NeedsServiceEquipmentReceptacle"] = hasServiceWork;
      flags["NeedsArcFlashWarning"] = (hasServiceWork || hasDistributionWork) && !occupancyDwelling;
      flags["NeedsServiceIncidentEnergyLabel"] = hasServiceWork &&
        scope.HasService1200AmpsOrMore && !occupancyDwelling;
      flags["NeedsArcEnergyReduction"] = (hasServiceWork || hasDistributionWork) && scope.HasOcpd1200AmpsOrMore;
      flags["NeedsTransformerReview"] = scope.HasNewTransformers;
      flags["NeedsHighLegDeltaReview"] = (
        string.Equals(scope.ServiceVoltage, "120/240V 3PH 4W High-Leg Delta", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(scope.ServiceVoltage, "120/240V 3PH Delta", StringComparison.OrdinalIgnoreCase)
      ) && (hasServiceWork || hasDistributionWork);
      flags["NeedsTitle24Metering"] = isCalifornia && hasServiceWork;
      flags["NeedsTitle24LoadSeparation"] = isCalifornia && completeDistribution;
      flags["NeedsVoltageDropReview"] = hasDistributionWork;

      // Receptacles, rooms, and utilization equipment.
      bool controlledReceptacleSpace = scope.HasOfficeSpaces || scope.HasHotelMotelGuestRooms;
      bool controlledReceptacleProject = isNewBuildingOrAddition || completeDistribution;
      bool needsControlled = isCalifornia && controlledReceptacleSpace &&
        controlledReceptacleProject && !scope.HasControlledReceptacleException;
      flags["NeedsControlledReceptacles"] = needsControlled;
      flags["ControlledReceptaclesExempt"] = isCalifornia && controlledReceptacleSpace &&
        (!controlledReceptacleProject || scope.HasControlledReceptacleException);
      flags["NeedsTamperResistant"] = occupancyDwelling || occupancyHotelDorm || occupancyEducation ||
        occupancyHealthCare || occupancyAssembly;
      flags["NeedsGfciProtection"] = scope.HasSinksOrWetLocations || scope.HasMultiStallRestrooms ||
        scope.HasCommercialKitchen || scope.HasGaragesServiceBaysOrIndoorParking || scope.HasExteriorLighting ||
        scope.HasRooftopUnits || scope.HasPoolsSpasOrFountains;
      flags["NeedsModularFurniturePoc"] = scope.HasModularFurniture;
      flags["NeedsMeetingRoomReceptacles"] = scope.HasMeetingRoomsAtOrBelow1000SqFt ||
        scope.HasMeetingRoomFloorOutletTrigger;
      flags["NeedsMeetingRoomFloorOutlet"] = scope.HasMeetingRoomFloorOutletTrigger;
      flags["NeedsHandDryerCoordination"] = scope.HasHandDryers;
      flags["NeedsCommercialKitchenCoordination"] = scope.HasCommercialKitchen;
      flags["NeedsFoodDisposerCoordination"] = scope.HasFoodDisposer;
      flags["NeedsGreaseTrapCoordination"] = scope.HasGreaseTrap;
      flags["NeedsElectricWaterHeater"] = scope.HasElectricWaterHeater;

      // Mechanical and plumbing coordination.
      flags["NeedsRooftopCoordination"] = scope.HasRooftopUnits;
      flags["NeedsRoofMotorDisconnect"] = scope.HasRoofExhaustFans;
      flags["NeedsDuctSmokeDetectors"] = scope.HasHvacOver2000Cfm;
      flags["HasDuctSmokeAlternative"] = scope.HasDuctSmokeAlternative;
      flags["NeedsMechanicalScheduleReview"] = scope.HasRooftopUnits || scope.HasHvacOver2000Cfm;

      // California lighting and power controls.
      flags["NeedsDaylightHarvesting"] = isCalifornia && fullLightingControls &&
        scope.HasDaylightWindowsOrSkylights && !scope.HasDaylightControlException;
      flags["DaylightControlsExempt"] = isCalifornia && scope.HasDaylightWindowsOrSkylights &&
        scope.HasDaylightControlException;
      flags["NeedsDemandResponse"] = isCalifornia && fullLightingControls && scope.HasLightingOver4000W &&
        !scope.HasDemandResponseHealthLifeSafetyException;
      flags["DemandResponseExempt"] = isCalifornia && scope.HasLightingOver4000W &&
        scope.HasDemandResponseHealthLifeSafetyException;
      flags["NeedsOccupancyControls"] = isCalifornia && hasRegulatedLightingWork && scope.HasOccupancyControlSpaces;
      flags["NeedsMultiLevelLighting"] = isCalifornia && fullLightingControls && scope.HasMultiLevelLightingSpaces;
      flags["NeedsAutomaticShutoffControls"] = isCalifornia && hasRegulatedLightingWork;
      flags["NeedsExteriorLightingControls"] = isCalifornia && scope.HasExteriorLighting &&
        !scope.HasOnlyExteriorLightingControlExceptions;
      flags["ExteriorLightingControlsExempt"] = isCalifornia && scope.HasExteriorLighting &&
        scope.HasOnlyExteriorLightingControlExceptions;
      flags["NeedsTitle24AcceptanceTesting"] = isCalifornia && (
        flags["NeedsDaylightHarvesting"] || flags["NeedsDemandResponse"] ||
        flags["NeedsOccupancyControls"] || flags["NeedsMultiLevelLighting"] ||
        flags["NeedsAutomaticShutoffControls"] || flags["NeedsExteriorLightingControls"] ||
        flags["NeedsControlledReceptacles"]
      );
      flags["NeedsMultiStallRestroomSensors"] = isCalifornia && hasRegulatedLightingWork &&
        scope.HasMultiStallRestrooms;

      // Renewable energy and EV scope are intentionally separate; Article 690, 705, and 706 are not interchangeable.
      flags["NeedsEvChargingCalculations"] = scope.HasEvCharging;
      flags["NeedsEvInfrastructureReview"] = isCalifornia && scope.HasParkingEvInfrastructureTrigger;
      flags["NeedsSolarPvInterconnection"] = scope.HasSolarPv;
      flags["NeedsBessReview"] = scope.HasBessStorage;

      // Emergency, standby, and special systems.
      bool emergencyArchitectureSelected = !string.Equals(
        scope.EmergencyPowerType,
        "None",
        StringComparison.OrdinalIgnoreCase
      );
      bool needsArticle700 = scope.HasArticle700EmergencyLoads && emergencyArchitectureSelected;
      flags["NeedsEmergencyPowerArchitectureResolution"] = (
        scope.HasRequiredExitSigns || scope.HasEmergencyEgressLighting || scope.HasArticle700EmergencyLoads
      ) && !emergencyArchitectureSelected;
      flags["NeedsArticle700Review"] = needsArticle700;
      flags["NeedsArticle701Review"] = scope.HasArticle701LegallyRequiredLoads;
      flags["NeedsArticle702Review"] = scope.HasArticle702OptionalStandbyLoads;
      flags["NeedsSelectiveCoordination"] = needsArticle700 || scope.HasArticle701LegallyRequiredLoads;
      flags["NeedsEmergencySystemSpd"] = needsArticle700 && scope.HasArticle700SwitchboardOrPanelboard;
      flags["NeedsEmergencyInverter"] = needsArticle700 && string.Equals(
        scope.EmergencyPowerType,
        "Central Battery Inverter / Micro-Inverter",
        StringComparison.OrdinalIgnoreCase
      );
      flags["NeedsIntegralBatteryUnits"] = needsArticle700 && string.Equals(
        scope.EmergencyPowerType,
        "Integral Emergency Battery Packs",
        StringComparison.OrdinalIgnoreCase
      );
      flags["NeedsGeneratorReview"] = string.Equals(
        scope.EmergencyPowerType,
        "Emergency Generator",
        StringComparison.OrdinalIgnoreCase
      ) || scope.HasArticle701LegallyRequiredLoads || scope.HasArticle702OptionalStandbyLoads;
      flags["NeedsExitSignReview"] = scope.HasRequiredExitSigns;
      flags["NeedsEgressIllumination"] = scope.HasEmergencyEgressLighting;
      flags["NeedsFireAlarmReview"] = scope.HasFireAlarm;
      flags["NeedsFirePumpReview"] = scope.HasFirePump;
      flags["NeedsElevatorReview"] = scope.HasElevatorsOrEscalators;
      flags["NeedsHealthCareArticle517"] = occupancyHealthCare;
      flags["NeedsHazardousLocationReview"] = occupancyHazardous;
      flags["NeedsPoolSpaReview"] = scope.HasPoolsSpasOrFountains;
      flags["NeedsAfciReview"] = occupancyDwelling || occupancyHotelDorm;

      return flags;
    }

    public static List<InferredCodeMandate> GetInferredMandatesSummary(ProjectScopeConfig scope)
    {
      var list = new List<InferredCodeMandate>();
      if (scope == null) return list;

      var flags = DeriveInferredFlags(scope);

      if (flags["NeedsCodeBasisConfirmation"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "AHJ / Permit-Date Confirmation Required",
          CodeCitation = "CEC/NEC 90.4; adopted-code ordinance",
          TriggerReason = "The governing code edition and local amendment status have not yet been confirmed with the authority having jurisdiction.",
          IsActive = true
        });
      }

      if (flags["NeedsLocalAmendmentReview"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "Local Amendments Apply",
          CodeCitation = "AHJ adopting ordinance",
          TriggerReason = "The scope identifies local amendments; add the ordinance and amendment-specific notes to the review basis.",
          IsActive = true
        });
      }

      // Controlled Receptacles
      if (flags["NeedsControlledReceptacles"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "Controlled Receptacles Required",
          CodeCitation = "Title 24 §130.5(d)",
          TriggerReason = "California scope includes a covered office/lobby/conference/copy/kitchen or hotel space and new construction/addition or an entirely new/fully replaced distribution system.",
          IsActive = true
        });
      }
      else if (flags["ControlledReceptaclesExempt"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "Controlled Receptacle Trigger Not Met",
          CodeCitation = "Title 24 §§130.5(d), 141.0(b)",
          TriggerReason = scope.HasControlledReceptacleException
            ? "A project-specific §130.5(d) receptacle exception was selected and must be documented in the checklist."
            : "The work is not a new building/addition or an entirely new/complete replacement of the building electrical distribution system.",
          IsActive = false
        });
      }

      // Tamper-Resistant Receptacles
      if (flags["NeedsTamperResistant"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "Tamper-Resistant (TR) Receptacles Required",
          CodeCitation = "CEC 406.12",
          TriggerReason = "A listed 406.12 occupancy is present (dwelling, guest/dormitory, child care/education, specified healthcare, or specified assembly area).",
          IsActive = true
        });
      }

      // Daylight Harvesting
      if (flags["NeedsDaylightHarvesting"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "Daylight-Responsive Controls Required",
          CodeCitation = "Title 24 §130.1(d)",
          TriggerReason = "A regulated daylit zone has at least 120 W of general lighting and no documented shading, glazing-area, retail, or parking exception was selected.",
          IsActive = true
        });
      }

      // Demand Response
      if (flags["NeedsDemandResponse"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "Automated Demand Response (ADR) Controls Required",
          CodeCitation = "Title 24 §110.12(c)",
          TriggerReason = "The California project has more than 4,000 W of non-exempt general lighting in a scope requiring full lighting controls.",
          IsActive = true
        });
      }

      // Meeting room receptacle outlets
      if (flags["NeedsMeetingRoomReceptacles"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "Meeting Room Receptacle Layout Required",
          CodeCitation = "CEC 210.65",
          TriggerReason = "Triggered because at least one meeting room is not more than 1,000 sq ft; floor outlets are a separate sub-trigger for rooms at least 215 sq ft with a 12-ft dimension.",
          IsActive = true
        });
      }

      // Rooftop Units
      if (flags["NeedsRooftopCoordination"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "Rooftop HVAC/R Disconnect & 125V GFCI Service Receptacle Required",
          CodeCitation = "CEC 210.63(A) & 440.14",
          TriggerReason = "Triggered because rooftop heating, air-conditioning, or refrigeration equipment is present; general exhaust fans use separate motor-disconnect rules.",
          IsActive = true
        });
      }

      // Duct Smoke
      if (flags["NeedsDuctSmokeDetectors"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = scope.HasDuctSmokeAlternative
            ? "Duct-Smoke Alternative Requires Documentation"
            : "Duct Smoke Detector / HVAC Shutdown Review Required",
          CodeCitation = "CMC 609.1 / CBC 907.3.1",
          TriggerReason = scope.HasDuctSmokeAlternative
            ? "Air-moving capacity exceeds 2,000 CFM, but an area-detection or engineered exception is claimed and must be documented with the fire/mechanical AHJ."
            : "Triggered because an air-moving system supplies more than 2,000 CFM; verify detector location, automatic shutdown, supervision, and exceptions.",
          IsActive = true
        });
      }

      if (flags["NeedsDwellingServiceSpd"] || flags["NeedsEmergencySystemSpd"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "Surge Protection Required",
          CodeCitation = flags["NeedsEmergencySystemSpd"] ? "CEC/NEC 700.8" : "CEC/NEC 230.67",
          TriggerReason = flags["NeedsEmergencySystemSpd"]
            ? "Article 700 emergency switchboards and panelboards require surge protection."
            : "The service supplies an occupancy covered by the selected edition of 230.67 and the service is new or replaced.",
          IsActive = true
        });
      }

      if (flags["NeedsEvInfrastructureReview"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "CALGreen EV Infrastructure Review",
          CodeCitation = "CALGreen §§5.106.5.3–5.106.5.4",
          TriggerReason = "The California scope affects parking/new construction or includes a qualifying service-capacity increase; EV-capable/EVCS quantities and feasibility exceptions must be checked even if no EVSE is shown.",
          IsActive = true
        });
      }

      if (flags["NeedsBessReview"])
      {
        list.Add(new InferredCodeMandate
        {
          MandateName = "Energy Storage System Review",
          CodeCitation = "CEC/NEC Article 706; CFC/IFC §1207",
          TriggerReason = "A BESS is present; Article 706 electrical requirements and fire-code capacity, location, detection, separation, and hazard-mitigation thresholds apply independently of PV rules.",
          IsActive = true
        });
      }

      return list;
    }

    public static bool IsRuleApplicable(
      AuditRuleDefinition rule,
      ProjectScopeConfig scope,
      Dictionary<string, bool> derivedFlags
    )
    {
      if (rule == null || scope == null)
      {
        return false;
      }

      // Discipline check
      if (!string.IsNullOrWhiteSpace(scope.Discipline) &&
          !string.Equals(scope.Discipline, "Full MEP", StringComparison.OrdinalIgnoreCase))
      {
        if (!string.IsNullOrWhiteSpace(rule.Discipline) &&
            !string.Equals(rule.Discipline, "General", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(rule.Discipline, scope.Discipline, StringComparison.OrdinalIgnoreCase))
        {
          return false;
        }
      }

      var cond = rule.Condition;
      if (cond == null)
      {
        return true;
      }

      // Allowed Disciplines
      if (cond.AllowedDisciplines != null && cond.AllowedDisciplines.Count > 0)
      {
        if (!cond.AllowedDisciplines.Any(d => string.Equals(d, scope.Discipline, StringComparison.OrdinalIgnoreCase) ||
                                              string.Equals(scope.Discipline, "Full MEP", StringComparison.OrdinalIgnoreCase)))
        {
          return false;
        }
      }

      // Allowed Phases
      if (cond.AllowedPhases != null && cond.AllowedPhases.Count > 0)
      {
        if (!cond.AllowedPhases.Any(p => string.Equals(p, scope.Phase, StringComparison.OrdinalIgnoreCase)))
        {
          return false;
        }
      }

      // Allowed Clients
      if (cond.AllowedClients != null && cond.AllowedClients.Count > 0)
      {
        if (!cond.AllowedClients.Any(c => string.Equals(c, scope.ClientStandard, StringComparison.OrdinalIgnoreCase)))
        {
          return false;
        }
      }

      // Allowed Jurisdictions
      if (cond.AllowedJurisdictions != null && cond.AllowedJurisdictions.Count > 0)
      {
        if (!cond.AllowedJurisdictions.Any(j => string.Equals(j, scope.CodeJurisdiction, StringComparison.OrdinalIgnoreCase)))
        {
          return false;
        }
      }

      // Emergency Power Types
      if (cond.EmergencyPowerTypes != null && cond.EmergencyPowerTypes.Count > 0)
      {
        if (!cond.EmergencyPowerTypes.Any(e => string.Equals(e, scope.EmergencyPowerType, StringComparison.OrdinalIgnoreCase)))
        {
          return false;
        }
      }

      // Required Flags (All must be true)
      if (cond.RequiredFlags != null && cond.RequiredFlags.Count > 0)
      {
        foreach (var req in cond.RequiredFlags)
        {
          if (!derivedFlags.TryGetValue(req, out bool val) || !val)
          {
            return false;
          }
        }
      }

      // Excluded Flags (None must be true)
      if (cond.ExcludedFlags != null && cond.ExcludedFlags.Count > 0)
      {
        foreach (var exc in cond.ExcludedFlags)
        {
          if (derivedFlags.TryGetValue(exc, out bool val) && val)
          {
            return false;
          }
        }
      }

      // Any Flags (At least one must be true if defined)
      if (cond.AnyFlags != null && cond.AnyFlags.Count > 0)
      {
        bool anyMatched = false;
        foreach (var any in cond.AnyFlags)
        {
          if (derivedFlags.TryGetValue(any, out bool val) && val)
          {
            anyMatched = true;
            break;
          }
        }
        if (!anyMatched)
        {
          return false;
        }
      }

      return true;
    }

    public static AuditMetrics CalculateMetrics(
      IEnumerable<AuditRuleDefinition> activeRules,
      ProjectAuditState state
    )
    {
      var metrics = new AuditMetrics();
      if (activeRules == null) return metrics;

      var checks = state?.Checks ?? new Dictionary<string, AuditCheckItemState>(StringComparer.OrdinalIgnoreCase);

      foreach (var rule in activeRules)
      {
        metrics.TotalApplicable++;
        bool isDone = checks.TryGetValue(rule.Id, out var item) && item != null && item.IsCompleted;
        if (isDone)
        {
          metrics.TotalCompleted++;
        }

        string sev = rule.Severity ?? "High";
        if (string.Equals(sev, "Critical", StringComparison.OrdinalIgnoreCase))
          metrics.CriticalTotal++;
        else if (string.Equals(sev, "High", StringComparison.OrdinalIgnoreCase))
          metrics.HighTotal++;
        else
          metrics.StandardTotal++;

        if (isDone)
        {
          if (string.Equals(sev, "Critical", StringComparison.OrdinalIgnoreCase))
            metrics.CriticalCompleted++;
          else if (string.Equals(sev, "High", StringComparison.OrdinalIgnoreCase))
            metrics.HighCompleted++;
          else
            metrics.StandardCompleted++;
        }
      }

      metrics.OverallPercent = metrics.TotalApplicable > 0
        ? Math.Round((double)metrics.TotalCompleted / metrics.TotalApplicable * 100.0, 1)
        : 100.0;

      metrics.CriticalPercent = metrics.CriticalTotal > 0
        ? Math.Round((double)metrics.CriticalCompleted / metrics.CriticalTotal * 100.0, 1)
        : 100.0;

      return metrics;
    }

    public static ProjectAuditState LoadAuditState(string folderPath)
    {
      string filePath = ResolveAuditStatePath(folderPath);
      if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
      {
        return new ProjectAuditState();
      }

      try
      {
        string json = File.ReadAllText(filePath, Encoding.UTF8);
        var state = JsonConvert.DeserializeObject<ProjectAuditState>(json) ?? new ProjectAuditState();
        if (state.Scope == null) state.Scope = new ProjectScopeConfig();
        if (state.Checks == null) state.Checks = new Dictionary<string, AuditCheckItemState>(StringComparer.OrdinalIgnoreCase);
        return state;
      }
      catch
      {
        return new ProjectAuditState();
      }
    }

    public static void SaveAuditState(string folderPath, ProjectAuditState state)
    {
      string filePath = ResolveAuditStatePath(folderPath);
      if (string.IsNullOrWhiteSpace(filePath))
      {
        return;
      }

      try
      {
        if (!Directory.Exists(folderPath))
        {
          Directory.CreateDirectory(folderPath);
        }

        state.Version = Math.Max(1, state.Version + 1);
        state.LastModifiedUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);

        string json = JsonConvert.SerializeObject(state, Formatting.Indented);
        string tempPath = filePath + ".tmp";
        File.WriteAllText(tempPath, json, Encoding.UTF8);

        if (File.Exists(filePath))
        {
          File.Delete(filePath);
        }
        File.Move(tempPath, filePath);
      }
      catch
      {
        try
        {
          string json = JsonConvert.SerializeObject(state, Formatting.Indented);
          File.WriteAllText(filePath, json, Encoding.UTF8);
        }
        catch
        {
        }
      }
    }

    public static string ResolveAuditStatePath(string folderPath)
    {
      if (string.IsNullOrWhiteSpace(folderPath))
      {
        return string.Empty;
      }
      return Path.Combine(folderPath, AuditStateFileName);
    }
  }
}
