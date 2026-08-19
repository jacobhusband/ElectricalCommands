using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace ElectricalCommands
{
  public sealed class ProjectScopeConfig
  {
    [JsonProperty("phase")]
    public string Phase { get; set; } = "IFP / Permit";

    [JsonProperty("discipline")]
    public string Discipline { get; set; } = "Electrical";

    [JsonProperty("codeJurisdiction")]
    public string CodeJurisdiction { get; set; } = "California (CEC 2022 / T24 2022)";

    [JsonProperty("clientStandard")]
    public string ClientStandard { get; set; } = "Generic Commercial";

    [JsonProperty("codeBasisStatus")]
    public string CodeBasisStatus { get; set; } = "Not Yet Confirmed with AHJ";

    [JsonProperty("authorityHavingJurisdiction")]
    public string AuthorityHavingJurisdiction { get; set; } = string.Empty;

    [JsonProperty("permitApplicationDate")]
    public string PermitApplicationDate { get; set; } = string.Empty;

    [JsonProperty("localAmendmentReference")]
    public string LocalAmendmentReference { get; set; } = string.Empty;

    [JsonProperty("buildingOccupancy")]
    public string BuildingOccupancy { get; set; } = "Business / Mercantile (B / M)";

    [JsonProperty("constructionScopeNature")]
    public string ConstructionScopeNature { get; set; } = "New Construction / 1st Time Tenant Space";

    [JsonProperty("serviceVoltage")]
    public string ServiceVoltage { get; set; } = "120/208V 3PH 4W";

    [JsonProperty("faultCurrentLevel")]
    public string FaultCurrentLevel { get; set; } = "<= 10kAIC (Standard)";

    [JsonProperty("electricalServiceScope")]
    public string ElectricalServiceScope { get; set; } = "Existing Service - No Service Work";

    [JsonProperty("electricalDistributionScope")]
    public string ElectricalDistributionScope { get; set; } = "Entirely New / Complete Replacement";

    [JsonProperty("lightingComplianceScope")]
    public string LightingComplianceScope { get; set; } = "New / Complete Lighting System - Full Controls";

    [JsonProperty("hasLocalAmendments")]
    public bool HasLocalAmendments { get; set; } = false;

    // Factual Interior Spaces & Uses
    [JsonProperty("hasOfficeSpaces")]
    public bool HasOfficeSpaces { get; set; } = true;

    [JsonProperty("hasControlledReceptacleException")]
    public bool HasControlledReceptacleException { get; set; } = false;

    [JsonProperty("hasPublicOrCustomerAreas")]
    public bool HasPublicOrCustomerAreas { get; set; } = true;

    [JsonProperty("hasDwellingUnits")]
    public bool HasDwellingUnits { get; set; } = false;

    [JsonProperty("hasHotelDormitoryOrAssistedLiving")]
    public bool HasHotelDormitoryOrAssistedLiving { get; set; } = false;

    [JsonProperty("hasHotelMotelGuestRooms")]
    public bool HasHotelMotelGuestRooms { get; set; } = false;

    [JsonProperty("has2023NecSpdSleepingOccupancy")]
    public bool Has2023NecSpdSleepingOccupancy { get; set; } = false;

    [JsonProperty("hasChildCareOrEducationAreas")]
    public bool HasChildCareOrEducationAreas { get; set; } = false;

    [JsonProperty("hasHealthCareAreas")]
    public bool HasHealthCareAreas { get; set; } = false;

    [JsonProperty("hasAssemblyAreas")]
    public bool HasAssemblyAreas { get; set; } = false;

    [JsonProperty("hasHazardousClassifiedLocations")]
    public bool HasHazardousClassifiedLocations { get; set; } = false;

    [JsonProperty("hasModularFurniture")]
    public bool HasModularFurniture { get; set; } = true;

    [JsonProperty("hasMeetingRoomsOver1000SqFt")]
    public bool HasMeetingRoomsOver1000SqFt { get; set; } = false;

    // Retained separately from the legacy, incorrectly oriented ">= 1,000 sq ft" answer above.
    // NEC/CEC 210.65 applies to meeting rooms not more than 1,000 sq ft.
    [JsonProperty("hasMeetingRoomsAtOrBelow1000SqFt")]
    public bool HasMeetingRoomsAtOrBelow1000SqFt { get; set; } = false;

    [JsonProperty("hasMeetingRoomFloorOutletTrigger")]
    public bool HasMeetingRoomFloorOutletTrigger { get; set; } = false;

    [JsonProperty("hasMultiStallRestrooms")]
    public bool HasMultiStallRestrooms { get; set; } = true;

    [JsonProperty("hasCommercialKitchen")]
    public bool HasCommercialKitchen { get; set; } = false;

    [JsonProperty("hasFoodDisposer")]
    public bool HasFoodDisposer { get; set; } = false;

    [JsonProperty("hasGreaseTrap")]
    public bool HasGreaseTrap { get; set; } = false;

    [JsonProperty("hasSinksOrWetLocations")]
    public bool HasSinksOrWetLocations { get; set; } = true;

    [JsonProperty("hasGaragesServiceBaysOrIndoorParking")]
    public bool HasGaragesServiceBaysOrIndoorParking { get; set; } = false;

    [JsonProperty("hasHandDryers")]
    public bool HasHandDryers { get; set; } = false;

    // Factual Mechanical & Plumbing Coordination
    [JsonProperty("hasRooftopUnits")]
    public bool HasRooftopUnits { get; set; } = false;

    [JsonProperty("hasHvacOver2000Cfm")]
    public bool HasHvacOver2000Cfm { get; set; } = false;

    [JsonProperty("hasDuctSmokeAlternative")]
    public bool HasDuctSmokeAlternative { get; set; } = false;

    [JsonProperty("hasRoofExhaustFans")]
    public bool HasRoofExhaustFans { get; set; } = false;

    [JsonProperty("hasElectricWaterHeater")]
    public bool HasElectricWaterHeater { get; set; } = false;

    [JsonProperty("hasNewTransformers")]
    public bool HasNewTransformers { get; set; } = false;

    [JsonProperty("hasService1200AmpsOrMore")]
    public bool HasService1200AmpsOrMore { get; set; } = false;

    [JsonProperty("hasOcpd1200AmpsOrMore")]
    public bool HasOcpd1200AmpsOrMore { get; set; } = false;

    // Factual Lighting, Daylighting & Emergency
    [JsonProperty("hasDaylightWindowsOrSkylights")]
    public bool HasDaylightWindowsOrSkylights { get; set; } = true;

    [JsonProperty("hasDaylightControlException")]
    public bool HasDaylightControlException { get; set; } = false;

    [JsonProperty("hasLightingOver4000W")]
    public bool HasLightingOver4000W { get; set; } = false;

    [JsonProperty("hasDemandResponseHealthLifeSafetyException")]
    public bool HasDemandResponseHealthLifeSafetyException { get; set; } = false;

    [JsonProperty("hasMultiLevelLightingSpaces")]
    public bool HasMultiLevelLightingSpaces { get; set; } = true;

    [JsonProperty("hasOccupancyControlSpaces")]
    public bool HasOccupancyControlSpaces { get; set; } = true;

    [JsonProperty("hasExteriorLighting")]
    public bool HasExteriorLighting { get; set; } = false;

    [JsonProperty("hasOnlyExteriorLightingControlExceptions")]
    public bool HasOnlyExteriorLightingControlExceptions { get; set; } = false;

    [JsonProperty("emergencyPowerType")]
    public string EmergencyPowerType { get; set; } = "Central Battery Inverter / Micro-Inverter";

    // Factual Renewable & EV
    [JsonProperty("hasEvCharging")]
    public bool HasEvCharging { get; set; } = false;

    [JsonProperty("hasParkingEvInfrastructureTrigger")]
    public bool HasParkingEvInfrastructureTrigger { get; set; } = false;

    [JsonProperty("hasSolarPv")]
    public bool HasSolarPv { get; set; } = false;

    [JsonProperty("hasBessStorage")]
    public bool HasBessStorage { get; set; } = false;

    // Life-safety and special systems
    [JsonProperty("hasRequiredExitSignsOrEmergencyLighting")]
    public bool HasRequiredExitSignsOrEmergencyLighting { get; set; } = true;

    [JsonProperty("hasRequiredExitSigns")]
    public bool HasRequiredExitSigns { get; set; } = true;

    [JsonProperty("hasEmergencyEgressLighting")]
    public bool HasEmergencyEgressLighting { get; set; } = true;

    [JsonProperty("hasArticle700EmergencyLoads")]
    public bool HasArticle700EmergencyLoads { get; set; } = true;

    [JsonProperty("hasArticle700SwitchboardOrPanelboard")]
    public bool HasArticle700SwitchboardOrPanelboard { get; set; } = true;

    [JsonProperty("hasArticle701LegallyRequiredLoads")]
    public bool HasArticle701LegallyRequiredLoads { get; set; } = false;

    [JsonProperty("hasArticle702OptionalStandbyLoads")]
    public bool HasArticle702OptionalStandbyLoads { get; set; } = false;

    [JsonProperty("hasFireAlarm")]
    public bool HasFireAlarm { get; set; } = false;

    [JsonProperty("hasFirePump")]
    public bool HasFirePump { get; set; } = false;

    [JsonProperty("hasElevatorsOrEscalators")]
    public bool HasElevatorsOrEscalators { get; set; } = false;

    [JsonProperty("hasPoolsSpasOrFountains")]
    public bool HasPoolsSpasOrFountains { get; set; } = false;

    public ProjectScopeConfig Clone()
    {
      string json = JsonConvert.SerializeObject(this);
      return JsonConvert.DeserializeObject<ProjectScopeConfig>(json) ?? new ProjectScopeConfig();
    }
  }

  public sealed class InferredCodeMandate
  {
    public string MandateName { get; set; }
    public string CodeCitation { get; set; }
    public string TriggerReason { get; set; }
    public bool IsActive { get; set; }
  }

  public sealed class AuditRuleDefinition
  {
    [JsonProperty("id")]
    public string Id { get; set; }

    [JsonProperty("title")]
    public string Title { get; set; }

    [JsonProperty("description")]
    public string Description { get; set; }

    [JsonProperty("codeCitation")]
    public string CodeCitation { get; set; }

    [JsonProperty("discipline")]
    public string Discipline { get; set; } = "Electrical";

    [JsonProperty("category")]
    public string Category { get; set; } = "Code Compliance";

    [JsonProperty("topic")]
    public string Topic { get; set; } = "General Setup";

    [JsonProperty("severity")]
    public string Severity { get; set; } = "High"; // Critical, High, Standard, Aesthetic

    [JsonProperty("whyItMatters")]
    public string WhyItMatters { get; set; }

    [JsonProperty("exceptionsAndNuances")]
    public string ExceptionsAndNuances { get; set; }

    [JsonProperty("codeCycleNotes")]
    public string CodeCycleNotes { get; set; }

    [JsonProperty("order")]
    public int Order { get; set; }

    [JsonProperty("condition")]
    public RuleCondition Condition { get; set; } = new RuleCondition();
  }

  public sealed class RuleCondition
  {
    [JsonProperty("requiredFlags")]
    public List<string> RequiredFlags { get; set; } = new List<string>();

    [JsonProperty("anyFlags")]
    public List<string> AnyFlags { get; set; } = new List<string>();

    [JsonProperty("excludedFlags")]
    public List<string> ExcludedFlags { get; set; } = new List<string>();

    [JsonProperty("allowedDisciplines")]
    public List<string> AllowedDisciplines { get; set; } = new List<string>();

    [JsonProperty("allowedPhases")]
    public List<string> AllowedPhases { get; set; } = new List<string>();

    [JsonProperty("allowedJurisdictions")]
    public List<string> AllowedJurisdictions { get; set; } = new List<string>();

    [JsonProperty("allowedClients")]
    public List<string> AllowedClients { get; set; } = new List<string>();

    [JsonProperty("emergencyPowerTypes")]
    public List<string> EmergencyPowerTypes { get; set; } = new List<string>();
  }

  public sealed class MasterAuditCatalogFile
  {
    [JsonProperty("version")]
    public string Version { get; set; } = "3.0.0";

    [JsonProperty("rules")]
    public List<AuditRuleDefinition> Rules { get; set; } = new List<AuditRuleDefinition>();
  }

  public sealed class ProjectAuditState
  {
    [JsonProperty("version")]
    public long Version { get; set; } = 1;

    [JsonProperty("lastModifiedUtc")]
    public string LastModifiedUtc { get; set; } = DateTime.UtcNow.ToString("o");

    [JsonProperty("scope")]
    public ProjectScopeConfig Scope { get; set; } = new ProjectScopeConfig();

    [JsonProperty("checks")]
    public Dictionary<string, AuditCheckItemState> Checks { get; set; } =
      new Dictionary<string, AuditCheckItemState>(StringComparer.OrdinalIgnoreCase);
  }

  public sealed class AuditCheckItemState
  {
    [JsonProperty("isCompleted")]
    public bool IsCompleted { get; set; }

    [JsonProperty("completedAtUtc")]
    public string CompletedAtUtc { get; set; }

    [JsonProperty("notes")]
    public string Notes { get; set; } = string.Empty;

    [JsonProperty("status")]
    public string Status { get; set; } = "Pending"; // Pass, Fail, NA, Pending
  }

  public sealed class AuditMetrics
  {
    public int TotalApplicable { get; set; }
    public int TotalCompleted { get; set; }
    public double OverallPercent { get; set; }

    public int CriticalTotal { get; set; }
    public int CriticalCompleted { get; set; }
    public double CriticalPercent { get; set; }

    public int HighTotal { get; set; }
    public int HighCompleted { get; set; }

    public int StandardTotal { get; set; }
    public int StandardCompleted { get; set; }

    public bool IsPermitReady => CriticalTotal == 0 || CriticalCompleted == CriticalTotal;
  }
}
