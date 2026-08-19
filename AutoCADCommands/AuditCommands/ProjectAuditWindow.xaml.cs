using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace ElectricalCommands
{
  public partial class ProjectAuditWindow : Window
  {
    private string _folderPath;
    private string _dwgPath;
    private List<AuditRuleDefinition> _masterCatalog;
    private ProjectAuditState _state;
    private List<AuditRuleDefinition> _filteredRules;
    private bool _isUpdatingUi;

    internal ProjectAuditWindow(
      string folderPath,
      string dwgPath,
      List<AuditRuleDefinition> masterCatalog,
      ProjectAuditState state
    )
    {
      InitializeComponent();

      _folderPath = folderPath ?? string.Empty;
      _dwgPath = dwgPath ?? string.Empty;
      _masterCatalog = masterCatalog ?? new List<AuditRuleDefinition>();
      _state = state ?? new ProjectAuditState();
      if (_state.Scope == null) _state.Scope = new ProjectScopeConfig();
      if (_state.Checks == null) _state.Checks = new Dictionary<string, AuditCheckItemState>(StringComparer.OrdinalIgnoreCase);

      FolderPathTextBlock.Text = !string.IsNullOrWhiteSpace(_folderPath) ? _folderPath : "Unsaved / Unknown";
      DwgNameTextBlock.Text = !string.IsNullOrWhiteSpace(_dwgPath) ? Path.GetFileName(_dwgPath) : "Active Drawing";

      InitScopeDropdowns();
      LoadScopeFromState();
      ReevaluateRules();
    }

    internal void SwitchFolderOrDrawing(string newFolderPath, string newDwgPath)
    {
      bool folderChanged = !string.Equals(_folderPath, newFolderPath, StringComparison.OrdinalIgnoreCase);
      _folderPath = newFolderPath ?? string.Empty;
      _dwgPath = newDwgPath ?? string.Empty;

      FolderPathTextBlock.Text = !string.IsNullOrWhiteSpace(_folderPath) ? _folderPath : "Unsaved / Unknown";
      DwgNameTextBlock.Text = !string.IsNullOrWhiteSpace(_dwgPath) ? Path.GetFileName(_dwgPath) : "Active Drawing";

      if (folderChanged)
      {
        _state = AuditEngine.LoadAuditState(_folderPath);
        LoadScopeFromState();
      }

      ReevaluateRules();
      SetStatus($"Active drawing: {DwgNameTextBlock.Text}", isError: false);
    }

    private void InitScopeDropdowns()
    {
      _isUpdatingUi = true;
      try
      {
        PhaseComboBox.ItemsSource = new[] { "DD", "CD50", "CD90", "IFP / Permit", "Bid", "CA / As-Built" };
        DisciplineComboBox.ItemsSource = new[] { "Electrical", "Full MEP", "Mechanical", "Plumbing" };
        JurisdictionComboBox.ItemsSource = new[] {
          "California (CEC 2022 / T24 2022)",
          "California (CEC 2025 / T24 2025)",
          "National (NEC 2020 / IECC)",
          "National (NEC 2023 / IECC)",
          "Washington State",
          "Texas / Other"
        };
        ClientStandardComboBox.ItemsSource = new[] {
          "Generic Commercial",
          "Bank of America (Prototype)",
          "Gensler Commercial",
          "JPMorgan Chase",
          "Retail Standard"
        };
        BuildingOccupancyComboBox.ItemsSource = new[] {
          "Business / Mercantile (B / M)",
          "Assembly (A)",
          "Education / Child Care (E / I-4)",
          "Factory / Storage (F / S)",
          "Health Care / Institutional (I)",
          "Hotel / Dormitory / Assisted Living (R-1 / R-2)",
          "Dwelling Units / Multifamily Residential (R-2 / R-3)",
          "Hazardous (H)",
          "Mixed / Other - Verify with AHJ"
        };
        CodeBasisStatusComboBox.ItemsSource = new[] {
          "Not Yet Confirmed with AHJ",
          "Confirmed - No Local Amendments",
          "Confirmed - Local Amendments Apply"
        };
        ConstructionScopeNatureComboBox.ItemsSource = new[] {
          "New Building / Addition",
          "First-Time Tenant Build-Out / Shell Completion",
          "New Construction / 1st Time Tenant Space",
          "Existing Space - Remodel with Completely NEW Electrical Distribution",
          "Existing Space - Remodel with Existing Electrical Distribution To Remain"
        };
        ServiceVoltageComboBox.ItemsSource = new[] {
          "120/208V 3PH 4W",
          "277/480V 3PH 4W",
          "120/240V 1PH 3W",
          "120/240V 3PH 4W High-Leg Delta"
        };
        FaultCurrentComboBox.ItemsSource = new[] {
          "<= 10kAIC (Standard)",
          "> 10kAIC / High Fault Current (22k-65kAIC)"
        };
        EmergencyPowerComboBox.ItemsSource = new[] {
          "Central Battery Inverter / Micro-Inverter",
          "Integral Emergency Battery Packs",
          "Emergency Generator",
          "None"
        };
        ElectricalServiceScopeComboBox.ItemsSource = new[] {
          "New Service / Service Replacement",
          "Service Upgrade / Modification",
          "Existing Service - No Service Work",
          "No Electrical Service in Scope"
        };
        ElectricalDistributionScopeComboBox.ItemsSource = new[] {
          "Entirely New / Complete Replacement",
          "Partial Alteration / Tenant Distribution",
          "Existing Distribution - No Distribution Work"
        };
        LightingComplianceScopeComboBox.ItemsSource = new[] {
          "New / Complete Lighting System - Full Controls",
          "Alteration - Full Controls per Table 141.0-F",
          "Alteration - Reduced Controls Path",
          "No Regulated Lighting Work",
          "Unknown - Verify Alteration Path"
        };
      }
      finally
      {
        _isUpdatingUi = false;
      }
    }

    private void LoadScopeFromState()
    {
      _isUpdatingUi = true;
      try
      {
        var sc = _state.Scope ?? new ProjectScopeConfig();

        SetComboValue(PhaseComboBox, sc.Phase);
        SetComboValue(DisciplineComboBox, sc.Discipline);
        SetComboValue(JurisdictionComboBox, sc.CodeJurisdiction);
        SetComboValue(ClientStandardComboBox, sc.ClientStandard);
        SetComboValue(BuildingOccupancyComboBox, sc.BuildingOccupancy);
        SetComboValue(CodeBasisStatusComboBox, sc.CodeBasisStatus);
        SetComboValue(ConstructionScopeNatureComboBox, sc.ConstructionScopeNature);
        SetComboValue(ServiceVoltageComboBox, sc.ServiceVoltage);
        SetComboValue(FaultCurrentComboBox, sc.FaultCurrentLevel);
        SetComboValue(EmergencyPowerComboBox, sc.EmergencyPowerType);
        SetComboValue(ElectricalServiceScopeComboBox, sc.ElectricalServiceScope);
        SetComboValue(ElectricalDistributionScopeComboBox, sc.ElectricalDistributionScope);
        SetComboValue(LightingComplianceScopeComboBox, sc.LightingComplianceScope);
        AuthorityHavingJurisdictionTextBox.Text = sc.AuthorityHavingJurisdiction ?? string.Empty;
        PermitApplicationDateTextBox.Text = sc.PermitApplicationDate ?? string.Empty;
        LocalAmendmentReferenceTextBox.Text = sc.LocalAmendmentReference ?? string.Empty;
        LocalAmendmentsCheckBox.IsChecked = sc.HasLocalAmendments;

        // Factual space flags
        OfficeSpacesCheckBox.IsChecked = sc.HasOfficeSpaces;
        ControlledReceptacleExceptionCheckBox.IsChecked = sc.HasControlledReceptacleException;
        PublicAreasCheckBox.IsChecked = sc.HasPublicOrCustomerAreas;
        DwellingUnitsCheckBox.IsChecked = sc.HasDwellingUnits;
        HotelDormitoryCheckBox.IsChecked = sc.HasHotelDormitoryOrAssistedLiving;
        HotelGuestRoomsCheckBox.IsChecked = sc.HasHotelMotelGuestRooms;
        SpdSleepingOccupancyCheckBox.IsChecked = sc.Has2023NecSpdSleepingOccupancy;
        ChildCareEducationCheckBox.IsChecked = sc.HasChildCareOrEducationAreas;
        HealthCareCheckBox.IsChecked = sc.HasHealthCareAreas;
        AssemblyAreasCheckBox.IsChecked = sc.HasAssemblyAreas;
        HazardousLocationsCheckBox.IsChecked = sc.HasHazardousClassifiedLocations;
        ModularFurnitureCheckBox.IsChecked = sc.HasModularFurniture;
        MeetingRoomsAtMost1000CheckBox.IsChecked = sc.HasMeetingRoomsAtOrBelow1000SqFt;
        MeetingRoomFloorOutletTriggerCheckBox.IsChecked = sc.HasMeetingRoomFloorOutletTrigger;
        MultiStallRestroomsCheckBox.IsChecked = sc.HasMultiStallRestrooms;
        SinksWetLocationsCheckBox.IsChecked = sc.HasSinksOrWetLocations;
        GaragesServiceBaysCheckBox.IsChecked = sc.HasGaragesServiceBaysOrIndoorParking;
        HandDryersCheckBox.IsChecked = sc.HasHandDryers;
        CommercialKitchenCheckBox.IsChecked = sc.HasCommercialKitchen;
        FoodDisposerCheckBox.IsChecked = sc.HasFoodDisposer;
        GreaseTrapCheckBox.IsChecked = sc.HasGreaseTrap;

        // Mechanical / Plumbing flags
        RooftopUnitsCheckBox.IsChecked = sc.HasRooftopUnits;
        RoofExhaustFansCheckBox.IsChecked = sc.HasRoofExhaustFans;
        HvacOver2000CfmCheckBox.IsChecked = sc.HasHvacOver2000Cfm;
        DuctSmokeAlternativeCheckBox.IsChecked = sc.HasDuctSmokeAlternative;
        ElectricWaterHeaterCheckBox.IsChecked = sc.HasElectricWaterHeater;

        // Distribution flags
        NewTransformersCheckBox.IsChecked = sc.HasNewTransformers;
        Service1200AmpsCheckBox.IsChecked = sc.HasService1200AmpsOrMore;
        Ocpd1200AmpsCheckBox.IsChecked = sc.HasOcpd1200AmpsOrMore;

        // Lighting flags
        DaylightWindowsCheckBox.IsChecked = sc.HasDaylightWindowsOrSkylights;
        DaylightExceptionCheckBox.IsChecked = sc.HasDaylightControlException;
        LightingOver4000WCheckBox.IsChecked = sc.HasLightingOver4000W;
        DemandResponseExceptionCheckBox.IsChecked = sc.HasDemandResponseHealthLifeSafetyException;
        MultiLevelLightingSpacesCheckBox.IsChecked = sc.HasMultiLevelLightingSpaces;
        OccupancyControlSpacesCheckBox.IsChecked = sc.HasOccupancyControlSpaces;
        ExteriorLightingCheckBox.IsChecked = sc.HasExteriorLighting;
        ExteriorLightingExceptionCheckBox.IsChecked = sc.HasOnlyExteriorLightingControlExceptions;

        // Renewable, EV, and life-safety flags
        EvChargingCheckBox.IsChecked = sc.HasEvCharging;
        ParkingEvInfrastructureCheckBox.IsChecked = sc.HasParkingEvInfrastructureTrigger;
        SolarPvCheckBox.IsChecked = sc.HasSolarPv;
        BessStorageCheckBox.IsChecked = sc.HasBessStorage;
        RequiredExitSignsCheckBox.IsChecked = sc.HasRequiredExitSigns;
        EmergencyEgressLightingCheckBox.IsChecked = sc.HasEmergencyEgressLighting;
        Article700LoadsCheckBox.IsChecked = sc.HasArticle700EmergencyLoads;
        Article700PanelboardCheckBox.IsChecked = sc.HasArticle700SwitchboardOrPanelboard;
        Article701LoadsCheckBox.IsChecked = sc.HasArticle701LegallyRequiredLoads;
        Article702LoadsCheckBox.IsChecked = sc.HasArticle702OptionalStandbyLoads;
        FireAlarmCheckBox.IsChecked = sc.HasFireAlarm;
        FirePumpCheckBox.IsChecked = sc.HasFirePump;
        ElevatorsCheckBox.IsChecked = sc.HasElevatorsOrEscalators;
        PoolsSpasFountainsCheckBox.IsChecked = sc.HasPoolsSpasOrFountains;
      }
      finally
      {
        _isUpdatingUi = false;
      }
    }

    private void SetComboValue(ComboBox combo, string value)
    {
      if (combo == null || string.IsNullOrWhiteSpace(value)) return;
      if (ReferenceEquals(combo, ServiceVoltageComboBox) &&
          string.Equals(value, "120/240V 3PH Delta", StringComparison.OrdinalIgnoreCase))
      {
        value = "120/240V 3PH 4W High-Leg Delta";
      }
      if (ReferenceEquals(combo, BuildingOccupancyComboBox) &&
          string.Equals(value, "Hotel / Residential / Dormitory (R)", StringComparison.OrdinalIgnoreCase))
      {
        value = "Mixed / Other - Verify with AHJ";
      }
      foreach (var item in combo.Items)
      {
        if (string.Equals(item?.ToString(), value, StringComparison.OrdinalIgnoreCase))
        {
          combo.SelectedItem = item;
          return;
        }
      }
      if (combo.Items.Count > 0)
      {
        combo.SelectedIndex = 0;
      }
    }

    private void SyncScopeFromUi()
    {
      if (_state.Scope == null)
      {
        _state.Scope = new ProjectScopeConfig();
      }

      var sc = _state.Scope;

      sc.Phase = PhaseComboBox.SelectedItem?.ToString() ?? "IFP / Permit";
      sc.Discipline = DisciplineComboBox.SelectedItem?.ToString() ?? "Electrical";
      sc.CodeJurisdiction = JurisdictionComboBox.SelectedItem?.ToString() ?? "California (CEC 2022 / T24 2022)";
      sc.ClientStandard = ClientStandardComboBox.SelectedItem?.ToString() ?? "Generic Commercial";
      sc.BuildingOccupancy = BuildingOccupancyComboBox.SelectedItem?.ToString() ?? "Business / Mercantile (B / M)";
      sc.CodeBasisStatus = CodeBasisStatusComboBox.SelectedItem?.ToString() ?? "Not Yet Confirmed with AHJ";
      sc.ConstructionScopeNature = ConstructionScopeNatureComboBox.SelectedItem?.ToString() ?? "New Construction / 1st Time Tenant Space";
      sc.ServiceVoltage = ServiceVoltageComboBox.SelectedItem?.ToString() ?? "120/208V 3PH 4W";
      sc.FaultCurrentLevel = FaultCurrentComboBox.SelectedItem?.ToString() ?? "<= 10kAIC (Standard)";
      sc.EmergencyPowerType = EmergencyPowerComboBox.SelectedItem?.ToString() ?? "Central Battery Inverter / Micro-Inverter";
      sc.ElectricalServiceScope = ElectricalServiceScopeComboBox.SelectedItem?.ToString() ?? "Existing Service - No Service Work";
      sc.ElectricalDistributionScope = ElectricalDistributionScopeComboBox.SelectedItem?.ToString() ?? "Entirely New / Complete Replacement";
      sc.LightingComplianceScope = LightingComplianceScopeComboBox.SelectedItem?.ToString() ?? "New / Complete Lighting System - Full Controls";
      sc.AuthorityHavingJurisdiction = AuthorityHavingJurisdictionTextBox.Text?.Trim() ?? string.Empty;
      sc.PermitApplicationDate = PermitApplicationDateTextBox.Text?.Trim() ?? string.Empty;
      sc.LocalAmendmentReference = LocalAmendmentReferenceTextBox.Text?.Trim() ?? string.Empty;
      sc.HasLocalAmendments = LocalAmendmentsCheckBox.IsChecked == true;

      sc.HasOfficeSpaces = OfficeSpacesCheckBox.IsChecked == true;
      sc.HasControlledReceptacleException = ControlledReceptacleExceptionCheckBox.IsChecked == true;
      sc.HasPublicOrCustomerAreas = PublicAreasCheckBox.IsChecked == true;
      sc.HasDwellingUnits = DwellingUnitsCheckBox.IsChecked == true;
      sc.HasHotelDormitoryOrAssistedLiving = HotelDormitoryCheckBox.IsChecked == true;
      sc.HasHotelMotelGuestRooms = HotelGuestRoomsCheckBox.IsChecked == true;
      sc.Has2023NecSpdSleepingOccupancy = SpdSleepingOccupancyCheckBox.IsChecked == true;
      sc.HasChildCareOrEducationAreas = ChildCareEducationCheckBox.IsChecked == true;
      sc.HasHealthCareAreas = HealthCareCheckBox.IsChecked == true;
      sc.HasAssemblyAreas = AssemblyAreasCheckBox.IsChecked == true;
      sc.HasHazardousClassifiedLocations = HazardousLocationsCheckBox.IsChecked == true;
      sc.HasModularFurniture = ModularFurnitureCheckBox.IsChecked == true;
      sc.HasMeetingRoomsAtOrBelow1000SqFt = MeetingRoomsAtMost1000CheckBox.IsChecked == true;
      sc.HasMeetingRoomFloorOutletTrigger = MeetingRoomFloorOutletTriggerCheckBox.IsChecked == true;
      sc.HasMultiStallRestrooms = MultiStallRestroomsCheckBox.IsChecked == true;
      sc.HasSinksOrWetLocations = SinksWetLocationsCheckBox.IsChecked == true;
      sc.HasGaragesServiceBaysOrIndoorParking = GaragesServiceBaysCheckBox.IsChecked == true;
      sc.HasHandDryers = HandDryersCheckBox.IsChecked == true;
      sc.HasCommercialKitchen = CommercialKitchenCheckBox.IsChecked == true;
      sc.HasFoodDisposer = FoodDisposerCheckBox.IsChecked == true;
      sc.HasGreaseTrap = GreaseTrapCheckBox.IsChecked == true;

      sc.HasRooftopUnits = RooftopUnitsCheckBox.IsChecked == true;
      sc.HasRoofExhaustFans = RoofExhaustFansCheckBox.IsChecked == true;
      sc.HasHvacOver2000Cfm = HvacOver2000CfmCheckBox.IsChecked == true;
      sc.HasDuctSmokeAlternative = DuctSmokeAlternativeCheckBox.IsChecked == true;
      sc.HasElectricWaterHeater = ElectricWaterHeaterCheckBox.IsChecked == true;
      sc.HasNewTransformers = NewTransformersCheckBox.IsChecked == true;
      sc.HasService1200AmpsOrMore = Service1200AmpsCheckBox.IsChecked == true;
      sc.HasOcpd1200AmpsOrMore = Ocpd1200AmpsCheckBox.IsChecked == true;

      sc.HasDaylightWindowsOrSkylights = DaylightWindowsCheckBox.IsChecked == true;
      sc.HasDaylightControlException = DaylightExceptionCheckBox.IsChecked == true;
      sc.HasLightingOver4000W = LightingOver4000WCheckBox.IsChecked == true;
      sc.HasDemandResponseHealthLifeSafetyException = DemandResponseExceptionCheckBox.IsChecked == true;
      sc.HasMultiLevelLightingSpaces = MultiLevelLightingSpacesCheckBox.IsChecked == true;
      sc.HasOccupancyControlSpaces = OccupancyControlSpacesCheckBox.IsChecked == true;
      sc.HasExteriorLighting = ExteriorLightingCheckBox.IsChecked == true;
      sc.HasOnlyExteriorLightingControlExceptions = ExteriorLightingExceptionCheckBox.IsChecked == true;

      sc.HasEvCharging = EvChargingCheckBox.IsChecked == true;
      sc.HasParkingEvInfrastructureTrigger = ParkingEvInfrastructureCheckBox.IsChecked == true;
      sc.HasSolarPv = SolarPvCheckBox.IsChecked == true;
      sc.HasBessStorage = BessStorageCheckBox.IsChecked == true;
      sc.HasRequiredExitSigns = RequiredExitSignsCheckBox.IsChecked == true;
      sc.HasEmergencyEgressLighting = EmergencyEgressLightingCheckBox.IsChecked == true;
      sc.HasRequiredExitSignsOrEmergencyLighting = sc.HasRequiredExitSigns || sc.HasEmergencyEgressLighting;
      sc.HasArticle700EmergencyLoads = Article700LoadsCheckBox.IsChecked == true;
      sc.HasArticle700SwitchboardOrPanelboard = Article700PanelboardCheckBox.IsChecked == true;
      sc.HasArticle701LegallyRequiredLoads = Article701LoadsCheckBox.IsChecked == true;
      sc.HasArticle702OptionalStandbyLoads = Article702LoadsCheckBox.IsChecked == true;
      sc.HasFireAlarm = FireAlarmCheckBox.IsChecked == true;
      sc.HasFirePump = FirePumpCheckBox.IsChecked == true;
      sc.HasElevatorsOrEscalators = ElevatorsCheckBox.IsChecked == true;
      sc.HasPoolsSpasOrFountains = PoolsSpasFountainsCheckBox.IsChecked == true;
    }

    private void OnScopeInputChanged(object sender, RoutedEventArgs e)
    {
      if (_isUpdatingUi) return;
      SyncScopeFromUi();
      ReevaluateRules();
      SaveStateToDisk();
    }

    private void ReevaluateRules()
    {
      _filteredRules = AuditEngine.FilterRules(_masterCatalog, _state.Scope);
      MatchingRuleCountTextBlock.Text = $"{_filteredRules.Count} Checks";
      RenderInferredMandates();
      RenderChecklistItems();
      UpdateMetricsDisplay();
    }

    private void RenderInferredMandates()
    {
      var mandates = AuditEngine.GetInferredMandatesSummary(_state.Scope);
      var viewModels = mandates.Select(m => new InferredMandateViewModel(m)).ToList();
      InferredMandatesItemsControl.ItemsSource = viewModels;
    }

    private void RenderChecklistItems()
    {
      if (_filteredRules == null) return;

      string search = (SearchTextBox?.Text ?? string.Empty).Trim();
      bool hideCompleted = HideCompletedCheckBox?.IsChecked == true;
      string sevFilter = (SeverityFilterComboBox?.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "All Severities";
      string topicFilter = (TopicFilterComboBox?.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "All Topics";

      var displayList = new List<AuditItemViewModel>();

      foreach (var rule in _filteredRules)
      {
        bool isDone = _state.Checks.TryGetValue(rule.Id, out var checkState) && checkState != null && checkState.IsCompleted;
        string notes = checkState?.Notes ?? string.Empty;

        if (hideCompleted && isDone)
        {
          continue;
        }

        // Search Filter (checks title, description, code citation, exceptions, code cycle notes, and reviewer notes)
        if (!string.IsNullOrWhiteSpace(search))
        {
          bool match = (rule.Title?.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (rule.Description?.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (rule.CodeCitation?.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (rule.ExceptionsAndNuances?.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (rule.CodeCycleNotes?.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0) ||
                       (notes?.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0);
          if (!match) continue;
        }

        // Severity Filter
        if (!string.Equals(sevFilter, "All Severities", StringComparison.OrdinalIgnoreCase))
        {
          if (string.Equals(sevFilter, "Critical Only", StringComparison.OrdinalIgnoreCase) &&
              !string.Equals(rule.Severity, "Critical", StringComparison.OrdinalIgnoreCase))
          {
            continue;
          }
          if (string.Equals(sevFilter, "High & Critical", StringComparison.OrdinalIgnoreCase) &&
              !string.Equals(rule.Severity, "Critical", StringComparison.OrdinalIgnoreCase) &&
              !string.Equals(rule.Severity, "High", StringComparison.OrdinalIgnoreCase))
          {
            continue;
          }
          if (string.Equals(sevFilter, "Standard", StringComparison.OrdinalIgnoreCase) &&
              !string.Equals(rule.Severity, "Standard", StringComparison.OrdinalIgnoreCase))
          {
            continue;
          }
        }

        // Topic Filter (flexible topic matching instead of rigid sheet numbers)
        if (!string.Equals(topicFilter, "All Topics", StringComparison.OrdinalIgnoreCase))
        {
          if (!string.Equals(rule.Topic, topicFilter, StringComparison.OrdinalIgnoreCase))
          {
            continue;
          }
        }

        displayList.Add(new AuditItemViewModel(rule, isDone, notes));
      }

      AuditItemsControl.ItemsSource = displayList;
    }

    private void UpdateMetricsDisplay()
    {
      var metrics = AuditEngine.CalculateMetrics(_filteredRules, _state);

      ProgressCountTextBlock.Text = $"{metrics.TotalCompleted} / {metrics.TotalApplicable} Checks";
      ProgressPercentTextBlock.Text = $"{metrics.OverallPercent}%";
      AuditProgressBar.Value = metrics.OverallPercent;

      if (metrics.TotalApplicable == 0)
      {
        ReadinessVerdictTextBlock.Text = "No Applicable Checks";
        ReadinessVerdictTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(88, 96, 105));
        ReadinessBadgeBorder.Background = new SolidColorBrush(Color.FromRgb(246, 248, 250));
        ReadinessBadgeBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(209, 213, 218));
      }
      else if (metrics.CriticalTotal > 0 && metrics.CriticalCompleted < metrics.CriticalTotal)
      {
        int remainingCrit = metrics.CriticalTotal - metrics.CriticalCompleted;
        ReadinessVerdictTextBlock.Text = $"⚠ {remainingCrit} Critical Check{(remainingCrit > 1 ? "s" : "")} Remaining";
        ReadinessVerdictTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(203, 36, 49));
        ReadinessBadgeBorder.Background = new SolidColorBrush(Color.FromRgb(255, 235, 233));
        ReadinessBadgeBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(255, 129, 130));
      }
      else if (metrics.TotalCompleted == metrics.TotalApplicable)
      {
        ReadinessVerdictTextBlock.Text = "✔ 100% Verified - Ready for Permit Issue";
        ReadinessVerdictTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(40, 167, 69));
        ReadinessBadgeBorder.Background = new SolidColorBrush(Color.FromRgb(220, 255, 230));
        ReadinessBadgeBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(52, 208, 88));
      }
      else
      {
        ReadinessVerdictTextBlock.Text = $"✔ Critical Checks Clear ({metrics.TotalCompleted}/{metrics.TotalApplicable})";
        ReadinessVerdictTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(3, 102, 214));
        ReadinessBadgeBorder.Background = new SolidColorBrush(Color.FromRgb(241, 248, 255));
        ReadinessBadgeBorder.BorderBrush = new SolidColorBrush(Color.FromRgb(200, 225, 255));
      }
    }

    private void AuditItemCheckBox_Click(object sender, RoutedEventArgs e)
    {
      if (sender is CheckBox cb && cb.Tag is string ruleId)
      {
        bool isChecked = cb.IsChecked == true;
        if (!_state.Checks.TryGetValue(ruleId, out var itemState) || itemState == null)
        {
          itemState = new AuditCheckItemState();
          _state.Checks[ruleId] = itemState;
        }

        itemState.IsCompleted = isChecked;
        itemState.CompletedAtUtc = isChecked ? DateTime.UtcNow.ToString("o") : null;
        itemState.Status = isChecked ? "Pass" : "Pending";

        UpdateMetricsDisplay();
        SaveStateToDisk();

        if (HideCompletedCheckBox.IsChecked == true)
        {
          RenderChecklistItems();
        }
      }
    }

    private void NotesTextBox_LostFocus(object sender, RoutedEventArgs e)
    {
      if (sender is TextBox tb && tb.Tag is string ruleId)
      {
        if (!_state.Checks.TryGetValue(ruleId, out var itemState) || itemState == null)
        {
          itemState = new AuditCheckItemState();
          _state.Checks[ruleId] = itemState;
        }

        string newNotes = tb.Text ?? string.Empty;
        if (!string.Equals(itemState.Notes, newNotes, StringComparison.Ordinal))
        {
          itemState.Notes = newNotes;
          SaveStateToDisk();
        }
      }
    }

    private void GoToChecklistTab_Click(object sender, RoutedEventArgs e)
    {
      MainTabControl.SelectedIndex = 1;
    }

    private void FilterCriteriaChanged(object sender, RoutedEventArgs e)
    {
      if (_isUpdatingUi) return;
      RenderChecklistItems();
    }

    private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
    {
      string q = SearchTextBox.Text ?? string.Empty;
      SearchPlaceholder.Visibility = string.IsNullOrEmpty(q) ? Visibility.Visible : Visibility.Collapsed;
      ClearSearchButton.Visibility = string.IsNullOrEmpty(q) ? Visibility.Collapsed : Visibility.Visible;
      RenderChecklistItems();
    }

    private void ClearSearchButton_Click(object sender, RoutedEventArgs e)
    {
      SearchTextBox.Text = string.Empty;
    }

    private void ResetChecksButton_Click(object sender, RoutedEventArgs e)
    {
      var result = MessageBox.Show(
        "Are you sure you want to uncheck all completed verification checks for this project?",
        "Reset QA/QC Checks",
        MessageBoxButton.YesNo,
        MessageBoxImage.Question
      );

      if (result == MessageBoxResult.Yes)
      {
        _state.Checks.Clear();
        SaveStateToDisk();
        RenderChecklistItems();
        UpdateMetricsDisplay();
        SetStatus("All verification checks have been reset.", isError: false);
      }
    }

    private void ReloadButton_Click(object sender, RoutedEventArgs e)
    {
      _state = AuditEngine.LoadAuditState(_folderPath);
      LoadScopeFromState();
      ReevaluateRules();
      SetStatus($"Reloaded from {AuditEngine.AuditStateFileName} ({DateTime.Now:hh:mm:ss tt})", isError: false);
    }

    private void CloseButton_Click(object sender, RoutedEventArgs e)
    {
      SaveStateToDisk();
      Close();
    }

    protected override void OnClosed(EventArgs e)
    {
      SaveStateToDisk();
      base.OnClosed(e);
    }

    private void SaveStateToDisk()
    {
      if (string.IsNullOrWhiteSpace(_folderPath))
      {
        SetStatus("Cannot save: project folder path is unknown.", isError: true);
        return;
      }

      try
      {
        AuditEngine.SaveAuditState(_folderPath, _state);
        SetStatus($"Saved to {AuditEngine.AuditStateFileName} ({DateTime.Now:hh:mm:ss tt})", isError: false);
      }
      catch (Exception ex)
      {
        SetStatus($"Save failed: {ex.Message}", isError: true);
      }
    }

    private void SetStatus(string message, bool isError)
    {
      StatusTextBlock.Text = message ?? string.Empty;
      if (isError)
      {
        StatusIcon.Text = "⚠";
        StatusIcon.Foreground = new SolidColorBrush(Color.FromRgb(215, 58, 73));
        StatusTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(215, 58, 73));
      }
      else
      {
        StatusIcon.Text = "✔";
        StatusIcon.Foreground = new SolidColorBrush(Color.FromRgb(40, 167, 69));
        StatusTextBlock.Foreground = new SolidColorBrush(Color.FromRgb(88, 96, 105));
      }
    }
  }

  public sealed class InferredMandateViewModel
  {
    public InferredMandateViewModel(InferredCodeMandate mandate)
    {
      Mandate = mandate;
    }

    public InferredCodeMandate Mandate { get; }
    public string MandateName => Mandate.MandateName;
    public string TriggerReason => Mandate.TriggerReason;

    public Brush BadgeBackground => Mandate.IsActive
      ? new SolidColorBrush(Color.FromRgb(220, 255, 230))
      : new SolidColorBrush(Color.FromRgb(246, 248, 250));

    public Brush BadgeForeground => Mandate.IsActive
      ? new SolidColorBrush(Color.FromRgb(40, 167, 69))
      : new SolidColorBrush(Color.FromRgb(88, 96, 105));
  }

  public sealed class AuditItemViewModel
  {
    public AuditItemViewModel(AuditRuleDefinition rule, bool isCompleted, string notes)
    {
      Rule = rule;
      IsCompleted = isCompleted;
      Notes = notes ?? string.Empty;
    }

    public AuditRuleDefinition Rule { get; }
    public bool IsCompleted { get; set; }
    public string Notes { get; set; }

    public Visibility ExceptionsVisibility => !string.IsNullOrWhiteSpace(Rule.ExceptionsAndNuances)
      ? Visibility.Visible
      : Visibility.Collapsed;

    public Visibility CodeCycleVisibility => !string.IsNullOrWhiteSpace(Rule.CodeCycleNotes)
      ? Visibility.Visible
      : Visibility.Collapsed;

    public string WhyItMattersText => !string.IsNullOrWhiteSpace(Rule.WhyItMatters)
      ? $"💡 Why it matters: {Rule.WhyItMatters}"
      : string.Empty;

    public Brush TitleForeground => IsCompleted
      ? new SolidColorBrush(Color.FromRgb(106, 115, 125))
      : new SolidColorBrush(Color.FromRgb(36, 41, 46));

    public Brush SeverityBackground
    {
      get
      {
        string sev = Rule.Severity ?? "High";
        if (string.Equals(sev, "Critical", StringComparison.OrdinalIgnoreCase))
          return new SolidColorBrush(Color.FromRgb(255, 235, 233));
        if (string.Equals(sev, "High", StringComparison.OrdinalIgnoreCase))
          return new SolidColorBrush(Color.FromRgb(255, 248, 225));
        return new SolidColorBrush(Color.FromRgb(241, 248, 255));
      }
    }

    public Brush SeverityForeground
    {
      get
      {
        string sev = Rule.Severity ?? "High";
        if (string.Equals(sev, "Critical", StringComparison.OrdinalIgnoreCase))
          return new SolidColorBrush(Color.FromRgb(203, 36, 49));
        if (string.Equals(sev, "High", StringComparison.OrdinalIgnoreCase))
          return new SolidColorBrush(Color.FromRgb(176, 136, 0));
        return new SolidColorBrush(Color.FromRgb(3, 102, 214));
      }
    }
  }
}
