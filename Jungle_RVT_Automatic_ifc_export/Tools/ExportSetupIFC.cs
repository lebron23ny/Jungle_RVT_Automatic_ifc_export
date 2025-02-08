using System;

namespace Jungle_RVT_Automatic_ifc_export.Tools
{
    public class Rootobject
    {
        public int IFCVersion { get; set; }
        public int ExchangeRequirement { get; set; }
        public int IFCFileType { get; set; }


        //public int ActivePhaseId { get; set; }
        public int SpaceBoundaries { get; set; }
        public bool SplitWallsAndColumns { get; set; }
        public bool IncludeSteelElements { get; set; }

        //public Projectaddress ProjectAddress { get; set; }
        public bool Export2DElements { get; set; }
        public bool ExportLinkedFiles { get; set; }
        public bool VisibleElementsOfCurrentView { get; set; }
        public bool ExportRoomsInView { get; set; }
        public bool ExportInternalRevitPropertySets { get; set; }
        public bool ExportIFCCommonPropertySets { get; set; }
        public bool ExportBaseQuantities { get; set; }
        public bool ExportMaterialPsets { get; set; }
        public bool ExportSchedulesAsPsets { get; set; }
        public bool ExportSpecificSchedules { get; set; }
        public bool ExportUserDefinedPsets { get; set; }
        public string ExportUserDefinedPsetsFileName { get; set; }
        public bool ExportUserDefinedParameterMapping { get; set; }
        public string ExportUserDefinedParameterMappingFileName { get; set; }
        //public Classificationsettings ClassificationSettings { get; set; }
        public int TessellationLevelOfDetail { get; set; }
        public bool ExportPartsAsBuildingElements { get; set; }
        public bool ExportSolidModelRep { get; set; }
        public bool UseActiveViewGeometry { get; set; }
        public bool UseFamilyAndTypeNameForReference { get; set; }
        public bool Use2DRoomBoundaryForVolume { get; set; }
        public bool IncludeSiteElevation { get; set; }
        public bool StoreIFCGUID { get; set; }
        public bool ExportBoundingBox { get; set; }
        public bool UseOnlyTriangulation { get; set; }
        public bool UseTypeNameOnlyForIfcType { get; set; }
        public bool UseVisibleRevitNameAsEntityName { get; set; }
        public string SelectedSite { get; set; }
        public int SitePlacement { get; set; }
        public string GeoRefCRSName { get; set; }
        public string GeoRefCRSDesc { get; set; }
        public string GeoRefEPSGCode { get; set; }
        public string GeoRefGeodeticDatum { get; set; }
        public string GeoRefMapUnit { get; set; }
        public string ExcludeFilter { get; set; }
        public string COBieCompanyInfo { get; set; }
        public string COBieProjectInfo { get; set; }
        public string Name { get; set; }
    }

    public class Projectaddress
    {
        public bool UpdateProjectInformation { get; set; }
        public bool AssignAddressToSite { get; set; }
        public bool AssignAddressToBuilding { get; set; }
    }

    public class Classificationsettings
    {
        public object ClassificationName { get; set; }
        public object ClassificationEdition { get; set; }
        public object ClassificationSource { get; set; }
        public DateTime ClassificationEditionDate { get; set; }
        public object ClassificationLocation { get; set; }
        public object ClassificationFieldName { get; set; }
    }
}
