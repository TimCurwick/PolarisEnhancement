@{
    RootModule         = "$PSScriptRoot\PolarisEnhancement.psm1"
    
    ModuleVersion      = '1.0.0'
    GUID               = 'cf35428d-bd49-423b-b2a4-5057f807da1d'
    Author             = 'Tim Curwick'
    Copyright          = '© Tim Curwick'
    
    Description = 'Enhanced functionality for Polaris'
    
    FunctionsToExport  = @(
        'New-WrappedSQLQueryRoute'
        'New-WrappedScriptRoute'
        'New-WrappedCommandRoute'
        'New-SQLReportSite'
        'Build-SQLLogSite' )
}
