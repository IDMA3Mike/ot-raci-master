
# OT Master RACI & Resource Planning — Recipe Pack

This pack contains:
- **CSV templates** for all master tables (schema-only, ready to load).
- **Power Query M** code to ingest from SharePoint folders, normalize RACI and staffing, build role maps, activities, org nodes/edges.
- **DAX measures** for the Data Model.
- **VBA macros** for department filtering and per-department CSV export + basic org chart.

## How to use
1. Open **OT_Master_RACI_MasterWorkbook.xlsx** in Excel Desktop.
2. Save as **OT_Master_RACI_MasterWorkbook.xlsm** (macro-enabled).
3. Press `Alt+F11` > Import the provided `VBA_Macros.bas` from this pack (or copy from the `Code_VBA` sheet).
4. Go to **Data > Get Data > Launch Power Query Editor**:
   - Create Parameters: `SharePointFolders`, `SourceExcelFiles`, `FuzzyMatchThreshold`, `AllowAutoFinalize`, `RoleCanonicalSource`.
   - Add **Blank Query** and paste sections from `PowerQuery_M_All.txt` into queries named: `Ingest`, `RACI_Normalized`, `Staffing_Master`, `Role_Map`, `Master_Activities`, `Master_RACI_Assignments`, `Activity_MergeProposals`, `Staffing_Ratio_Models`, `Questionnaire_Responses`, `Dependencies_Register`, `OrgNodes`, `OrgEdges`.
5. Load all these queries **To Data Model** and also **as Tables** on dedicated sheets with matching table names (so macros can find them).
6. On the Dashboard sheet, insert slicers for `Department`, `RACI`, `Category`, `RoleCanonical`, and `SourceFile`. Name the Department slicer **slcDepartment**.
7. Insert two **Form Controls** buttons and assign macros: `btnFilterDepartment_Click` and `btnExportCSVByDepartment_Click`.
8. Refresh All. Validate counts vs source files using the SourceFile slicer.

## Notes
- Exports are UTF-8 with headers.
- Org chart macro adds a basic SmartArt; for detailed network, use OrgNodes/OrgEdges in Visio/Power BI.
- Fuzzy merge proposals appear in `Activity_MergeProposals`; accept/reject by editing and re-running.
