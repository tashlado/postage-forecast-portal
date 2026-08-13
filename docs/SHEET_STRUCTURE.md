# Workbook structure snapshot
Generated: 2026-08-13T12:38:44.875Z
Spreadsheet: Version 1 Postage Forecast — Data

## Tab: Dim_Reference
- hidden: false
- used range: 30 rows x 7 cols
- frozen rows/cols: 1 / 0
- row 1: Ref_ID | List_Name | Code | Label | Sort_Order | Active | Notes

## Tab: Actuals_Import
- hidden: false
- used range: 4957 rows x 8 cols
- frozen rows/cols: 1 / 0
- row 1: Country Code | Brand | Delivery Carrier | Delivery Method | Treatment Type | Dispatched Date: Month | Count | Sum of Cost Of Shipping (Â£)

## Tab: Snapshots
- hidden: false
- used range: 3 rows x 10 cols
- frozen rows/cols: 1 / 0
- row 1: Snapshot_ID | Taken_TS | Taken_By | Trigger | Scenario_ID | Label | Rows_Stored | Changed_Rows | Calc_Run_ID | Notes

## Tab: Snapshot_Values
- hidden: false
- used range: 1261 rows x 5 cols
- frozen rows/cols: 1 / 0
- row 1: Snapshot_ID | High_Level_ID | Date_ID | Month_Start | Forecast_Rate_Per_Order

## Tab: Actuals
- hidden: false
- used range: 131 rows x 15 cols
- frozen rows/cols: 1 / 0
- row 1: Actual_ID | High_Level_ID | Month_Start | Orders | Total_Spend | Blended_Rate | Currency | Source | Source_Version | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Actuals_Amends
- hidden: false
- used range: 261 rows x 19 cols
- frozen rows/cols: 1 / 0
- row 1: Amend_ID | Amend_TS | Amend_By | Amend_Type | Actual_ID | High_Level_ID | Month_Start | Orders | Total_Spend | Blended_Rate | Currency | Source | Source_Version | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Dim_Carrier
- hidden: false
- used range: 8 rows x 5 cols
- frozen rows/cols: 1 / 0
- row 1: Carrier_Code | Carrier_Name | Default_Currency | Active | Notes

## Tab: Dim_Method
- hidden: false
- used range: 36 rows x 7 cols
- frozen rows/cols: 1 / 0
- row 1: Method_Code | Carrier_Code | Method_Name | Service_Level | Is_Tracked | Active | Notes

## Tab: Dim_Surcharge
- hidden: false
- used range: 5 rows x 8 cols
- frozen rows/cols: 1 / 0
- row 1: Surcharge_Code | Surcharge_Name | Value_Type | Applies_To | Apply_Order | Proration | Active | Notes

## Tab: Dim_Calendar
- hidden: false
- used range: 37 rows x 9 cols
- frozen rows/cols: 1 / 0
- row 1: Date_ID | Month_Start | Month_End | Days_In_Month | Year | Month_No | Month_Label | Quarter | In_Horizon

## Tab: High_Level_IDs
- hidden: false
- used range: 19 rows x 14 cols
- frozen rows/cols: 1 / 0
- row 1: High_Level_ID | High_Level_Code | Brand | Geo | Treatment_Type | WL_Split | Currency | Active | Sort_Order | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Modelling_IDs
- hidden: false
- used range: 255 rows x 12 cols
- frozen rows/cols: 1 / 0
- row 1: Modelling_ID | High_Level_ID | Carrier_Code | Method_Code | Letter_Parcel | Modelling_Code | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Rate_Base
- hidden: false
- used range: 287 rows x 14 cols
- frozen rows/cols: 1 / 0
- row 1: Rate_ID | Modelling_ID | Valid_From | Valid_To | Base_Rate | Currency | Scenario_ID | Source_Ref | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Rate_Surcharge
- hidden: false
- used range: 1650 rows x 15 cols
- frozen rows/cols: 1 / 0
- row 1: Surcharge_Rate_ID | Modelling_ID | Surcharge_Code | Valid_From | Valid_To | Value | Currency | Scenario_ID | Source_Ref | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Mix_Method
- hidden: false
- used range: 375 rows x 13 cols
- frozen rows/cols: 1 / 0
- row 1: Mix_ID | Modelling_ID | Temp_Regime | Valid_From | Valid_To | Mix_Pct | Scenario_ID | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Mix_LetterParcel
- hidden: false
- used range: 19 rows x 12 cols
- frozen rows/cols: 1 / 0
- row 1: LP_Mix_ID | High_Level_ID | Valid_From | Valid_To | Letter_Mix_Pct | Scenario_ID | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Mix_ColdChain
- hidden: false
- used range: 142 rows x 12 cols
- frozen rows/cols: 1 / 0
- row 1: CC_Mix_ID | High_Level_ID | Valid_From | Valid_To | CC_Mix_Pct | Scenario_ID | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: OUTPUT
- hidden: false
- used range: 649 rows x 11 cols
- frozen rows/cols: 1 / 0
- row 1: High_Level_ID | Date_ID | Month_Start | Brand | Geo | Treatment_Type | WL_Split | Currency | Forecast_Rate_Per_Order | Scenario_ID | Calc_Run_ID

## Tab: OUTPUT_Detail
- hidden: false
- used range: 9145 rows x 22 cols
- frozen rows/cols: 1 / 0
- row 1: Modelling_ID | High_Level_ID | Date_ID | Month_Start | Brand | Geo | Treatment_Type | WL_Split | Carrier_Code | Method_Code | Letter_Parcel | Base_Rate | Surcharge_Pct_Total | Surcharge_Amt_Total | Rate_Per_Parcel | CC_Share | Method_Mix | LP_Mix | Rate_Contribution | Rate_CTS | Scenario_ID | Calc_Run_ID

## Tab: Calc_Runs
- hidden: false
- used range: 7 rows x 11 cols
- frozen rows/cols: 1 / 0
- row 1: Calc_Run_ID | Run_TS | Run_By | Scenario_ID | Trigger | Rows_Output | Rows_Output_Detail | Duration_Ms | Validation_Status | Validation_Summary | Engine_Version

## Tab: Permissions
- hidden: false
- used range: 3 rows x 7 cols
- frozen rows/cols: 1 / 0
- row 1: Email | Display_Name | Role | Active | Tab_Visibility | Last_Login_TS | Notes

## Tab: Portal_Roles
- hidden: false
- used range: 5 rows x 10 cols
- frozen rows/cols: 1 / 0
- row 1: Role | All_Access | Write_Access | Can_Edit_Rates | Can_Edit_Mixes | Can_Edit_Structure | Can_Run_Calc | Can_Publish_Output | Can_Manage_Users | Can_View_Audit

## Tab: Scope_Mapping
- hidden: false
- used range: 1 rows x 7 cols
- frozen rows/cols: 1 / 0
- row 1: Scope_ID | Email | Scope_Type | Scope_Value | Can_View | Can_Edit | Active

## Tab: Audit_Log
- hidden: false
- used range: 768 rows x 11 cols
- frozen rows/cols: 1 / 0
- row 1: Log_ID | TS | Email | Action | Entity | Entity_ID | Field | Old_Value | New_Value | Detail | Success

## Tab: Scenarios
- hidden: false
- used range: 2 rows x 9 cols
- frozen rows/cols: 1 / 0
- row 1: Scenario_ID | Scenario_Name | Description | Parent_Scenario_ID | Is_Default | Locked | Active | Created_TS | Created_By

## Tab: Config
- hidden: false
- used range: 14 rows x 5 cols
- frozen rows/cols: 1 / 0
- row 1: Key | Value | Description | Updated_TS | Updated_By

## Tab: FX_Rates
- hidden: false
- used range: 769 rows x 6 cols
- frozen rows/cols: 1 / 0
- row 1: FX_ID | Month_Start | Currency | Rate_To_GBP | Source | Active

## Tab: Validation_Results
- hidden: false
- used range: 3 rows x 11 cols
- frozen rows/cols: 1 / 0
- row 1: Result_ID | Calc_Run_ID | Rule_Code | Severity | Entity | Entity_ID | High_Level_ID | Modelling_ID | Date_ID | Message | Resolved

## Tab: High_Level_IDs_Amends
- hidden: false
- used range: 2 rows x 18 cols
- frozen rows/cols: 1 / 0
- row 1: Amend_ID | Amend_TS | Amend_By | Amend_Type | High_Level_ID | High_Level_Code | Brand | Geo | Treatment_Type | WL_Split | Currency | Active | Sort_Order | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Modelling_IDs_Amends
- hidden: false
- used range: 25 rows x 16 cols
- frozen rows/cols: 1 / 0
- row 1: Amend_ID | Amend_TS | Amend_By | Amend_Type | Modelling_ID | High_Level_ID | Carrier_Code | Method_Code | Letter_Parcel | Modelling_Code | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Rate_Base_Amends
- hidden: false
- used range: 23 rows x 18 cols
- frozen rows/cols: 1 / 0
- row 1: Amend_ID | Amend_TS | Amend_By | Amend_Type | Rate_ID | Modelling_ID | Valid_From | Valid_To | Base_Rate | Currency | Scenario_ID | Source_Ref | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Rate_Surcharge_Amends
- hidden: false
- used range: 232 rows x 19 cols
- frozen rows/cols: 1 / 0
- row 1: Amend_ID | Amend_TS | Amend_By | Amend_Type | Surcharge_Rate_ID | Modelling_ID | Surcharge_Code | Valid_From | Valid_To | Value | Currency | Scenario_ID | Source_Ref | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Mix_Method_Amends
- hidden: false
- used range: 45 rows x 17 cols
- frozen rows/cols: 1 / 0
- row 1: Amend_ID | Amend_TS | Amend_By | Amend_Type | Mix_ID | Modelling_ID | Temp_Regime | Valid_From | Valid_To | Mix_Pct | Scenario_ID | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Mix_LetterParcel_Amends
- hidden: false
- used range: 2 rows x 16 cols
- frozen rows/cols: 1 / 0
- row 1: Amend_ID | Amend_TS | Amend_By | Amend_Type | LP_Mix_ID | High_Level_ID | Valid_From | Valid_To | Letter_Mix_Pct | Scenario_ID | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By

## Tab: Mix_ColdChain_Amends
- hidden: false
- used range: 3 rows x 16 cols
- frozen rows/cols: 1 / 0
- row 1: Amend_ID | Amend_TS | Amend_By | Amend_Type | CC_Mix_ID | High_Level_ID | Valid_From | Valid_To | CC_Mix_Pct | Scenario_ID | Active | Notes | Created_TS | Created_By | Updated_TS | Updated_By
