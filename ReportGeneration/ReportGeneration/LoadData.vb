Imports Microsoft.Office.Interop.Excel
Public Class LoadData
    Sub New()

    End Sub

    Public Function loadKPMSObject(ByRef objSheet As Worksheet)

        Dim ObjKPMSItem As New KPMS
        Dim CollectionKPMSItem As New List(Of KPMS)

        Dim lineOFSheet As Long
        Dim ColumnsOfSheet As Long

        lineOFSheet = 7
        ColumnsOfSheet = 1
        Do While objSheet.Cells(lineOFSheet, 1).Value <> ""
            ObjKPMSItem = New KPMS

            ObjKPMSItem.BU_Name = objSheet.Cells(lineOFSheet, 1).value
            ObjKPMSItem.Sub_BU_Name = objSheet.Cells(lineOFSheet, 2).value
            ObjKPMSItem.Track_Name = objSheet.Cells(lineOFSheet, 3).value
            ObjKPMSItem.Remedy_Status = objSheet.Cells(lineOFSheet, 4).value
            ObjKPMSItem.KPMS_Status = objSheet.Cells(lineOFSheet, 5).value
            ObjKPMSItem.Ticket_with_KPIT_Or_OGPV = objSheet.Cells(lineOFSheet, 6).value
            ObjKPMSItem.Week_Received = objSheet.Cells(lineOFSheet, 7).value
            ObjKPMSItem.Week_Closed = objSheet.Cells(lineOFSheet, 8).value
            ObjKPMSItem.Incident_No = objSheet.Cells(lineOFSheet, 9).value
            ObjKPMSItem.Site_Name = objSheet.Cells(lineOFSheet, 10).value
            ObjKPMSItem.Impact = objSheet.Cells(lineOFSheet, 11).value
            ObjKPMSItem.WWID_of_User = objSheet.Cells(lineOFSheet, 12).value
            ObjKPMSItem.User_Name = objSheet.Cells(lineOFSheet, 13).value
            ObjKPMSItem.Assigned_To = objSheet.Cells(lineOFSheet, 14).value
            ObjKPMSItem.Instance_Name = objSheet.Cells(lineOFSheet, 15).value
            ObjKPMSItem.Ticket_Arrival_Time = objSheet.Cells(lineOFSheet, 16).value
            ObjKPMSItem.Last_Modified_Date = objSheet.Cells(lineOFSheet, 17).value
            ObjKPMSItem.Ticket_Closed_Time = objSheet.Cells(lineOFSheet, 18).value
            ObjKPMSItem.Hours_to_Close_Gross_Hrs = objSheet.Cells(lineOFSheet, 19).value
            ObjKPMSItem.Pending_Info_Time_Hrs = objSheet.Cells(lineOFSheet, 20).value
            ObjKPMSItem.Weekend_Holiday_Duration_Hrs = objSheet.Cells(lineOFSheet, 21).value
            ObjKPMSItem.Net_Duration_Hrs = objSheet.Cells(lineOFSheet, 22).value
            ObjKPMSItem.Resolution_SLA_Agreed_Time = objSheet.Cells(lineOFSheet, 23).value
            ObjKPMSItem.Resolution_SLA_Status_System_Generated = objSheet.Cells(lineOFSheet, 24).value
            ObjKPMSItem.Resolution_SLA_Status_User_Modified = objSheet.Cells(lineOFSheet, 25).value
            ObjKPMSItem.Resolution_Adjustment = objSheet.Cells(lineOFSheet, 26).value
            ObjKPMSItem.KPMS_Efforts = objSheet.Cells(lineOFSheet, 27).value
            ObjKPMSItem.Response_SLA_Agreed_Time = objSheet.Cells(lineOFSheet, 28).value
            ObjKPMSItem.Time_to_respond_Hrs = objSheet.Cells(lineOFSheet, 29).value
            ObjKPMSItem.Response_SLA_Status_System_Generated = objSheet.Cells(lineOFSheet, 30).value
            ObjKPMSItem.Response_SLA_Status_User_Modified = objSheet.Cells(lineOFSheet, 31).value
            ObjKPMSItem.Response_Adjustment = objSheet.Cells(lineOFSheet, 32).value
            ObjKPMSItem.Reopened_Ticket_Yes_No = objSheet.Cells(lineOFSheet, 33).value
            ObjKPMSItem.Rework_Ticket_Yes_No = objSheet.Cells(lineOFSheet, 34).value
            ObjKPMSItem.Rework_Efforts_Hrs = objSheet.Cells(lineOFSheet, 35).value
            ObjKPMSItem.WWID_of_Performer = objSheet.Cells(lineOFSheet, 36).value
            ObjKPMSItem.Name_of_Performer = objSheet.Cells(lineOFSheet, 37).value
            ObjKPMSItem.Name_of_Multiple_Performer = objSheet.Cells(lineOFSheet, 38).value
            ObjKPMSItem.Problem_Summary = objSheet.Cells(lineOFSheet, 39).value
            ObjKPMSItem.SR_Number = objSheet.Cells(lineOFSheet, 40).value
            ObjKPMSItem.Organization = objSheet.Cells(lineOFSheet, 41).value
            ObjKPMSItem.Initial_Support_3_Months = objSheet.Cells(lineOFSheet, 42).value
            ObjKPMSItem.Ticket_Complexity = objSheet.Cells(lineOFSheet, 43).value
            ObjKPMSItem.Issue_Type = objSheet.Cells(lineOFSheet, 44).value
            ObjKPMSItem.L0_Support = objSheet.Cells(lineOFSheet, 45).value
            ObjKPMSItem.Notes = objSheet.Cells(lineOFSheet, 46).value
            ObjKPMSItem.Resolution = objSheet.Cells(lineOFSheet, 47).value
            ObjKPMSItem.Module_Application_Name = objSheet.Cells(lineOFSheet, 48).value
            ObjKPMSItem.Vendor_Last_Name = objSheet.Cells(lineOFSheet, 49).value
            ObjKPMSItem.Vendor_Phone = objSheet.Cells(lineOFSheet, 50).value
            ObjKPMSItem.Problem_Type = objSheet.Cells(lineOFSheet, 51).value
            ObjKPMSItem.Solution_Provided = objSheet.Cells(lineOFSheet, 52).value
            ObjKPMSItem.KPMS_Closed_Date = objSheet.Cells(lineOFSheet, 53).value
            ObjKPMSItem.Patch_Number = objSheet.Cells(lineOFSheet, 54).value
            ObjKPMSItem.CnI_Number = objSheet.Cells(lineOFSheet, 55).value
            ObjKPMSItem.PRODUCT_CATEGORIZATION_TIER = objSheet.Cells(lineOFSheet, 56).value
            ObjKPMSItem.Issue_Location = objSheet.Cells(lineOFSheet, 57).value
            ObjKPMSItem.Perf_Site = objSheet.Cells(lineOFSheet, 58).value
            ObjKPMSItem.Macro_Category = objSheet.Cells(lineOFSheet, 59).value
            ObjKPMSItem.KPMS_Resolved_Date = objSheet.Cells(lineOFSheet, 60).value
            ObjKPMSItem.Reporting_Manager = objSheet.Cells(lineOFSheet, 61).value
            ObjKPMSItem.Issue_and_Problem_Type = objSheet.Cells(lineOFSheet, 62).value
            ObjKPMSItem.Resolution_Categorization_Tier = objSheet.Cells(lineOFSheet, 63).value
            ObjKPMSItem.Resolution_Categorization_Tier_2 = objSheet.Cells(lineOFSheet, 64).value
            ObjKPMSItem.Resolution_Categorization_Tier_3 = objSheet.Cells(lineOFSheet, 65).value
            ObjKPMSItem.Production_Support_Percent = objSheet.Cells(lineOFSheet, 66).value
            ObjKPMSItem.Model_VersionR = objSheet.Cells(lineOFSheet, 67).value
            ObjKPMSItem.Remedy_Version = objSheet.Cells(lineOFSheet, 68).value
            ObjKPMSItem.KP_Code = objSheet.Cells(lineOFSheet, 69).value
            ObjKPMSItem.Project_Code = objSheet.Cells(lineOFSheet, 70).value
            ObjKPMSItem.Breach_Exceptions = objSheet.Cells(lineOFSheet, 71).value
            ObjKPMSItem.Product_Categarization_Version = objSheet.Cells(lineOFSheet, 72).value
            ObjKPMSItem.Reprocessing_Flag = objSheet.Cells(lineOFSheet, 73).value
            ObjKPMSItem.ITS_Status = objSheet.Cells(lineOFSheet, 74).value
            ObjKPMSItem.Resolution_ITS_Adjustment_SLA = objSheet.Cells(lineOFSheet, 75).value
            ObjKPMSItem.Resolution_ITS_Adjustment_Reason = objSheet.Cells(lineOFSheet, 76).value
            ObjKPMSItem.Resolution_ITS_Adjustment_comments = objSheet.Cells(lineOFSheet, 77).value
            ObjKPMSItem.Response_ITS_Adjustment_SLA = objSheet.Cells(lineOFSheet, 78).value
            ObjKPMSItem.Response_ITS_Adjustment_Reason = objSheet.Cells(lineOFSheet, 79).value
            ObjKPMSItem.Response_ITS_Adjustment_Reason = objSheet.Cells(lineOFSheet, 80).value
            ObjKPMSItem.Responded_Date = objSheet.Cells(lineOFSheet, 81).value
            ObjKPMSItem.SLA_Breach_Reason = objSheet.Cells(lineOFSheet, 82).value

            Call ObjKPMSItem.splitHashTags()

            Call CollectionKPMSItem.Add(ObjKPMSItem)

            lineOFSheet = lineOFSheet + 1
        Loop

        Return CollectionKPMSItem

    End Function


End Class
