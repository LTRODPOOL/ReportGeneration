Public Class KPMS

    Public Assigned_To As String
    Public Breach_Exceptions As String
    Public BU_Name As String
    Public CnI_Number As String
    Public Hours_to_Close_Gross_Hrs As String
    Public Impact As String
    Public Incident_No As String
    Public Initial_Support_3_Months As String
    Public Instance_Name As String
    Public Issue_and_Problem_Type As String
    Public Issue_Location As String
    Public Issue_Type As String
    Public ITS_Status As String
    Public KP_Code As String
    Public KPMS_Closed_Date As String
    Public KPMS_Efforts As String
    Public KPMS_Resolved_Date As String
    Public KPMS_Status As String
    Public L0_Support As String
    Public Last_Modified_Date As String
    Public Macro_Category As String
    Public Model_VersionR As String
    Public Module_Application_Name As String
    Public Name_of_Multiple_Performer As String
    Public Name_of_Performer As String
    Public Net_Duration_Hrs As String
    Public Notes As String
    Public Organization As String
    Public Patch_Number As String
    Public Pending_Info_Time_Hrs As String
    Public Perf_Site As String
    Public Problem_Summary As String
    Public Problem_Type As String
    Public Product_Categarization_Version As String
    Public PRODUCT_CATEGORIZATION_TIER As String
    Public Production_Support_Percent As String
    Public Project_Code As String
    Public Remedy_Status As String
    Public Remedy_Version As String
    Public Reopened_Ticket_Yes_No As String
    Public Reporting_Manager As String
    Public Reprocessing_Flag As String
    Public Resolution As String
    Public Resolution_Adjustment As String
    Public Resolution_Categorization_Tier As String
    Public Resolution_Categorization_Tier_2 As String
    Public Resolution_Categorization_Tier_3 As String
    Public Resolution_ITS_Adjustment_comments As String
    Public Resolution_ITS_Adjustment_Reason As String
    Public Resolution_ITS_Adjustment_SLA As String
    Public Resolution_SLA_Agreed_Time As String
    Public Resolution_SLA_Status_System_Generated As String
    Public Resolution_SLA_Status_User_Modified As String
    Public Responded_Date As String
    Public Response_Adjustment As String
    Public Response_ITS_Adjustment_comments As String
    Public Response_ITS_Adjustment_Reason As String
    Public Response_ITS_Adjustment_SLA As String
    Public Response_SLA_Agreed_Time As String
    Public Response_SLA_Status_System_Generated As String
    Public Response_SLA_Status_User_Modified As String
    Public Rework_Efforts_Hrs As String
    Public Rework_Ticket_Yes_No As String
    Public Site_Name As String
    Public SLA_Breach_Reason As String
    Public Solution_Provided As String
    Public SR_Number As String
    Public Sub_BU_Name As String
    Public Ticket_Arrival_Time As String
    Public Ticket_Closed_Time As String
    Public Ticket_Complexity As String
    Public Ticket_with_KPIT_Or_OGPV As String
    Public Time_to_respond_Hrs As String
    Public Track_Name As String
    Public User_Name As String
    Public Vendor_Last_Name As String
    Public Vendor_Phone As String
    Public Week_Closed As String
    Public Week_Received As String
    Public Weekend_Holiday_Duration_Hrs As String
    Public WWID_of_Performer As String
    Public WWID_of_User As String


    Public HashTags() As String


    Sub New()

    End Sub

    Sub splitHashTags()

        Dim varHashTags() As String
        Dim varHashTag() As String

        varHashTags = Split(Resolution, "|")
        varHashTag = Split(varHashTags(0), ";")

        Me.HashTags = varHashTag

    End Sub

End Class
