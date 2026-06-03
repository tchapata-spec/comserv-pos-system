Public Class GlobalVariables
    Public Shared CurrentUserID As Integer
    Public Shared CurrentUsername As String
    Public Shared CurrentUserRole As String
    Public Shared DBConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\ComServPOS.accdb"
    
    ' Company Information
    Public Shared ReadOnly CompanyName As String = "ComServ Africa PBC"
    Public Shared ReadOnly CompanyTradingName As String = "T/A Compass Services Africa"
    Public Shared ReadOnly ZMRA_TIN As String = "2001429560"
    Public Shared ReadOnly PRAZ_No As String = "PR71894435240"
    Public Shared ReadOnly CompanyAddress As String = "A16 1st Floor, Thompson House, 130 Harare Street Corner Speke"
    Public Shared ReadOnly CompanyPhone As String = "0771058311"
    Public Shared ReadOnly CompanyEmail As String = "sales@comservafrica.co.zw"
    Public Shared ReadOnly CompanyWebsite As String = "www.comservafrica.co.zw"
    Public Shared ReadOnly CompanyMotto As String = "Pointing Africa towards the Future"
    
    ' Tax Rate
    Public Shared ReadOnly TAX_RATE As Decimal = 0.15D ' 15%
    
    ' Return Policy
    Public Shared ReadOnly RETURN_HANDLING_FEE As Decimal = 0.10D ' 10%
    Public Shared ReadOnly WARRANTY_PERIOD As String = "6 Months"
    
    ' Terms and Conditions
    Public Shared ReadOnly TermsAndConditions As String = _
        "TERMS AND CONDITIONS" & vbCrLf & _
        "• Price in USD valid for 30 days" & vbCrLf & _
        "• 6 Months warranty on all our products" & vbCrLf & _
        "• 6 Months warranty on service and maintenance" & vbCrLf & _
        "• 10% Handling fee on all returns" & vbCrLf & _
        "• Returns of Laptops with scratches or damaged are not accepted" & vbCrLf & _
        "• No refunds but replacement" & vbCrLf & _
        "• No warranty on charger, LCD, screen, earphones and all small gadgets"
    
    ' Non-Warranty Items
    Public Shared ReadOnly NonWarrantyItems As String() = {"Charger", "LCD", "Screen", "Earphones", "Small Gadgets"}
End Class