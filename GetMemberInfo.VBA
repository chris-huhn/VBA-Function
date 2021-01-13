' This is a fucntion to return five fields (Member ID, Last Name, First Name, State, and  Age Status) from a SQL Server database.

' Single argument 'mbrid'
Public Function GetMemberInfo(mbrid As Long) As Variant
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String
 
    ' Create the connection string.
    sConnString = "Provider=SQLOLEDB;Data Source=SERVERNAMEHERE;" & _
                  "Initial Catalog=DATABASENAMEHERE;" & _
                  "Integrated Security=SSPI;"
    
    ' Create the Connection and Recordset objects.
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' Open the connection and execute.
    conn.Open sConnString
    ' Select fields from 'members' table.
    Set rs = conn.Execute(
                        "SELECT 
                            Member_ID,
                            Last_Name,
                            First_Name, 
                            State,
                            ' If the member is under 18 years old, return "Yes"; else, return "No"
                            CASE 
                                WHEN FLOOR(DATEDIFF(DAY, Date_Of_Birth, CONVERT(VARCHAR(10), GETDATE(), 101))/365) < 18 
                                THEN 'Yes' 
                                ELSE 'No' 
                                END AS Age 
                        FROM members 
                        WHERE Member_ID= " & empID & ";")

        
    ' Verify that data is returned.
    If Not rs.EOF Then
        ' Set function equal to all fields as rows, transpose to columns.
        GetMemberInfo = Application.WorksheetFunction.Transpose(rs.GetRows)
    Else
        'Not found; return #N/A! error
        GetMemberInfo = CVErr(xlErrNA)
    End If
    ' Close the recordset
    rs.Close
    
    ' Clean up
    If CBool(conn.State And adStateOpen) Then conn.Close
    Set conn = Nothing
    Set rs = Nothing
    
End Function
