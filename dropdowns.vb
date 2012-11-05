Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Data
Imports System.Data.SqlClient

Namespace DDL

    Public Module mDropdowns

        ''' <summary>
        ''' Loads a dropdown list with values from the database
        ''' </summary>
        ''' <param name="ddlControl">The drop-down list control to be filled</param>
        ''' <param name="strTable">The database table to pull the data from</param>
        ''' <param name="strText">The name of the field to be used as the displayed text</param>
        ''' <param name="strValue">The name of the field to be used as the item value</param>
        ''' <param name="strWhere">A where clause containing criteria that the record set should be filtered by (do not include the WHERE keyword)</param>
        ''' <param name="strOrderBy">An order by clause containing a list of fields that the record set should be sorted by (do not include the ORDER BY keywords)</param>
        ''' <param name="blnSelect">Boolean value indicating whether a "Select an Item" item should be included</param>
        ''' <param name="blnShortSelect">Boolean value indicating if the "Select an Item" item should be shortened to just "--"</param>
        ''' <param name="strCustomSelText">A custom text value to be displayed instead of "Select an Item"</param>
        ''' <param name="blnAll">Boolean value indicating whether an "All Items" item should be included</param>
        ''' <param name="strCustomAllText">A custom text value to be displayed instead of "All Items"</param>
        Public Sub LoadDDL(ByRef ddlControl As DropDownList, ByVal strTable As String, _
            Optional ByVal strText As String = "Name", Optional ByVal strValue As String = "ID", _
            Optional ByVal strWhere As String = "", Optional ByVal strOrderBy As String = "", _
            Optional ByVal blnSelect As Boolean = True, Optional ByVal blnShortSelect As Boolean = False, Optional ByVal strCustomSelText As String = "", _
            Optional ByVal blnAll As Boolean = False, Optional ByVal strCustomAllText As String = "All")
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Loads a dropdown list with values from the database
            ' Modified:     2012.07.23 by JR - Added AppendDataBoundItems to make sure that 
            '               the select and all list items are preserved
            ' *****************************************************************************
            Dim objConn As New SqlConnection(Configuration.WebConfigurationManager.ConnectionStrings(My.Settings.ConnName).ToString)
            Dim cmdGet As New SqlCommand()
            Dim objAdapter As New SqlDataAdapter(cmdGet)
            Dim objTable As New DataTable
            Dim strQuery As String

            'clear the existing list items
            ddlControl.Items.Clear()
            ddlControl.AppendDataBoundItems = True

            'check if multiple fields will be used to make up the text field
            If strText.IndexOf("+") < 0 And strText.IndexOf("(") < 0 Then
                strText = "[" & strText & "]"
            End If

            'create the sql query
            strQuery = "SELECT DISTINCT [" & strValue & "] AS [Value], " & _
                            strText & " AS [Text]"
            If strOrderBy <> "" Then
                strQuery &= ", " & strOrderBy.Replace("DESC", "").Replace("ASC", "")
            End If
            strQuery &= " FROM " & strTable & " "
            If strWhere.Trim <> "" Then
                strQuery = strQuery & "WHERE " & strWhere.Trim & " "
            End If
            If strOrderBy.Trim <> "" Then
                strQuery = strQuery & "ORDER BY " & strOrderBy & ";"
            Else
                strQuery = strQuery & "ORDER BY " & strText & ";"
            End If

            'set up the command object
            cmdGet.Connection = objConn
            cmdGet.CommandType = CommandType.Text
            cmdGet.CommandText = strQuery

            'open the connection
            objConn.Open()

            'execute the query
            objAdapter.Fill(objTable)
            objConn.Close()

            'add each of the records to the ddl
            ddlControl.DataSource = objTable
            ddlControl.DataTextField = "Text"
            ddlControl.DataValueField = "Value"
            ddlControl.DataBind()

            'add a "select all" option
            If blnAll Then
                Dim objItem As New ListItem(strCustomAllText, 0)
                ddlControl.Items.Insert(0, objItem)

                'TODO: figure out why this call doesn't work
                'DDLInsert_All(ddlControl, strCustomAllText)
            End If

            'add a "select an item" option
            If blnSelect Then
                Dim objItem As ListItem

                'create the item
                If strCustomSelText = "" Then
                    If blnShortSelect Then
                        objItem = New ListItem("--", "-1")
                    Else
                        objItem = New ListItem("[Select an Item]", "-1")
                    End If
                Else
                    objItem = New ListItem(strCustomSelText, "-1")
                End If

                'add the item
                ddlControl.Items.Insert(0, objItem)

                'TODO: figure out why this call doesn't work
                'DDLInsert_SelectItem(ddlControl, strCustomSelText, blnShortSelect)
            End If

            'select the first item
            ddlControl.ClearSelection()

            'clean up
            objConn = Nothing
        End Sub

        ''' <summary>
        ''' Loads a dropdown list with an "All" item
        ''' </summary>
        ''' <param name="ddlDropDown">The drop-down list to be loaded</param>
        ''' <param name="strCustomText">Custom text to be used instead of "All"</param>
        Public Sub DDLInsert_All(ByRef ddlDropDown As DropDownList, Optional ByRef strCustomText As String = "All")
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Loads a dropdown list with an "All" item
            ' *****************************************************************************
            'insert a "select all" option
            Dim objItem As New ListItem(strCustomText, 0)
            ddlDropDown.Items.Insert(0, objItem)
        End Sub

        ''' <summary>
        ''' Loads a dropdown list with a "Select an Item" item
        ''' </summary>
        ''' <param name="ddlDropDown">The drop-down list to be loaded</param>
        ''' <param name="strCustomText">Custom text to be used instead of "Select an Item"</param>
        ''' <param name="blnShort">Boolean value indicating if the "Select an Item" item should be shortened to just "--"</param>
        ''' <remarks>If strCustomText is provided, then blnShort will be ignored</remarks>
        Public Sub DDLInsert_SelectItem(ByRef ddlDropDown As DropDownList, Optional ByVal strCustomText As String = "", Optional ByVal blnShort As Boolean = False)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Loads a dropdown list with a "Select an Item" item
            ' *****************************************************************************
            Dim objItem As ListItem

            'create the item
            If strCustomText = "" Then
                If blnShort Then
                    objItem = New ListItem("--", "-1")
                Else
                    objItem = New ListItem("[Select an Item]", "-1")
                End If
            Else
                objItem = New ListItem(strCustomText, "-1")
            End If

            'add the item
            ddlDropDown.Items.Insert(0, objItem)
        End Sub

        ''' <summary>
        ''' Selects the item with the give value in the dropdown list
        ''' </summary>
        ''' <param name="ddlDropDown">The dropdown list to be selected</param>
        ''' <param name="strValue">The value of the item to be selected</param>
        ''' <remarks></remarks>
        Public Sub SelectDDL(ByRef ddlDropDown As DropDownList, ByVal strValue As String)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Selects the item with the give value in the dropdown list
            ' *****************************************************************************
            Dim liItem As ListItem

            ddlDropDown.ClearSelection()

            'search each ddl item for the value
            For Each liItem In ddlDropDown.Items
                If liItem.Value = strValue Then
                    liItem.Selected = True
                Else
                    liItem.Selected = False
                End If
            Next
        End Sub

        ''' <summary>
        ''' Adds a new item to the given drop-down list
        ''' </summary>
        ''' <param name="ddlControl">The drop-down list to add the item to</param>
        ''' <param name="strText">The text value of the new item</param>
        ''' <param name="strValue">The value of the new item</param>
        Public Sub AddDDLItem(ByRef ddlControl As DropDownList, ByVal strText As String, ByVal strValue As String)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Adds a new item to the given drop-down list
            ' *****************************************************************************
            Dim objItem As ListItem

            'add the item
            objItem = New ListItem(strText, strValue)
            ddlControl.Items.Add(objItem)
        End Sub

        ''' <summary>
        ''' Adds items for every US state to the given drop-down list
        ''' </summary>
        ''' <param name="ddlState">The drop-down list to add the items to</param>
        ''' <param name="strSelected">The state to be selected by default</param>
        ''' <param name="blnSelect">Boolean value indicating is a "Select an Item" item should be included</param>
        Public Sub StateDDLInit(ByRef ddlState As DropDownList, ByVal strSelected As String, Optional ByVal blnSelect As Boolean = True)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Adds items for every US state to the given drop-down list
            ' *****************************************************************************
            ddlState.Items.Clear()

            'add a default select value
            If blnSelect Then
                DDLInsert_SelectItem(ddlState, True)
            End If

            'add all of the states
            ddlState.Items.Add(New ListItem("AL", "AL"))
            ddlState.Items.Add(New ListItem("AK", "AK"))
            ddlState.Items.Add(New ListItem("AZ", "AZ"))
            ddlState.Items.Add(New ListItem("AR", "AR"))
            ddlState.Items.Add(New ListItem("CA", "CA"))
            ddlState.Items.Add(New ListItem("CO", "CO"))
            ddlState.Items.Add(New ListItem("CT", "CT"))
            ddlState.Items.Add(New ListItem("DE", "DE"))
            ddlState.Items.Add(New ListItem("DC", "DC"))
            ddlState.Items.Add(New ListItem("FL", "FL"))
            ddlState.Items.Add(New ListItem("GA", "GA"))
            ddlState.Items.Add(New ListItem("HI", "HI"))
            ddlState.Items.Add(New ListItem("ID", "ID"))
            ddlState.Items.Add(New ListItem("IL", "IL"))
            ddlState.Items.Add(New ListItem("IN", "IN"))
            ddlState.Items.Add(New ListItem("IA", "IA"))
            ddlState.Items.Add(New ListItem("KS", "KS"))
            ddlState.Items.Add(New ListItem("KY", "KY"))
            ddlState.Items.Add(New ListItem("LA", "LA"))
            ddlState.Items.Add(New ListItem("MA", "MA"))
            ddlState.Items.Add(New ListItem("ME", "ME"))
            ddlState.Items.Add(New ListItem("MD", "MD"))
            ddlState.Items.Add(New ListItem("MI", "MI"))
            ddlState.Items.Add(New ListItem("MN", "MN"))
            ddlState.Items.Add(New ListItem("MS", "MS"))
            ddlState.Items.Add(New ListItem("MO", "MO"))
            ddlState.Items.Add(New ListItem("MT", "MT"))
            ddlState.Items.Add(New ListItem("NE", "NE"))
            ddlState.Items.Add(New ListItem("NV", "NV"))
            ddlState.Items.Add(New ListItem("NH", "NH"))
            ddlState.Items.Add(New ListItem("NJ", "NJ"))
            ddlState.Items.Add(New ListItem("NM", "NM"))
            ddlState.Items.Add(New ListItem("NY", "NY"))
            ddlState.Items.Add(New ListItem("NC", "NC"))
            ddlState.Items.Add(New ListItem("ND", "ND"))
            ddlState.Items.Add(New ListItem("OH", "OH"))
            ddlState.Items.Add(New ListItem("OK", "OK"))
            ddlState.Items.Add(New ListItem("OR", "OR"))
            ddlState.Items.Add(New ListItem("PA", "PA"))
            ddlState.Items.Add(New ListItem("RI", "RI"))
            ddlState.Items.Add(New ListItem("SC", "SC"))
            ddlState.Items.Add(New ListItem("SD", "SD"))
            ddlState.Items.Add(New ListItem("TN", "TN"))
            ddlState.Items.Add(New ListItem("TX", "TX"))
            ddlState.Items.Add(New ListItem("UT", "UT"))
            ddlState.Items.Add(New ListItem("VT", "VT"))
            ddlState.Items.Add(New ListItem("VA", "VA"))
            ddlState.Items.Add(New ListItem("WA", "WA"))
            ddlState.Items.Add(New ListItem("WV", "WV"))
            ddlState.Items.Add(New ListItem("WI", "WI"))
            ddlState.Items.Add(New ListItem("WY", "WY"))

            'select the given state
            ddlState.SelectedValue = strSelected
        End Sub

        ''' <summary>
        ''' Adds items for every country to the given drop-down list
        ''' </summary>
        ''' <param name="ddlCountries">The drop-down list to add the items to</param>
        ''' <param name="strSelected">The country to be selected by default</param>
        ''' <param name="blnSelect">Boolean value indicating is a "Select an Item" item should be included</param>
        Public Sub CountryDDLInit(ByRef ddlCountries As DropDownList, ByVal strSelected As String, Optional ByVal blnSelect As Boolean = True)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Adds items for every country to the given drop-down list
            ' *****************************************************************************
            ddlCountries.Items.Clear()

            'add a default select value
            If blnSelect Then
                DDLInsert_SelectItem(ddlCountries)
            End If

            'add all of the countries
            ddlCountries.Items.Add(New ListItem("United States", "United States"))
            ddlCountries.Items.Add(New ListItem("Canada", "Canada"))
            ddlCountries.Items.Add(New ListItem("Mexico", "Mexico"))
            ddlCountries.Items.Add(New ListItem("Afghanistan", "Afghanistan"))
            ddlCountries.Items.Add(New ListItem("Albania", "Albania"))
            ddlCountries.Items.Add(New ListItem("Algeria", "Algeria"))
            ddlCountries.Items.Add(New ListItem("Andorra", "Andorra"))
            ddlCountries.Items.Add(New ListItem("Angola", "Angola"))
            ddlCountries.Items.Add(New ListItem("Antigua &amp; Barbuda", "Antigua &amp; Barbuda"))
            ddlCountries.Items.Add(New ListItem("Argentina", "Argentina"))
            ddlCountries.Items.Add(New ListItem("Armenia", "Armenia"))
            ddlCountries.Items.Add(New ListItem("Australia", "Australia"))
            ddlCountries.Items.Add(New ListItem("Austria", "Austria"))
            ddlCountries.Items.Add(New ListItem("Azerbaijan", "Azerbaijan"))
            ddlCountries.Items.Add(New ListItem("Bahamas", "Bahamas"))
            ddlCountries.Items.Add(New ListItem("Bahrain", "Bahrain"))
            ddlCountries.Items.Add(New ListItem("Bangladesh", "Bangladesh"))
            ddlCountries.Items.Add(New ListItem("Barbados", "Barbados"))
            ddlCountries.Items.Add(New ListItem("Belarus", "Belarus"))
            ddlCountries.Items.Add(New ListItem("Belgium", "Belgium"))
            ddlCountries.Items.Add(New ListItem("Belize", "Belize"))
            ddlCountries.Items.Add(New ListItem("Benin", "Benin"))
            ddlCountries.Items.Add(New ListItem("Bhutan", "Bhutan"))
            ddlCountries.Items.Add(New ListItem("Bolivia", "Bolivia"))
            ddlCountries.Items.Add(New ListItem("Bosnia &amp; Herzegovina", "Bosnia &amp; Herzegovina"))
            ddlCountries.Items.Add(New ListItem("Botswana", "Botswana"))
            ddlCountries.Items.Add(New ListItem("Brazil", "Brazil"))
            ddlCountries.Items.Add(New ListItem("Brunei", "Brunei"))
            ddlCountries.Items.Add(New ListItem("Bulgaria", "Bulgaria"))
            ddlCountries.Items.Add(New ListItem("Burkina Faso", "Burkina Faso"))
            ddlCountries.Items.Add(New ListItem("Burundi", "Burundi"))
            ddlCountries.Items.Add(New ListItem("Cambodia", "Cambodia"))
            ddlCountries.Items.Add(New ListItem("Cameroon", "Cameroon"))
            ddlCountries.Items.Add(New ListItem("Cape Verde", "Cape Verde"))
            ddlCountries.Items.Add(New ListItem("Central African Republic", "Central African Republic"))
            ddlCountries.Items.Add(New ListItem("Chad", "Chad"))
            ddlCountries.Items.Add(New ListItem("Chile", "Chile"))
            ddlCountries.Items.Add(New ListItem("China", "China"))
            ddlCountries.Items.Add(New ListItem("Colombia", "Colombia"))
            ddlCountries.Items.Add(New ListItem("Comoros", "Comoros"))
            ddlCountries.Items.Add(New ListItem("Congo", "Congo"))
            ddlCountries.Items.Add(New ListItem("Costa Rica", "Costa Rica"))
            ddlCountries.Items.Add(New ListItem("Croatia", "Croatia"))
            ddlCountries.Items.Add(New ListItem("Cuba", "Cuba"))
            ddlCountries.Items.Add(New ListItem("Cyprus", "Cyprus"))
            ddlCountries.Items.Add(New ListItem("Czech Republic", "Czech Republic"))
            ddlCountries.Items.Add(New ListItem("Côte d'Ivoire", "Côte d'Ivoire"))
            ddlCountries.Items.Add(New ListItem("Denmark", "Denmark"))
            ddlCountries.Items.Add(New ListItem("Djibouti", "Djibouti"))
            ddlCountries.Items.Add(New ListItem("Dominica", "Dominica"))
            ddlCountries.Items.Add(New ListItem("Dominican Republic", "Dominican Republic"))
            ddlCountries.Items.Add(New ListItem("East Timor", "East Timor"))
            ddlCountries.Items.Add(New ListItem("Ecuador", "Ecuador"))
            ddlCountries.Items.Add(New ListItem("Egypt", "Egypt"))
            ddlCountries.Items.Add(New ListItem("El Salvador", "El Salvador"))
            ddlCountries.Items.Add(New ListItem("Equatorial Guinea", "Equatorial Guinea"))
            ddlCountries.Items.Add(New ListItem("Eritrea", "Eritrea"))
            ddlCountries.Items.Add(New ListItem("Estonia", "Estonia"))
            ddlCountries.Items.Add(New ListItem("Ethiopia", "Ethiopia"))
            ddlCountries.Items.Add(New ListItem("Fiji", "Fiji"))
            ddlCountries.Items.Add(New ListItem("Finland", "Finland"))
            ddlCountries.Items.Add(New ListItem("France", "France"))
            ddlCountries.Items.Add(New ListItem("Gabon", "Gabon"))
            ddlCountries.Items.Add(New ListItem("Gambia, The", "Gambia, The"))
            ddlCountries.Items.Add(New ListItem("Georgia", "Georgia"))
            ddlCountries.Items.Add(New ListItem("Germany", "Germany"))
            ddlCountries.Items.Add(New ListItem("Ghana", "Ghana"))
            ddlCountries.Items.Add(New ListItem("Greece", "Greece"))
            ddlCountries.Items.Add(New ListItem("Grenada", "Grenada"))
            ddlCountries.Items.Add(New ListItem("Guatemala", "Guatemala"))
            ddlCountries.Items.Add(New ListItem("Guinea", "Guinea"))
            ddlCountries.Items.Add(New ListItem("Guinea-Bissau", "Guinea-Bissau"))
            ddlCountries.Items.Add(New ListItem("Guyana", "Guyana"))
            ddlCountries.Items.Add(New ListItem("Haiti", "Haiti"))
            ddlCountries.Items.Add(New ListItem("Honduras", "Honduras"))
            ddlCountries.Items.Add(New ListItem("Hungary", "Hungary"))
            ddlCountries.Items.Add(New ListItem("Iceland", "Iceland"))
            ddlCountries.Items.Add(New ListItem("India", "India"))
            ddlCountries.Items.Add(New ListItem("Indonesia", "Indonesia"))
            ddlCountries.Items.Add(New ListItem("Iran", "Iran"))
            ddlCountries.Items.Add(New ListItem("Iraq", "Iraq"))
            ddlCountries.Items.Add(New ListItem("Ireland", "Ireland"))
            ddlCountries.Items.Add(New ListItem("Israel", "Israel"))
            ddlCountries.Items.Add(New ListItem("Italy", "Italy"))
            ddlCountries.Items.Add(New ListItem("Jamaica", "Jamaica"))
            ddlCountries.Items.Add(New ListItem("Japan", "Japan"))
            ddlCountries.Items.Add(New ListItem("Jordan", "Jordan"))
            ddlCountries.Items.Add(New ListItem("Kazakhstan", "Kazakhstan"))
            ddlCountries.Items.Add(New ListItem("Kenya", "Kenya"))
            ddlCountries.Items.Add(New ListItem("Kiribati", "Kiribati"))
            ddlCountries.Items.Add(New ListItem("Kuwait", "Kuwait"))
            ddlCountries.Items.Add(New ListItem("Kyrgyzstan", "Kyrgyzstan"))
            ddlCountries.Items.Add(New ListItem("Laos", "Laos"))
            ddlCountries.Items.Add(New ListItem("Latvia", "Latvia"))
            ddlCountries.Items.Add(New ListItem("Lebanon", "Lebanon"))
            ddlCountries.Items.Add(New ListItem("Lesotho", "Lesotho"))
            ddlCountries.Items.Add(New ListItem("Liberia", "Liberia"))
            ddlCountries.Items.Add(New ListItem("Libya", "Libya"))
            ddlCountries.Items.Add(New ListItem("Liechtenstein", "Liechtenstein"))
            ddlCountries.Items.Add(New ListItem("Lithuania", "Lithuania"))
            ddlCountries.Items.Add(New ListItem("Luxembourg", "Luxembourg"))
            ddlCountries.Items.Add(New ListItem("Macedonia", "Macedonia"))
            ddlCountries.Items.Add(New ListItem("Madagascar", "Madagascar"))
            ddlCountries.Items.Add(New ListItem("Malawi", "Malawi"))
            ddlCountries.Items.Add(New ListItem("Malaysia", "Malaysia"))
            ddlCountries.Items.Add(New ListItem("Maldives", "Maldives"))
            ddlCountries.Items.Add(New ListItem("Mali", "Mali"))
            ddlCountries.Items.Add(New ListItem("Malta", "Malta"))
            ddlCountries.Items.Add(New ListItem("Marshall Islands", "Marshall Islands"))
            ddlCountries.Items.Add(New ListItem("Mauritania", "Mauritania"))
            ddlCountries.Items.Add(New ListItem("Mauritius", "Mauritius"))
            ddlCountries.Items.Add(New ListItem("Micronesia", "Micronesia"))
            ddlCountries.Items.Add(New ListItem("Moldova", "Moldova"))
            ddlCountries.Items.Add(New ListItem("Monaco", "Monaco"))
            ddlCountries.Items.Add(New ListItem("Mongolia", "Mongolia"))
            ddlCountries.Items.Add(New ListItem("Montenegro", "Montenegro"))
            ddlCountries.Items.Add(New ListItem("Morocco", "Morocco"))
            ddlCountries.Items.Add(New ListItem("Mozambique", "Mozambique"))
            ddlCountries.Items.Add(New ListItem("Myanmar", "Myanmar"))
            ddlCountries.Items.Add(New ListItem("Namibia", "Namibia"))
            ddlCountries.Items.Add(New ListItem("Nauru", "Nauru"))
            ddlCountries.Items.Add(New ListItem("Nepal", "Nepal"))
            ddlCountries.Items.Add(New ListItem("Netherlands", "Netherlands"))
            ddlCountries.Items.Add(New ListItem("New Zealand", "New Zealand"))
            ddlCountries.Items.Add(New ListItem("Nicaragua", "Nicaragua"))
            ddlCountries.Items.Add(New ListItem("Niger", "Niger"))
            ddlCountries.Items.Add(New ListItem("Nigeria", "Nigeria"))
            ddlCountries.Items.Add(New ListItem("North Korea", "North Korea"))
            ddlCountries.Items.Add(New ListItem("Norway", "Norway"))
            ddlCountries.Items.Add(New ListItem("Oman", "Oman"))
            ddlCountries.Items.Add(New ListItem("Pakistan", "Pakistan"))
            ddlCountries.Items.Add(New ListItem("Palau", "Palau"))
            ddlCountries.Items.Add(New ListItem("Panama", "Panama"))
            ddlCountries.Items.Add(New ListItem("Papua New Guinea", "Papua New Guinea"))
            ddlCountries.Items.Add(New ListItem("Paraguay", "Paraguay"))
            ddlCountries.Items.Add(New ListItem("Peru", "Peru"))
            ddlCountries.Items.Add(New ListItem("Philippines", "Philippines"))
            ddlCountries.Items.Add(New ListItem("Poland", "Poland"))
            ddlCountries.Items.Add(New ListItem("Portugal", "Portugal"))
            ddlCountries.Items.Add(New ListItem("Qatar", "Qatar"))
            ddlCountries.Items.Add(New ListItem("Romania", "Romania"))
            ddlCountries.Items.Add(New ListItem("Russia", "Russia"))
            ddlCountries.Items.Add(New ListItem("Rwanda", "Rwanda"))
            ddlCountries.Items.Add(New ListItem("Saint Kitts &amp; Nevis", "Saint Kitts &amp; Nevis"))
            ddlCountries.Items.Add(New ListItem("Saint Lucia", "Saint Lucia"))
            ddlCountries.Items.Add(New ListItem("Saint Vincent", "Saint Vincent"))
            ddlCountries.Items.Add(New ListItem("Samoa", "Samoa"))
            ddlCountries.Items.Add(New ListItem("San Marino", "San Marino"))
            ddlCountries.Items.Add(New ListItem("Sao Tome &amp; Principe", "Sao Tome &amp; Principe"))
            ddlCountries.Items.Add(New ListItem("Saudi Arabia", "Saudi Arabia"))
            ddlCountries.Items.Add(New ListItem("Senegal", "Senegal"))
            ddlCountries.Items.Add(New ListItem("Serbia", "Serbia"))
            ddlCountries.Items.Add(New ListItem("Seychelles", "Seychelles"))
            ddlCountries.Items.Add(New ListItem("Sierra Leone", "Sierra Leone"))
            ddlCountries.Items.Add(New ListItem("Singapore", "Singapore"))
            ddlCountries.Items.Add(New ListItem("Slovakia", "Slovakia"))
            ddlCountries.Items.Add(New ListItem("Slovenia", "Slovenia"))
            ddlCountries.Items.Add(New ListItem("Solomon Islands", "Solomon Islands"))
            ddlCountries.Items.Add(New ListItem("Somalia", "Somalia"))
            ddlCountries.Items.Add(New ListItem("South Africa", "South Africa"))
            ddlCountries.Items.Add(New ListItem("South Korea", "South Korea"))
            ddlCountries.Items.Add(New ListItem("Spain", "Spain"))
            ddlCountries.Items.Add(New ListItem("Sri Lanka", "Sri Lanka"))
            ddlCountries.Items.Add(New ListItem("Sudan", "Sudan"))
            ddlCountries.Items.Add(New ListItem("Suriname", "Suriname"))
            ddlCountries.Items.Add(New ListItem("Swaziland", "Swaziland"))
            ddlCountries.Items.Add(New ListItem("Sweden", "Sweden"))
            ddlCountries.Items.Add(New ListItem("Switzerland", "Switzerland"))
            ddlCountries.Items.Add(New ListItem("Syria", "Syria"))
            ddlCountries.Items.Add(New ListItem("Taiwan", "Taiwan"))
            ddlCountries.Items.Add(New ListItem("Tajikistan", "Tajikistan"))
            ddlCountries.Items.Add(New ListItem("Tanzania", "Tanzania"))
            ddlCountries.Items.Add(New ListItem("Thailand", "Thailand"))
            ddlCountries.Items.Add(New ListItem("Togo", "Togo"))
            ddlCountries.Items.Add(New ListItem("Tonga", "Tonga"))
            ddlCountries.Items.Add(New ListItem("Trinidad &amp; Tobago", "Trinidad &amp; Tobago"))
            ddlCountries.Items.Add(New ListItem("Tunisia", "Tunisia"))
            ddlCountries.Items.Add(New ListItem("Turkey", "Turkey"))
            ddlCountries.Items.Add(New ListItem("Turkmenistan", "Turkmenistan"))
            ddlCountries.Items.Add(New ListItem("Tuvalu", "Tuvalu"))
            ddlCountries.Items.Add(New ListItem("Uganda", "Uganda"))
            ddlCountries.Items.Add(New ListItem("Ukraine", "Ukraine"))
            ddlCountries.Items.Add(New ListItem("United Arab Emirates", "United Arab Emirates"))
            ddlCountries.Items.Add(New ListItem("United Kingdom", "United Kingdom"))
            ddlCountries.Items.Add(New ListItem("Uruguay", "Uruguay"))
            ddlCountries.Items.Add(New ListItem("Uzbekistan", "Uzbekistan"))
            ddlCountries.Items.Add(New ListItem("Vanuatu", "Vanuatu"))
            ddlCountries.Items.Add(New ListItem("Vatican City", "Vatican City"))
            ddlCountries.Items.Add(New ListItem("Venezuela", "Venezuela"))
            ddlCountries.Items.Add(New ListItem("Vietnam", "Vietnam"))
            ddlCountries.Items.Add(New ListItem("Western Sahara", "Western Sahara"))
            ddlCountries.Items.Add(New ListItem("Yemen", "Yemen"))
            ddlCountries.Items.Add(New ListItem("Zambia", "Zambia"))
            ddlCountries.Items.Add(New ListItem("Zimbabwe", "Zimbabwe"))

            'select the given country
            ddlCountries.SelectedValue = strSelected
        End Sub

        ''' <summary>
        ''' Adds items for each month to the given drop-down list
        ''' </summary>
        ''' <param name="ddlMonths">The drop-down list to add the items to</param>
        ''' <param name="strSelected">The month to be selected by default</param>
        ''' <param name="blnSelect">Boolean value indicating is a "Select an Item" item should be included</param>
        ''' <param name="blnLongName">Boolean value indicating if full month names should be displayed</param>
        Public Sub MonthDDLInit(ByRef ddlMonths As DropDownList, ByVal strSelected As String, _
            Optional ByVal blnSelect As Boolean = True, Optional ByVal blnLongName As Boolean = False)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Adds items for each month to the given drop-down list
            ' *****************************************************************************
            ddlMonths.Items.Clear()

            'add a default null value
            If blnSelect Then
                DDLInsert_SelectItem(ddlMonths, , Not blnLongName)
            End If

            'add all of the months
            With ddlMonths
                If blnLongName Then
                    .Items.Add(New ListItem("January", "1"))
                    .Items.Add(New ListItem("February", "2"))
                    .Items.Add(New ListItem("March", "3"))
                    .Items.Add(New ListItem("April", "4"))
                    .Items.Add(New ListItem("May", "5"))
                    .Items.Add(New ListItem("June", "6"))
                    .Items.Add(New ListItem("July", "7"))
                    .Items.Add(New ListItem("August", "8"))
                    .Items.Add(New ListItem("September", "9"))
                    .Items.Add(New ListItem("October", "10"))
                    .Items.Add(New ListItem("November", "11"))
                    .Items.Add(New ListItem("December", "12"))
                Else
                    .Items.Add(New ListItem("Jan", "1"))
                    .Items.Add(New ListItem("Feb", "2"))
                    .Items.Add(New ListItem("Mar", "3"))
                    .Items.Add(New ListItem("Apr", "4"))
                    .Items.Add(New ListItem("May", "5"))
                    .Items.Add(New ListItem("Jun", "6"))
                    .Items.Add(New ListItem("Jul", "7"))
                    .Items.Add(New ListItem("Aug", "8"))
                    .Items.Add(New ListItem("Sep", "9"))
                    .Items.Add(New ListItem("Oct", "10"))
                    .Items.Add(New ListItem("Nov", "11"))
                    .Items.Add(New ListItem("Dec", "12"))
                End If
            End With

            'select the given state
            ddlMonths.SelectedValue = strSelected
        End Sub

        ''' <summary>
        ''' Adds items for each day of the week to the given drop-down list
        ''' </summary>
        ''' <param name="ddlDays">The drop-down list to add the items to</param>
        ''' <param name="strSelected">The day to be selected by default</param>
        ''' <param name="blnSelect">Boolean value indicating is a "Select an Item" item should be included</param>
        ''' <param name="blnLongName">Boolean value indicating if full day names should be displayed</param>
        Public Sub DayDDLInit(ByRef ddlDays As DropDownList, ByVal strSelected As String, _
            Optional ByVal blnSelect As Boolean = True, Optional ByVal blnLongName As Boolean = False)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2011.04.17
            ' Description:  Adds items for each day of the week to the given drop-down list
            ' *****************************************************************************
            ddlDays.Items.Clear()

            'add a default null value
            If blnSelect Then
                DDLInsert_SelectItem(ddlDays, , Not blnLongName)
            End If

            'add all of the days
            With ddlDays
                If blnLongName Then
                    .Items.Add(New ListItem("Sunday", "1"))
                    .Items.Add(New ListItem("Monday", "2"))
                    .Items.Add(New ListItem("Tuesday", "3"))
                    .Items.Add(New ListItem("Wednesday", "4"))
                    .Items.Add(New ListItem("Thursday", "5"))
                    .Items.Add(New ListItem("Friday", "6"))
                    .Items.Add(New ListItem("Saturday", "7"))
                Else
                    .Items.Add(New ListItem("Sun", "1"))
                    .Items.Add(New ListItem("Mon", "2"))
                    .Items.Add(New ListItem("Tue", "3"))
                    .Items.Add(New ListItem("Wed", "4"))
                    .Items.Add(New ListItem("Thu", "5"))
                    .Items.Add(New ListItem("Fri", "6"))
                    .Items.Add(New ListItem("Sat", "7"))
                End If
            End With

            'select the given state
            ddlDays.SelectedValue = strSelected
        End Sub

        ''' <summary>
        ''' Adds items for each integer within a given range to a drop-down list
        ''' </summary>
        ''' <param name="ddlNumeric">The drop-down list to add the items to</param>
        ''' <param name="strSelected">The country to be selected by default</param>
        ''' <param name="intMin">The minimum number to be added</param>
        ''' <param name="intMax">The maximum number to be added</param>
        ''' <param name="blnAscending">Boolean value indicating if the numbers should appear in acending or decending order</param>
        ''' <param name="blnSelect">Boolean value indicating is a "Select an Item" item should be included</param>
        ''' <param name="blnShortSelect">Boolean value indicating if the "Select an Item" item should be shortened to just "--"</param>
        Public Sub NumericDDLInit(ByRef ddlNumeric As DropDownList, ByVal strSelected As String, _
            ByVal intMin As Integer, ByVal intMax As Integer, Optional ByVal blnAscending As Boolean = True, _
            Optional ByVal blnSelect As Boolean = True, Optional ByVal blnShortSelect As Boolean = False)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.11.15
            ' Description:  Adds items for each integer within a given range to a drop-down list
            ' *****************************************************************************
            Dim i As Integer

            'check if the minimum number is smaller than the maximum number
            If intMin > intMax Then
                Throw New Exception("Minimum number must be smaller than the maximum number.")
            End If

            'clear out any existing items
            ddlNumeric.Items.Clear()

            'add a default null value
            If blnSelect Then
                DDLInsert_SelectItem(ddlNumeric, blnShortSelect)
            End If

            'add the numbers to the list
            If blnAscending Then
                For i = intMin To intMax
                    AddDDLItem(ddlNumeric, i, i)
                Next i
            Else
                For i = intMax To intMin Step -1
                    AddDDLItem(ddlNumeric, i, i)
                Next i
            End If

            'select the given number
            SelectDDL(ddlNumeric, strSelected)
        End Sub

        ''' <summary>
        ''' Adds items for the time of day based on the passed in options
        ''' </summary>
        ''' <param name="ddlTime">The dropdown list to add items to</param>
        ''' <param name="bln24Hour">Boolean value indicating if 24 hour time should be used</param>
        ''' <param name="blnAMPM">Boolean value indicating if AM/PM should be included in the items</param>
        ''' <param name="intInterval">The number of minutes that should elapse between items</param>
        Public Sub TimeDDLInit(ByRef ddlTime As DropDownList, Optional ByVal bln24Hour As Boolean = False, _
                               Optional ByVal blnAMPM As Boolean = False, Optional ByVal intInterval As Integer = 15)
            ' *****************************************************************************
            ' Author:       Jeff Richmond
            ' Create date:  2010.04.04
            ' Description:  Adds items for the time of day based on the passed in options
            ' *****************************************************************************
            Dim h, m As Integer
            Dim intMinHour As Integer = 1
            Dim intMaxHour As Integer = 12
            Dim strAMPM As String = ""

            'add 0 hour items for 24 hour time
            If bln24Hour Then
                intMaxHour = 23
                blnAMPM = False

                For m = 0 To 59 Step intInterval
                    AddDDLItem(ddlTime, "0:" & Formatting.PadLeft(m, 2), h & ":" & Formatting.PadLeft(m, 2))
                Next m
            End If

            'add 12am hour for AM/PM time
            If blnAMPM Then
                strAMPM = " AM"

                For m = 0 To 59 Step intInterval
                    AddDDLItem(ddlTime, "12:" & Formatting.PadLeft(m, 2) & strAMPM, h & ":" & Formatting.PadLeft(m, 2) & strAMPM)
                Next m

                intMaxHour -= 1
            End If

            'add the normal hours and minutes
            For h = intMinHour To intMaxHour
                For m = 0 To 59 Step intInterval
                    AddDDLItem(ddlTime, h & ":" & Formatting.PadLeft(m, 2) & strAMPM, h & ":" & Formatting.PadLeft(m, 2) & strAMPM)
                Next m
            Next h

            'add the PM hours for AM/PM time
            If blnAMPM Then
                strAMPM = " PM"

                For m = 0 To 59 Step intInterval
                    AddDDLItem(ddlTime, "12:" & Formatting.PadLeft(m, 2) & strAMPM, h & ":" & Formatting.PadLeft(m, 2) & strAMPM)
                Next m

                For h = intMinHour To intMaxHour
                    For m = 0 To 59 Step intInterval
                        AddDDLItem(ddlTime, h & ":" & Formatting.PadLeft(m, 2) & strAMPM, h & ":" & Formatting.PadLeft(m, 2) & strAMPM)
                    Next m
                Next h
            End If
        End Sub

    End Module

End Namespace