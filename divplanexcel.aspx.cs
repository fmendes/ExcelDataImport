using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;
using System.Data.Odbc;
using System.Diagnostics;
using ExcelLibrary.SpreadSheet;

public partial class entity : System.Web.UI.Page
{
    String strDataSource    = System.Configuration.ConfigurationManager.AppSettings["AppRootPath"].ToString(),
           strLoginId = "",
           strPassword = "";

    string strOriginalFile = "";
    string strTempXLSFile =
     System.Configuration.ConfigurationManager.AppSettings["TempXLSFile"].ToString();
    string EmptyTerm = "* EMPTY *";

    #region Default Page Initialization
    protected void Page_Load(object sender, EventArgs e)
    {
        // load up session variables
        if (!LoadUpSession())
            return;

        lblError.Text = "";
        lblError.ForeColor = System.Drawing.Color.Red;
        lblError.Visible = false;
    }

    private bool LoadUpSession()
    {
        // load up variables from session
        if (Session.IsNewSession || Session["LogIn"] == null)
        {
            // navigate to the default page if the session is invalid/expired
            string strURL = System.Configuration.ConfigurationManager.AppSettings["DefaultPage"].ToString();
            Response.RedirectLocation = strURL;
            Response.Redirect(strURL);
            return false;
        }

        strLoginId = Session["LogIn"].ToString();
        strPassword = Session["Password"].ToString();
        return true;
    }
    #endregion

	protected void btnUpload_Click(object sender, EventArgs e)
	{
		// clear any error messages
		lblError.Text = "";

		//// get the file name
		//strOriginalFile = objUpload.FileName;

		// delete any previously existing file
		try
		{
			System.IO.File.Delete(strTempXLSFile);
		}
		catch (Exception)
		{ }

		// save the file in the Temp directory with another name
		objUpload.SaveAs(strTempXLSFile);

		// load dropdownlist with the sheet names
		ddlstSheet.Items.Clear();
		int iIndex = 0;

		Workbook objWB = Workbook.Load(strTempXLSFile);
		foreach (Worksheet objWS in objWB.Worksheets)
		{
			ddlstSheet.Items.Add(new ListItem(objWS.Name, iIndex.ToString()));
			iIndex++;
		}
	}

	protected void btnExtract_Click(object sender, EventArgs e)
	{
		string strSheetNbr = ddlstSheet.SelectedValue;
		int iSheetNbr = Convert.ToInt32(strSheetNbr);

		// open the worksheet and extract data from it
		Workbook objWB = Workbook.Load( strTempXLSFile );
		Worksheet objWS = objWB.Worksheets[ iSheetNbr ];

		ProcessExcelWorksheet(objWS);

		btnSave.Enabled = true;

		RetrieveObjectiveOwner();
	}

	private void ProcessExcelWorksheet(Worksheet objWS)
	{
		if (objWS.Name.Equals("INSTRUCTIONS"))
			return;

		// this Excel library indexes cells starting from zero (0 = 1st row/col)
		string strCell = GetCellValue( objWS, 2, 1);
		if ( ! strCell.StartsWith( "Goal / Objective" ) )
			return;

		// load GOAL
		DataTable objTbl = new DataTable();
		objTbl.Columns.Add("GoalName", typeof(string));
		objTbl.Columns.Add("Description", typeof(string));
		DataRow objRow = objTbl.NewRow();
		string strGoal = GetCellValue( objWS,  4, 1 );
		strGoal = strGoal.Replace("oal ", "");
		objRow["GoalName"] = strGoal;
		objRow["Description"] = GetCellValue( objWS,  4, 2 );
		objTbl.Rows.Add(objRow);
		gvGoal.DataSource = objTbl;
		gvGoal.DataBind();

		// load OBJECTIVE
		objTbl = new DataTable();
		objTbl.Columns.Add("ObjectiveName", typeof(string));
		objTbl.Columns.Add("Description", typeof(string));
		objTbl.Columns.Add("Guidance", typeof(string));
		objRow = objTbl.NewRow();
		string strObjective = GetCellValue( objWS,  5, 1 );
		strObjective = strObjective.Replace("bjective ", "");
		objRow["ObjectiveName"] = strObjective;
		objRow["Description"] = GetCellValue( objWS,  5, 2 );
		objRow["Guidance"] = GetCellValue( objWS,  5, 6 );
		objTbl.Rows.Add(objRow);
		gvObjective.DataSource = objTbl;
		gvObjective.DataBind();

		// load DRIVERS
		strCell = GetCellValue( objWS,  8, 0 );
		if (strCell.StartsWith("Driver/ Key Performance Indicator"))
		{
			objTbl = new DataTable();
			objTbl.Columns.Add("Description", typeof(string));
			objTbl.Columns.Add("StrategicIndicator_ID", typeof(int));
			objTbl.Columns.Add("Measure", typeof(string));
			objTbl.Columns.Add("Frequency", typeof(string));
			objTbl.Columns.Add("Measure_ID", typeof(string));
			objTbl.Columns.Add("Frequency_ID", typeof(string));
			objTbl.Columns.Add("Baseline", typeof(string));
			objTbl.Columns.Add("Target", typeof(string));
			objTbl.Columns.Add("Selected", typeof(bool));

			try
			{
				int iIndex = 10;
				string strDriver = GetCellValue( objWS, iIndex,  0 );

				// added Tactical Plan and Tactic Number because Gretel's spreadsheet was missing Financial Return:
				while (!strDriver.Equals("Financial Return:")
					&& !strDriver.Equals("Tactical Plan")
					&& !strDriver.Equals("Tactic Number")
					&& iIndex < 100 )	// <- to prevent infinite loops
				{
					if (strDriver.Equals(""))
					{
						iIndex++;
						strDriver = GetCellValue( objWS, iIndex,  0);
						continue;
					}

					objRow = objTbl.NewRow();
					objRow["Selected"] = true;
					objRow["Description"] = CapitalizeInitial(strDriver);
					objRow["StrategicIndicator_ID"] = 0;

					// convert description to primary key
					string strMeasure = GetCellValue( objWS, iIndex,  6);
					int iMeasureId = GetMeasureFromDescription(strMeasure);
					objRow["Measure_ID"] = iMeasureId;
					if (iMeasureId == 0)
					{
						objRow["Measure"] = String.Format("Could not find '{0}'", strMeasure);
						objRow["Selected"] = false;
					}
					else
						objRow["Measure"] = "";

					// convert description to primary key
					string strFrequency = GetCellValue( objWS, iIndex,  8);
					strFrequency = CapitalizeInitial(strFrequency);
					int iFrequencyId = GetFrequencyFromDescription(strFrequency);
					objRow["Frequency_ID"] = iFrequencyId;
					if (iFrequencyId == 0)
					{
						objRow["Frequency"] = String.Format("Could not find '{0}'", strFrequency);
						objRow["Selected"] = false;
					}
					else
						objRow["Frequency"] = "";

					string strBaseline = GetCellValue( objWS, iIndex,  9);
					if (strMeasure.IndexOf("Percent") >= 0)
						strBaseline = ConvertPercentage(strBaseline);
					objRow["Baseline"] = strBaseline;

					string strTarget = GetCellValue( objWS, iIndex,  10);
					if (strMeasure.IndexOf("Percent") >= 0)
						strTarget = ConvertPercentage(strTarget);
					objRow["Target"] = strTarget;

					objTbl.Rows.Add(objRow);

					iIndex++;
					strDriver = GetCellValue( objWS, iIndex, 0 );
				}
				gvDriver.DataSource = objTbl;
				gvDriver.DataBind();
			}
			catch (Exception excpt)
			{
				throw excpt;
			}
		}

		try
		{
			// load TACTICS
			int iTacticIndex = FindStringInColumn( objWS, 0, "Tactic Number" );
			if (iTacticIndex > 0)
			{
				objTbl = new DataTable();
				objTbl.Columns.Add("Tactic", typeof(string));
				objTbl.Columns.Add("Description", typeof(string));
				objTbl.Columns.Add("TargetDate", typeof(string));
				objTbl.Columns.Add("Status_ID", typeof(int));
				objTbl.Columns.Add("Completion", typeof(string));
				objTbl.Columns.Add("UpdatePeriod", typeof(string));
				objTbl.Columns.Add("Selected", typeof(bool));

				int iIndex = iTacticIndex + 1;
				string strDescription = GetCellValue( objWS, iIndex,  1 );
				while (!strDescription.Equals(""))
				{
					objRow = objTbl.NewRow();
					objRow["Selected"] = true;
					objRow["Tactic"] = GetCellValue( objWS, iIndex,  0 ).Replace("actic ", "");
					objRow["Description"] = strDescription;

					string strDate = GetCellValue( objWS, iIndex,  6 );
					if (strDate.Equals("Monthly") || strDate.Equals("Ongoing"))
					{
						// default to last day of current month
						objRow["TargetDate"] = DateTime.Today.AddMonths( 1 ).AddDays( - DateTime.Today.AddMonths( 1 ).Day ).ToShortDateString();
					}
					else
					{
						strDate = ConvertExcelDate(strDate);
						if (strDate.Equals(""))
							objRow["TargetDate"] = null;
						else
							objRow["TargetDate"] = strDate;
					}

					// default to target
					objRow["UpdatePeriod"] = "T";
					objRow["Status_ID"] = 1;
					objRow["Completion"] = GetCellValue(objWS, iIndex, 7);

					objTbl.Rows.Add(objRow);

					iIndex++;
					strDescription = GetCellValue( objWS, iIndex,  1);
				}
				gvTactic.DataSource = objTbl;
				gvTactic.DataBind();

			}
		}
		catch (Exception excpt)
		{
			throw excpt;
		}



	}

	private string GetCellValue( Worksheet objWS, int iIndex, int iCol )
	{
		if (objWS.Cells[ iIndex, iCol ].Value != null)
			return objWS.Cells[ iIndex, iCol ].Value.ToString();

		return "";
	}


    private static string CapitalizeInitial(string strValue)
    {
        if (strValue.Length == 0)
            return strValue;

        return String.Concat( strValue.Substring( 0, 1 ).ToUpper(), 
                                strValue.Substring( 1 ) );
    }

    private static string ConvertPercentage(string strValue)
    {
        Decimal decValue = 0;
        if (Decimal.TryParse(strValue, out decValue))
            if( decValue < 2 )
                strValue = String.Format("{0:f}%", decValue * 100);

        return strValue;
    }

    private static string ConvertExcelDate(string strDate)
    {
        int iDate = 0;
        if (!int.TryParse(strDate, out iDate))
        {
            string strConvertedDate = strDate.Replace("st.", ",");
            strConvertedDate = strConvertedDate.Replace("nd.", ",");
            strConvertedDate = strConvertedDate.Replace("rd.", ",");
            strConvertedDate = strConvertedDate.Replace("th.", ",");
            if (!int.TryParse(strConvertedDate, out iDate))
            {
                DateTime dtConvertedDate = new DateTime();

                if( DateTime.TryParse( strConvertedDate, out dtConvertedDate ) )
                    return dtConvertedDate.ToShortDateString();

                return "";
            }
        }
        DateTime dtDate = new DateTime(1900, 1, 1);
        dtDate = dtDate.AddDays(iDate - 2);
        return dtDate.ToShortDateString();
    }

    private int FindStringInColumn( Worksheet objWS, int iCol, string strText )
    {
        int iIndex = 1;
		string strValue = GetCellValue( objWS, iIndex,  iCol);

		while (iIndex < objWS.Cells.LastRowIndex )
        {
            if (strValue.StartsWith(strText))
                return iIndex;

            iIndex++;
			strValue = GetCellValue( objWS, iIndex,  iCol );
        }

        return -1;
    }

    private void ParseStrategicIndicators(DataTable objTbl, string strStratInd)
    {
        string[] strIndicatorsArray = strStratInd.Split(
                        new char[] { ';', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string strIndicator in strIndicatorsArray)
        {
            //if (strIndicator.Trim().Equals(""))
                //continue;
            DataRow objRow = objTbl.NewRow();
            objRow["StrategicIndicator"] = strIndicator;
            objTbl.Rows.Add(objRow);
        }
    }

    private static string ConvertToNumber(string strValue)
    {
        strValue = strValue.Replace("%", "");
        return strValue;
    }

    private static bool StringsSimilarityConfirmed(string strFirst, string strSecond )
    {
        // the comparison will be made ignoring case, periods and spaces
        strFirst  = strFirst.Replace( ".", "" ).Replace( " ", "" );
        strSecond = strSecond.Replace( ".", "" ).Replace( " ", "" );

        return ! strFirst.StartsWith(strSecond, StringComparison.CurrentCultureIgnoreCase)
                    && ! strSecond.StartsWith(strFirst, StringComparison.CurrentCultureIgnoreCase);
    }
}