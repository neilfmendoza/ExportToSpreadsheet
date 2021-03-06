﻿using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExportToSpreadsheet
{
    public partial class Form1 : Form
    {
        protected static string oraDB = string.Empty;
        protected static OracleConnection oraConn = new OracleConnection();
        protected static DataTable dtTRIDS = new DataTable();

        Stopwatch timer = new Stopwatch();
        TimeSpan timespan = new TimeSpan();

        public Form1()
        {
            InitializeComponent();
        }

        protected void InitializeOracleConnection()
        {
            oraDB = "Data Source=(DESCRIPTION="
             + "(ADDRESS=(PROTOCOL=TCP)(HOST=db77prd189.cqxvyjemdhyr.us-west-2.rds.amazonaws.com)(PORT=1521))"
             + "(CONNECT_DATA=(SERVICE_NAME=osPRD2)));"
             + "User Id=CUSTOMER_PRD2_READ_ONLY;Password=tkECuaxB5y4myTzirBu93rF5A;";
            oraConn = new OracleConnection(oraDB);
            oraConn.Open();
        }

        protected void DisposeOracleConnection()
        {
            oraConn.Dispose();
        }

        private void btnExort_Click(object sender, EventArgs e)
        {
            CustomSetText("Exporting");
            btnExort.Enabled = false;

            InitializeOracleConnection();
            OracleCommand oraCmd = new OracleCommand();
            oraCmd.Connection = oraConn;
            oraCmd.CommandType = CommandType.Text;

            List<string> lstTRIDs = new List<string>() { "TR396QE95" };//, "TR128WY17" };//GetTRIDs(oraCmd);//

            DataTable dtAnalyteInfo = GetAnalyteInformation(oraCmd, lstTRIDs);

            new Thread(delegate ()
            {
                DataTable dtWaterInfo = GetWaterInformation(oraCmd, dtAnalyteInfo);

                DataTable dTable = dtWaterInfo;
                StreamWriter spreadWriter = new StreamWriter(".\\SampleExportFile.csv", false, Encoding.UTF8);

                timer = Stopwatch.StartNew();
                using (spreadWriter)
                {
                    foreach (DataColumn col in dTable.Columns)
                    {
                        string colName = col.ColumnName;
                        spreadWriter.Write(colName + ",");
                    }
                    spreadWriter.WriteLine();

                    foreach (DataRow row in dTable.Rows)
                    {
                        for (int rowctr = 0; rowctr < row.ItemArray.Length; rowctr++)
                        {
                            string strField = row[rowctr].ToString();
                            strField = (strField.Contains(",")) ? "\"" + strField + "\"" : strField;
                            spreadWriter.Write(strField + ",");
                        }
                        spreadWriter.WriteLine();
                    }
                    spreadWriter.Close();
                    timer.Stop();
                    timespan = timer.Elapsed;
                    CustomSetText(string.Format("Writing to spreadsheet done in {0:00}:{1:00}:{2:00}", timespan.Minutes, timespan.Seconds, timespan.Milliseconds / 10));
                }
            }).Start();

            btnExort.Enabled = true;
        }

        private List<string> GetTRIDs(OracleCommand cmd)
        {
            timer = Stopwatch.StartNew();
            cmd.CommandText = @"SELECT * FROM (SELECT DISTINCT OSADMIN_PRD2.OSUSR_HGM_LABANAL1.TRID FROM OSADMIN_PRD2.OSUSR_HGM_LABANAL1 WHERE OSADMIN_PRD2.OSUSR_HGM_LABANAL1.TRID IS NOT NULL AND OSADMIN_PRD2.OSUSR_HGM_LABANAL1.ISACTIVE IN (0,1) ORDER BY OSADMIN_PRD2.OSUSR_HGM_LABANAL1.TRID ASC) TBLTEMP WHERE rownum >= 1 AND rownum <= 100";
            DataSet dataSet = new DataSet();
            using (OracleDataAdapter oraDataAdapter = new OracleDataAdapter())
            {
                oraDataAdapter.SelectCommand = cmd;
                oraDataAdapter.Fill(dataSet);
            }

            timer.Stop();
            timespan = timer.Elapsed;
            this.CustomSetText(string.Format("TRID query done in {0:00}:{1:00}:{2:00}", timespan.Minutes, timespan.Seconds, timespan.Milliseconds / 10) );

            dtTRIDS = dataSet.Tables[0];

            return dataSet.Tables[0].AsEnumerable().Select(t => t[0].ToString()).ToList();
        }
        
        private DataTable GetAnalyteInformation(OracleCommand cmd, List<string> trids)
        {
            timer = Stopwatch.StartNew();
            List<string> lstTridChunks = new List<string>() { };
            string tridSet = string.Empty;
            string tridWhereClause = string.Empty;
            int intChunkSize = 1000;
            
            int intTridLength = trids.Count;
            for (int ctr = 0; ctr < intTridLength; ctr += intChunkSize)
            {
                lstTridChunks.Add("'" + string.Join("','", trids.GetRange(ctr, Math.Min(intChunkSize, trids.Count - ctr))) + "'");
            }

            int chunkCtr = 0;
            foreach (string chunkedTrids in lstTridChunks)
            {
                chunkCtr++;
                tridWhereClause += string.Format("OSADMIN_PRD2.OSUSR_HGM_LABANAL1.TRID IN ({0})", chunkedTrids);
                tridWhereClause += (chunkCtr < lstTridChunks.Count) ? " OR " : "";
            }
            
            cmd.CommandText = string.Format(@"SELECT DISTINCT ANALYTEDATA.TRID,ANALYTEDATA.AnalyteName, ANALYTEDATA.ANALYTEID, ANALYTEDATA.ANALYTETYPEID FROM ( 
                                            SELECT
                                            OSADMIN_PRD2.OSUSR_HGM_LABANAL1.TRID,
                                            CASE 
	                                            WHEN OSADMIN_PRD2.OSUSR_35O_LABORAT1.ANALYTETYPEID = 1 THEN OSADMIN_PRD2.OSUSR_35O_CHEMICA1.NAME
	                                            WHEN OSADMIN_PRD2.OSUSR_35O_LABORAT1.ANALYTETYPEID = 3 THEN OSADMIN_PRD2.OSUSR_35O_PARAMETE.NAME
                                            END AS AnalyteName, OSADMIN_PRD2.OSUSR_35O_LABORAT1.ANALYTEID, OSADMIN_PRD2.OSUSR_35O_LABORAT1.ANALYTETYPEID
                                            FROM OSADMIN_PRD2.OSUSR_HGM_LABANAL1
	                                            INNER JOIN OSADMIN_PRD2.OSUSR_35O_LABORAT3
                                            ON OSADMIN_PRD2.OSUSR_HGM_LABANAL1.LABANALYSISID = OSADMIN_PRD2.OSUSR_35O_LABORAT3.ID
	                                            LEFT JOIN OSADMIN_PRD2.OSUSR_35O_LABORAT1
                                            ON OSADMIN_PRD2.OSUSR_35O_LABORAT3.ID = OSADMIN_PRD2.OSUSR_35O_LABORAT1.LABORATORYANALYSISID
	                                            LEFT JOIN OSADMIN_PRD2.OSUSR_35O_CHEMICA1
                                            ON OSADMIN_PRD2.OSUSR_35O_LABORAT1.ANALYTEID = OSADMIN_PRD2.OSUSR_35O_CHEMICA1.ID
	                                            AND OSADMIN_PRD2.OSUSR_35O_LABORAT1.ANALYTETYPEID = 1
	                                            LEFT JOIN OSADMIN_PRD2.OSUSR_35O_PARAMETE
                                            ON OSADMIN_PRD2.OSUSR_35O_LABORAT1.ANALYTEID = OSADMIN_PRD2.OSUSR_35O_PARAMETE.ID
	                                            AND OSADMIN_PRD2.OSUSR_35O_LABORAT1.ANALYTETYPEID = 3
                                WHERE {0}
                                ORDER BY OSADMIN_PRD2.OSUSR_HGM_LABANAL1.TRID, OSADMIN_PRD2.OSUSR_35O_LABORAT1.ANALYTEID ASC
                                ) ANALYTEDATA
                                GROUP BY ANALYTEDATA.TRID,ANALYTEDATA.AnalyteName,ANALYTEDATA.ANALYTEID, ANALYTEDATA.ANALYTETYPEID
                                ORDER BY ANALYTEDATA.TRID, ANALYTEDATA.AnalyteName", tridWhereClause);


            DataSet dataSet = new DataSet();
            using (OracleDataAdapter oraDataAdapter = new OracleDataAdapter())
            {
                oraDataAdapter.SelectCommand = cmd;
                oraDataAdapter.Fill(dataSet);
            }

            timer.Stop();
            timespan = timer.Elapsed;

            this.CustomSetText(string.Format("Analyte Info query done in {0:00}:{1:00}:{2:00}", timespan.Minutes, timespan.Seconds, timespan.Milliseconds / 10) );

            return dataSet.Tables[0];
        }

        delegate void ProgressBarDelegate(int value);

        private void InvokeProgBarTick(int value)
        {
            if (progBar.InvokeRequired)
            {
                ProgressBarDelegate p = new ProgressBarDelegate(InvokeProgBarTick);
                this.Invoke(p, new object[] { value });
            }
            else
            {
                progBar.Value = value;
            }
        }

        delegate void ProgressBarMax(int value);

        private void InvokeProgBarMax(int value)
        {
            if (progBar.InvokeRequired)
            {
                ProgressBarMax p = new ProgressBarMax(InvokeProgBarMax);
                this.Invoke(p, new object[] { value });
            }
            else
            {
                progBar.Maximum = value;
            }
        }

        delegate void StringArgReturningVoidDelegate(string text);

        private void CustomSetText(string text)
        {
            if (txtState.InvokeRequired)
            {
                StringArgReturningVoidDelegate d = new StringArgReturningVoidDelegate(CustomSetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                txtState.Text += text + Environment.NewLine;
            }

        }
        
        private void CustomSetTextLabel(string text)
        {
            if (lblState.InvokeRequired)
            {
                StringArgReturningVoidDelegate d = new StringArgReturningVoidDelegate(CustomSetTextLabel);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                lblState.Text = text;
            }

        }

        private DataTable GetWaterInformation(OracleCommand cmd, DataTable analyteInfoTable)
        {
            Stopwatch timer = Stopwatch.StartNew();
            DataSet dataSet = new DataSet();
            DataTable dtWaterInfo = analyteInfoTable;
            string strSubCommand = string.Empty;

            List<string> lstProdGUIDs = new List<string>() {  "TREATED", "RAW" , "INCOMING","SLUDGE" };
            foreach (string prodGUID in lstProdGUIDs)
            {
                dtWaterInfo.Columns.Add(string.Format("RESULT{0}", prodGUID), typeof(string));
                dtWaterInfo.Columns.Add(string.Format("ABBREVIATION{0}", prodGUID), typeof(string));
                dtWaterInfo.Columns.Add(string.Format("HASVALUE{0}", prodGUID), typeof(bool));
                dtWaterInfo.Columns.Add(string.Format("SAMPLETYPEID{0}", prodGUID), typeof(int));
            }

            using (OracleDataAdapter oraDataAdapter = new OracleDataAdapter())
            {
                int ctrAnalyte = 0;
                InvokeProgBarMax(analyteInfoTable.Rows.Count);
                
                foreach (DataRow analyteInfo in analyteInfoTable.Rows)
                {
                    strSubCommand = string.Empty;
                    if (!string.IsNullOrEmpty(analyteInfo["ANALYTEID"].ToString()) && !string.IsNullOrEmpty(analyteInfo["ANALYTETYPEID"].ToString()))
                    {
                        //int guidctr = 0;
                        //foreach (string prodGUID in lstProdGUIDs)
                        for (int guidctr = 0; guidctr < lstProdGUIDs.Count; guidctr++)
                        {
                            string prodGUID = lstProdGUIDs[guidctr];
                            strSubCommand += string.Format(@"(SELECT
                                    OSADMIN_PRD2.OSUSR_HGM_LABANAL1.TRID,
	                                OSADMIN_PRD2.OSUSR_35O_LABORAT1.RESULT AS RESULT{0},
	                                OSADMIN_PRD2.OSUSR_VFG_MEASURE2.ABBREVIATION AS ABBREVIATION{0},
	                                CASE
		                                WHEN TRIM(OSADMIN_PRD2.OSUSR_35O_LABORAT1.RESULT) IS NULL THEN 0
		                                ELSE 1
	                                END AS HasValue{0},
	                                OSADMIN_PRD2.OSUSR_35O_LABORAT3.SAMPLETYPEID AS SAMPLETYPEID{0}
                                FROM OSADMIN_PRD2.OSUSR_HGM_LABANAL1
                                INNER JOIN OSADMIN_PRD2.OSUSR_35O_LABORAT3
                                    ON OSADMIN_PRD2.OSUSR_HGM_LABANAL1.LABANALYSISID = OSADMIN_PRD2.OSUSR_35O_LABORAT3.ID
                                INNER JOIN OSADMIN_PRD2.OSUSR_35O_LABORAT1
                                    ON OSADMIN_PRD2.OSUSR_35O_LABORAT1.LABORATORYANALYSISID = OSADMIN_PRD2.OSUSR_35O_LABORAT3.ID
                                LEFT JOIN OSADMIN_PRD2.OSUSR_VFG_MEASURE2
                                    ON OSADMIN_PRD2.OSUSR_35O_LABORAT1.UNITID = OSADMIN_PRD2.OSUSR_VFG_MEASURE2.ID
                                WHERE OSADMIN_PRD2.OSUSR_HGM_LABANAL1.TRID = '{1}'
                                AND OSADMIN_PRD2.OSUSR_35O_LABORAT1.ANALYTETYPEID = {2}
                                AND OSADMIN_PRD2.OSUSR_35O_LABORAT1.ANALYTEID = {3}
                                AND OSADMIN_PRD2.OSUSR_35O_LABORAT3.TYPEID = 2
                                AND OSADMIN_PRD2.OSUSR_35O_LABORAT3.PRODUCTGUID = 'WW-{0}') TBL{0}", prodGUID, analyteInfo["TRID"], analyteInfo["ANALYTETYPEID"], analyteInfo["ANALYTEID"]);
                            if (guidctr == 0)
                            {
                                strSubCommand += " LEFT JOIN ";
                            }
                            else if (guidctr > 0 && guidctr == lstProdGUIDs.Count - 1)
                            {
                                strSubCommand += string.Format(" ON TBL{0}.TRID = TBL{1}.TRID ", lstProdGUIDs[guidctr - 1], prodGUID);
                            }
                            else if (guidctr > 0)
                            {
                                strSubCommand += string.Format(" ON TBL{0}.TRID = TBL{1}.TRID LEFT JOIN ", lstProdGUIDs[guidctr - 1], prodGUID);
                            }
                        }
                        DataTable dtWtrInfo = new DataTable();
                        dataSet.Clear();
                        cmd.CommandText = string.Format("SELECT * FROM ({0})", strSubCommand);
                        oraDataAdapter.SelectCommand = cmd;
                        oraDataAdapter.Fill(dataSet);
                        dtWtrInfo = dataSet.Tables[0];

                        if (dtWtrInfo.Rows.Count > 0)
                        {
                            foreach (string pguid in lstProdGUIDs)
                            {
                                analyteInfo[string.Format("RESULT{0}", pguid)] = dtWtrInfo.Rows[0][string.Format("RESULT{0}", pguid)];
                                analyteInfo[string.Format("ABBREVIATION{0}", pguid)] = dtWtrInfo.Rows[0][string.Format("ABBREVIATION{0}", pguid)];
                                analyteInfo[string.Format("HASVALUE{0}", pguid)] = dtWtrInfo.Rows[0][string.Format("HASVALUE{0}", pguid)];
                                analyteInfo[string.Format("SAMPLETYPEID{0}", pguid)] = dtWtrInfo.Rows[0][string.Format("SAMPLETYPEID{0}", pguid)];
                            }
                        }
                    }
                    ctrAnalyte++;
                    this.CustomSetTextLabel("(" + (ctrAnalyte).ToString() + "/" + analyteInfoTable.Rows.Count + ") analyte info processed for water info.");
                    InvokeProgBarTick(ctrAnalyte);
                }
            }
            timer.Stop();
            timespan = timer.Elapsed;
            this.CustomSetText(string.Format("Water Info query done in {0:00}:{1:00}:{2:00}", timespan.Minutes, timespan.Seconds, timespan.Milliseconds / 10));
            return dtWaterInfo;
        }

        public Task taskAnalyteInfo(OracleCommand oraCmd, List<string> lstTRIDs)
        {
            return Task.Run(() => {
                DataTable dtAnalyteInfo = new DataTable();
                dtAnalyteInfo = GetAnalyteInformation(oraCmd, lstTRIDs);
            });
        }

        private Task DoAllTasks(OracleCommand oraCmd, List<string> lstTRIDs)
        {
            return Task.Run(
                async () =>
                {
                    await Task.WhenAll(taskAnalyteInfo(oraCmd, lstTRIDs));

                    DisposeOracleConnection();
                }
                );
        }
    }
}