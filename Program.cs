using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net.Mail;
using Oracle_Class;
using ExcelClass;
using System.Data;
using System.Text.RegularExpressions;
using System.Threading;
using System.Drawing;
using System.Net.Mime;
using System.Net;
using System.Xml;


namespace x_reports
{
    class Program
    {
        
        string pathWay = @"."; 
        string systemPathWay = ".";

        string fileName = @"\test.log";  // Log file name
        
        string adTableName = "tbl_multi_cs";  // table name where all the query all held      
        string MDYformat = "dd/MMM/yyyy HH:mm:ss"; 
        string failedReports = "";
        private static readonly object _lockNow = new object(); //use to lock the log file so the separate thread can write to the log file
        double version = 1.0;

        static void Main(string[] args) 
        {        

            Program a = new Program();
          
                     
           // one off table creation 
            OracleClass orC = new OracleClass();// call oracle class

            int tblRow = 0;         

            bool oneRunNoChange = false;
            string getStandarsql = "*** sql statement";
            string getFailedsql = "*** sql statement";

            // use to collect date & time to use to update the log file
            DateTime datetime = DateTime.Now;
            string justDate = datetime.ToString(a.MDYformat);
            // how we update the log file
            a.updateLog(a.pathWay, a.fileName, justDate, "Process started for the Report mailing process mltg.EXE: version " + a.version + ": Author nasadora", true);
			
           // create virtual tables to hold the reports and information
            System.Data.DataTable cReportTBL = a.cReport();
            System.Data.DataTable cReportTBL1 = a.cReport();
            System.Data.DataTable cReportTBL2 = a.cReport();
            System.Data.DataTable cReportTBL3 = a.cReport();
            System.Data.DataTable ccheckTBL = a.cReport();

         
            ccheckTBL = orC.readOracleToDatabtable(cReportTBL, getFailedsql); // everything in report table has failed in previous run
       
            if (ccheckTBL.Rows.Count > 0) // if found more than 0, validation check required
            {
                int tblError = 0;
                foreach (DataRow row in ccheckTBL.Rows)
                {

                    if (!row[3].ToString().Contains("1999") & (row[3].ToString() != ""))// if the row is not contain 1999 & is not blank then it is error, we need to know to email the team
                    {
                        // mail the error
                        a.mailTemplateImbeddedIcons("Report-ERROR: date in the past - Report ID " + ccheckTBL.Rows[tblError]["RID"].ToString(), "REPORT-ERROR-ERROR: date in the past - Report ID " + ccheckTBL.Rows[tblError]["RID"].ToString(), "Date is in the past on this record", "Update the date and manually check", "*******.com", "**********.com", "", "", "", "");


                    }
                    tblError++;
                }

            }
            ccheckTBL.Clear();// clear the table ready to be reuse

            cReportTBL = orC.readOracleToDatabtable(cReportTBL, getStandarsql); // tbl will now have only the ones we need to check for today
            DataRow[] foundRunNow = cReportTBL.Select("RNOW = 'y'");// to check if there something need to be done immediately
            if (foundRunNow.Length != 0)// if found, action now
            {
                oneRunNoChange = true;
                // set the report to run using the data in foundRunNow only
                int foundNowRow = 0;
                DataTable runNowtbl = a.RNowReport();
                foreach (DataRow row in foundRunNow)
                {
                    runNowtbl.Rows.Add("");
                    runNowtbl.Rows[foundNowRow]["RID"] = row[0].ToString();
                    runNowtbl.Rows[foundNowRow]["MID"] = row[1].ToString();
                    runNowtbl.Rows[foundNowRow]["MSUBJECT"] = row[2].ToString();
                    runNowtbl.Rows[foundNowRow]["RDATE"] = row[3].ToString();
                    runNowtbl.Rows[foundNowRow]["CDATE"] = row[4].ToString();
                    runNowtbl.Rows[foundNowRow]["RNOW"] = row[5].ToString();
                    runNowtbl.Rows[foundNowRow]["SREPORT"] = row[6].ToString();
                    runNowtbl.Rows[foundNowRow]["SDULE"] = row[7].ToString();
                    
                    if (row[8] != DBNull.Value)
                    {
                        runNowtbl.Rows[tblRow]["STAT"] = row[8].ToString();
                    }

                    runNowtbl.Rows[tblRow]["DCREATED"] = row[9].ToString();
                    if (row[11] != DBNull.Value)
                    {
                        runNowtbl.Rows[tblRow]["SYS_COMM"] = row[11].ToString();
                    }
                    foundNowRow++;

                }

                cReportTBL.Clear();  // first clear the customer report tbl
                cReportTBL = runNowtbl.Copy();
                datetime = DateTime.Now;
                string justDateN = datetime.ToString(a.MDYformat);
                a.updateLog(a.pathWay, a.fileName, justDateN, "forced run, Schedule dates will not be updated", false);

            }


            // split the customer report tbl into 3 and start 3 new threads 
            int tbl1 = 0;
            int tbl2 = 0;
            int tbl3 = 0;
            int caseLoopN = 0;//use to determine which table get written to
            
            foreach (DataRow row in cReportTBL.Rows)
            {
                switch (caseLoopN) //if its 0 you updating tbl1, if its 1 you update tbl2, if its 2 you update tbl3
                {
                    case 0:
                        cReportTBL1.Rows.Add("");
                        cReportTBL1.Rows[tbl1]["RID"] = row[0].ToString();
                        cReportTBL1.Rows[tbl1]["MID"] = row[1].ToString();
                        cReportTBL1.Rows[tbl1]["MSUBJECT"] = row[2].ToString();
                        cReportTBL1.Rows[tbl1]["RDATE"] = row[3].ToString();
                        cReportTBL1.Rows[tbl1]["CDATE"] = row[4].ToString();
                        cReportTBL1.Rows[tbl1]["RNOW"] = row[5].ToString();
                        cReportTBL1.Rows[tbl1]["SREPORT"] = row[6].ToString();
                        cReportTBL1.Rows[tbl1]["SDULE"] = row[7].ToString();
                        if (row[8] != DBNull.Value)
                        {
                            cReportTBL1.Rows[tbl1]["STAT"] = row[8].ToString();
                        }

                        cReportTBL1.Rows[tbl1]["DCREATED"] = row[9].ToString();
                        if (row[11] != DBNull.Value)
                        {
                            cReportTBL1.Rows[tbl1]["SYS_COMM"] = row[11].ToString();
                        }

                        tbl1++;
                        caseLoopN++;//adding 1 to force the update to tbl2
                        break;

                    case 1:
                        cReportTBL2.Rows.Add("");
                        cReportTBL2.Rows[tbl1]["RID"] = row[0].ToString();
                        cReportTBL2.Rows[tbl1]["MID"] = row[1].ToString();
                        cReportTBL2.Rows[tbl1]["MSUBJECT"] = row[2].ToString();
                        cReportTBL2.Rows[tbl1]["RDATE"] = row[3].ToString();
                        cReportTBL2.Rows[tbl1]["CDATE"] = row[4].ToString();
                        cReportTBL2.Rows[tbl1]["RNOW"] = row[5].ToString();
                        cReportTBL2.Rows[tbl1]["SREPORT"] = row[6].ToString();
                        cReportTBL2.Rows[tbl1]["SDULE"] = row[7].ToString();
                        if (row[8] != DBNull.Value)
                        {
                            cReportTBL2.Rows[tbl2]["STAT"] = row[8].ToString();
                        }

                        cReportTBL2.Rows[tbl2]["DCREATED"] = row[9].ToString();
                        if (row[11] != DBNull.Value)
                        {
                            cReportTBL2.Rows[tbl2]["SYS_COMM"] = row[11].ToString();
                        }
                        tbl2++;
                        caseLoopN++;//adding 1 to force the update to tbl3
                        break;

                    case 2:
                        cReportTBL3.Rows.Add("");

                        cReportTBL3.Rows[tbl3]["RID"] = row[0].ToString();
                        cReportTBL3.Rows[tbl3]["MID"] = row[1].ToString();
                        cReportTBL3.Rows[tbl3]["MSUBJECT"] = row[2].ToString();
                        cReportTBL3.Rows[tbl3]["RDATE"] = row[3].ToString();
                        cReportTBL3.Rows[tbl3]["CDATE"] = row[4].ToString();
                        cReportTBL3.Rows[tbl3]["RNOW"] = row[5].ToString();
                        cReportTBL3.Rows[tbl3]["SREPORT"] = row[6].ToString();
                        cReportTBL3.Rows[tbl3]["SDULE"] = row[7].ToString();
                        if (row[8] != DBNull.Value)
                        {
                            cReportTBL3.Rows[tbl3]["STAT"] = row[8].ToString();
                        }

                        cReportTBL3.Rows[tbl3]["DATECREATED"] = row[9].ToString();
                        if (row[11] != DBNull.Value)
                        {
                            cReportTBL3.Rows[tbl3]["SYS_COMM"] = row[11].ToString();
                        }
                        tbl3++;
                        caseLoopN=0;//this is to reset, so tbl1 will be return again, its important to reset, otherwise the tbl1 will never start again
                        break;

                }


            }       

            int numOfThreads = 3;//the no. of threads that the system will split into
            WaitHandle[] waitHandles = new WaitHandle[numOfThreads];// use to identify how many defined thread.....
            int Threadloop = 0;

            datetime = DateTime.Now;
            string justDateT = datetime.ToString(a.MDYformat);
            a.updateLog(a.pathWay, a.fileName, justDateT, "Thread started", false);

            for (int i = 0; i < numOfThreads; i++)//each thread need different tbl diff excels file
            {
                var handle = new EventWaitHandle(false, EventResetMode.ManualReset);

                switch (i)
                {
                    case 0:

                        Thread myThread1 = new Thread(new ThreadStart(() =>
                        {
                            a.MultithreadedAction(cReportTBL1, a.systemPathWay + "T1.xlsx", oneRunNoChange);//1st thread action kicks off
                            handle.Set();//the 1st wait use for the system when the 1st for thread has closed off

                        }));
                        myThread1.Start();// the actual separate thread is now running
                        break;

                    case 1:

                        Thread myThread2 = new Thread(new ThreadStart(() =>
                        {
                            a.MultithreadedAction(cReportTBL2, a.systemPathWay + "T2.xlsx", oneRunNoChange);// 2nd thread kicks off
                            handle.Set();// the 2nd wait use for the system when the 2nd for thread has closed off

                        }));
                        myThread2.Start();// the actual separate thread is now running
                        break;

                    case 2:

                        Thread myThread3 = new Thread(new ThreadStart(() =>
                        {
                            a.MultithreadedAction(cReportTBL3, a.systemPathWay + "T3.xlsx", oneRunNoChange);
                            handle.Set();

                        }));
                        myThread3.Start();
                        break;
                }



                waitHandles[Threadloop] = handle;

                Threadloop++;
            }

        
            WaitHandle.WaitAll(waitHandles);//to stop the code to going any further until all the thread has been collected
            // now write the data back using the table report number
          
            datetime = DateTime.Now;
            string ClosingDate = datetime.ToString(a.MDYformat);
            a.updateLog(a.pathWay, a.fileName, ClosingDate, "First run of the report have finished on " + numOfThreads + " of threads", false);

            cReportTBL.Clear();
            cReportTBL1.Clear();
            cReportTBL2.Clear();
            cReportTBL3.Clear();


            if (a.failedReports != "")//if any report failed, we write a msg in a failed report section, so we know that we do the separate run
            {
                datetime = DateTime.Now;
               string justDateN = datetime.ToString(a.MDYformat);
                a.updateLog(a.pathWay, a.fileName, justDateN, "running again on " + a.failedReports, false);
                a.rerunCheck();// rerun to try to pick up the failed report & complete this time
            }

            Environment.Exit(0);// closing off the program

        }

        private void rerunCheck()
        {
           
            DateTime datetime = DateTime.Now;
            string reRunDate = datetime.ToString(MDYformat);
            updateLog(pathWay, fileName, reRunDate, "Re-run of the report started", false);
            OracleClass orC = new OracleClass();
            string[] reportID = Regex.Split(failedReports, "::");//any report that failed last time, we gonna try run the again
            string sqlAppend = string.Empty;
            for (int i = 0; i < reportID.Count(); i++)// this is where we gonna start
            {
                if (reportID[i] != "")
                {
                    if (i == 0)
                    {
                        sqlAppend += " '" + reportID[i] + "'";
                    }
                    else
                    {
                        sqlAppend += " or Report_ID = '" + reportID[i] + "'";
                    }

                }
            }

           

            string getStandarsql = "**** sql statement;




            System.Data.DataTable cReportTBL = cReport();// brand new empty tbl are created
            System.Data.DataTable ccheckTBL = cReport();
            ccheckTBL = orC.readOracleToDatabtable(cReportTBL, getStandarsql);// we go off the cust reports tbl & we collect only the reports that failed on the last run
            MultithreadedAction(ccheckTBL, systemPathWay + "T4.xlsx", false);// 

            datetime = DateTime.Now;
            string ClosingDate = datetime.ToString(MDYformat);
            updateLog(pathWay, fileName, ClosingDate, "Report have finished", false);

            cReportTBL.Clear();

        }

        private void MultithreadedAction(DataTable cReportTBL, string systemPathWay, bool oneRunNoChange)
        {
            Program a = new Program();
            //updateLog(pathWay, fileName, "xxx", systemPathWay, false);
            OracleClass orC = new OracleClass();
            int tblRow = 0;
            System.Data.DataTable oneUpdatePerRun = a.cReport();
            
            foreach (DataRow row in cReportTBL.Rows)// split it per row 
            {
                
                try
                {
                    oneUpdatePerRun.Rows.Add("");
                    DateTime datetime = DateTime.Now;

                    string MDYformat = "dd/MMM/yyyy HH:mm:ss";
                    string justDate = datetime.ToString(MDYformat);

                    string comments = string.Empty;
                    string status = string.Empty;

                    string standardBlurb = "<br>Please find attached your report" +
                    "<br><br>******* " +
                    "***** " +
                    "***** " +
                    "***** " +
                    "<br><br>********* <br><br>";
                    string speedRead = "<br>*****<br><br>";
                    string whatAction = "****<br><br>" +
                        "******";
                    string sqlRestriction = row[6].ToString().ToUpper().Replace(" ", "");

                    try
                    {
                        oneUpdatePerRun.Rows[0]["RNOW"] = "n";
                        oneUpdatePerRun.Rows[0]["DDATE"] = datetime;
                        oneUpdatePerRun.Rows[0]["RID"] = cReportTBL.Rows[tblRow]["REPORT_ID"].ToString();
                        oneUpdatePerRun.Rows[0]["MID"] = cReportTBL.Rows[tblRow]["MAILID"].ToString();

                        oneUpdatePerRun.Rows[0]["MSUBJECT"] = cReportTBL.Rows[tblRow]["MAIL_SUBJECT"].ToString();
                        oneUpdatePerRun.Rows[0]["RDATE"] = cReportTBL.Rows[tblRow]["RUNDATE"].ToString();
                        oneUpdatePerRun.Rows[0]["CDATE"] = cReportTBL.Rows[tblRow]["CERTDATE"].ToString();

                        oneUpdatePerRun.Rows[0]["SREPORT"] = cReportTBL.Rows[tblRow]["SQL_REPORT"].ToString();
                        oneUpdatePerRun.Rows[0]["SDULE"] = cReportTBL.Rows[tblRow]["SCHEDULE"].ToString();


                        if (cReportTBL.Rows[tblRow]["STAT"] != DBNull.Value)
                        {
                            oneUpdatePerRun.Rows[0]["STAT"] = cReportTBL.Rows[tblRow]["STAT"].ToString();
                        }

                        oneUpdatePerRun.Rows[0]["DATECREATED"] = cReportTBL.Rows[tblRow]["DATECREATED"].ToString();
                        if (cReportTBL.Rows[tblRow]["SYS_COMMENTS"] != DBNull.Value)
                        {
                            oneUpdatePerRun.Rows[0]["SYS_COMMENTS"] = cReportTBL.Rows[tblRow]["SYS_COMMENTS"].ToString();
                        }

                    }
                    catch(Exception uD)
                    {
                        a.updateLog(pathWay, fileName, justDate, "report table update issue " + uD.Message.ToString(), false);
                    }
                    string mailNames = row[1].ToString();
                    string subject = row[2].ToString();

                    if (sqlRestriction.Contains("ALTERTABLE") | sqlRestriction.Contains("DROPTABLE")) // sql is trying to drop or rename a table ??? report violation or error
                    {
                        oneUpdatePerRun.Rows[0]["SYS_COMM"] = "report has excluded statments";
                        oneUpdatePerRun.Rows[0]["STAT"] = "rejected";
                        datetime = DateTime.Now;
                        string justDateN = datetime.ToString(MDYformat);
                        a.updateLog(pathWay, fileName, justDateN, "REPORT-ERROR: Excluded statment - R ID " + cReportTBL.Rows[tblRow]["RID"].ToString(), false);
                       
                        a.mailTemplateImbeddedIcons("REPORT-ERROR: Excluded statment - Report ID " + cReportTBL.Rows[tblRow]["RID"].ToString(), "REPORT-ERROR: Excluded statment - Report ID " + cReportTBL.Rows[tblRow]["RID"].ToString(), "Report has been excluded?", "Update if required and manually check", "bxx37070@gsk.com", "bxx37070@gsk.com", "", "", "", "");
                      
                    }
                    else
                    {
                      
                        DataTable cReport = orC.getOracleReport(row[6].ToString());  // to inform no records
                        if (cReport.Rows.Count < 1)
                        {
                            string zeroRecords = "**********";
                            zeroRecords = zeroRecords.Replace("[REPORTID]", row[0].ToString());
                            a.mailTemplateImbeddedIcons(subject, "No records found, please contact the Service Team if in doubt.", "Report has completed with 0 records found.", zeroRecords, mailNames, "********.com", "", "", "", "");
                     
                           
                            // record the report failed to generate 
                            oneUpdatePerRun.Rows[0]["SYS_COMM"] = "0 records found";
                            oneUpdatePerRun.Rows[0]["STAT"] = "Completed";
                            datetime = DateTime.Now;
                            justDate = datetime.ToString(MDYformat);
                            a.updateLog(pathWay, fileName, justDate, "REPORT-ERROR: 0 records found - Report ID: " + cReportTBL.Rows[tblRow]["RID"].ToString(), false);

                        }
                        else
                        {
                            // where we start to create excel file using cReport
                            datetime = DateTime.Now;
                            string justDateN = datetime.ToString(MDYformat);
                            a.updateLog(pathWay, fileName, justDateN, "Running Report ID: " + cReportTBL.Rows[tblRow]["RID"].ToString(), false);
                            ExcelClass1 exC = new ExcelClass1();// collecting excel class
                           
                            // remove duplicate excel file if found
                            try
                            {
                               
                                File.Delete(systemPathWay);// excel file shud have been remove, if it found try to remove it
                               
                            }
                            catch
                            {
                                // not needed as we should not have a file.
                               a.updateLog(pathWay, fileName, justDateN, systemPathWay + " could not be deleted?", false);
                            }
                            // create new file using report info
                            
                            int reportNo = exC.createFileUpdateFromDataTable(cReport, systemPathWay, 1, "Report");// create the excel file using the data from the sql query that we created

      

                            // get the standard report and then personalise it with the scheduled date and report id
                            string sendStandardBlurb = standardBlurb;
                            sendStandardBlurb = sendStandardBlurb.Replace("[SDULE]", row[7].ToString()); // will replace the text [SDULE] with the SDULE value from the table
                            sendStandardBlurb = sendStandardBlurb.Replace("[RID]", row[0].ToString()); // will replace the text [RID] with the report ID value from the table
                            whatAction = whatAction.Replace("[RID]", row[0].ToString());
                            speedRead = speedRead.Replace("[SDULE]", row[7].ToString());
                            
                            // now mail report out
                            bool mailSent = false;
                            try
                            {//mail the excel files to the customers
                                mailSent = a.mailTemplateImbeddedIcons(subject, speedRead, sendStandardBlurb, whatAction, mailNames, "*******.com", "", systemPathWay, "", "");
                            }
                            catch (Exception mailF)
                            {
                                a.updateLog(pathWay, fileName, justDateN, "mail failed to send with message: " + mailF.Message.ToString() + " for RID " + cReportTBL.Rows[tblRow]["RID"].ToString(), false);
                                failedReports += cReportTBL.Rows[tblRow]["RID"].ToString() + "::";

                            }

                            if (mailSent == true)//if we send the email, we gonna update the tbl that we sucessfully mailed
                            {
                                reportNo = reportNo - 1; // removing 1 as we have a header
                                oneUpdatePerRun.Rows[0]["SYS_COMM"] = "Report run and mailed on " + reportNo.ToString() + " records";
                                oneUpdatePerRun.Rows[0]["STAT"] = "Completed";
                                MDYformat = "dd/MMM/yyyy HH:mm:ss";
                                datetime = DateTime.Now;
                                justDateN = datetime.ToString(MDYformat);
                                a.updateLog(pathWay, fileName, justDateN, "Report run and mailed " + reportNo.ToString() + " records on RID " + cReportTBL.Rows[tblRow]["RID"].ToString(), false);

                            }
                            else
                            {//update the tbl that we could not email the customer
                                oneUpdatePerRun.Rows[0]["SYS_COMM"] = "report did not mail";
                                oneUpdatePerRun.Rows[0]["STAT"] = "failed";
                                MDYformat = "dd/MMM/yyyy HH:mm:ss";
                                datetime = DateTime.Now;
                                justDateN = datetime.ToString(MDYformat);
                                a.updateLog(pathWay, fileName, justDateN, "Report did not get mailed for Report ID " + cReportTBL.Rows[tblRow]["RID"].ToString(), false);

                            }

                            // remove excel file if found
                            try
                            {
                                File.Delete(systemPathWay);
                            }
                            catch
                            {
                                a.updateLog(pathWay, fileName, justDateN, systemPathWay + " could not be deleted?", false);


                            }
                        }


                        // set new dates    SDULE
                        if (oneRunNoChange == true)
                        {

                            oneUpdatePerRun.Rows[0]["RUNDATE"] = cReportTBL.Rows[tblRow]["RUNDATE"];

                        }
                        else
                        {

                            switch (cReportTBL.Rows[tblRow]["SDULE"].ToString().ToLower())
                            {// to update the report to the next schedule run
                                case "daily": // RDATE
                                    oneUpdatePerRun.Rows[0]["RDATE"] = datetime.AddDays(1);
                                    break;
                                case "weekly": // RDATE
                                    oneUpdatePerRun.Rows[0]["RDATE"] = datetime.AddDays(7);
                                    break;
                                case "monthly": // RDATE
                                    oneUpdatePerRun.Rows[0]["RDATE"] = datetime.AddMonths(1);
                                    break;

                            }
                        }
                    }
                    tblRow++;
                    orC.UpdateOracleFromDataTable(oneUpdatePerRun);
                    oneUpdatePerRun.Clear();
                }
                catch(Exception failed)
                {
                    string MDYformat = "dd/MMM/yyyy HH:mm:ss";
                    DateTime datetime = DateTime.Now;
                    string justDateN = datetime.ToString(MDYformat);
                    a.updateLog(pathWay, fileName, justDateN, "MultithreadedAction failed on Report " + cReportTBL.Rows[tblRow]["RID"].ToString() + " error message - " + failed.Message.ToString(), false);
                    failedReports += cReportTBL.Rows[tblRow]["RID"].ToString() + "::";
                    oneUpdatePerRun.Clear();
                }


            }
           
           
        }
    
        private void updateLog(string pathWay, string fileName, string justDate, string text, bool st)
        {

            lock (_lockNow)
            {
                try
                {
                    if (st == true)
                    {
                        // new file so should be a new day.
                        // Create a file to write to. 
                        FileStream fcreate = File.Open(pathWay + fileName, FileMode.Create);
                        fcreate.Close();


                        using (TextWriter sw = TextWriter.Synchronized(File.AppendText(pathWay + fileName)))
                        {

                            sw.WriteLine(text + " :: " + justDate);

                        }

                    }
                    else
                    {
                        using (TextWriter sw = TextWriter.Synchronized(File.AppendText(pathWay + fileName)))
                        {
                            sw.WriteLine(Environment.NewLine + text + " :: " + justDate);
                        }
                    }
                }
                catch (Exception e)
                {

                    updateLog(pathWay, fileName, justDate, text + e.Message.ToString(), st);

                }
            }
          

        }

        private DataTable cReport()
        {
            // creates the table of PAMS Priv groups and the memebers

            System.Data.DataTable cReportTBL = new System.Data.DataTable("cReportTBL");
            cReportTBL.Columns.Add("RID", typeof(string));  // ID
            cReportTBL.Columns.Add("MID", typeof(string));  // mail ID of who to get the report 
            cReportTBL.Columns.Add("MSUBJECT", typeof(string));  // the mail subject 
            cReportTBL.Columns.Add("RDATE", typeof(string));  // date this report is to be run 
            cReportTBL.Columns.Add("CDATE", typeof(string));  // last day the report will run 
            cReportTBL.Columns.Add("RNOW", typeof(string));  // last day the report will run 
            cReportTBL.Columns.Add("SREPORT", typeof(string));  // what SQL command does the mail ID's want 
            cReportTBL.Columns.Add("SDULE", typeof(string));  // how offten should the report run  
            cReportTBL.Columns.Add("STAT", typeof(string));  // status of the report 
            cReportTBL.Columns.Add("DCREATED", typeof(string));  // creation date
            cReportTBL.Columns.Add("COMMENTS", typeof(string));  // comments by the code
            cReportTBL.Columns.Add("SYS_COMM", typeof(string));  // comments by the code
            cReportTBL.Columns.Add("DDATE", typeof(string));  // comments by the code 
            return cReportTBL;

        }

        private DataTable RNowReport()
        {
            // creates the table of PAMS Priv groups and the memebers

            System.Data.DataTable RNowRTBL = new System.Data.DataTable("RNowRTBLTBL");
            RunNowReportTBL.Columns.Add("RID", typeof(string));  // ID
            RunNowReportTBL.Columns.Add("MAILID", typeof(string));  // mail ID of who to get the report 
            RunNowReportTBL.Columns.Add("MAIL_SUBJECT", typeof(string));  // the mail subject 
            RunNowReportTBL.Columns.Add("RDATE", typeof(string));  // date this report is to be run 
            RunNowReportTBL.Columns.Add("CDATE", typeof(string));  // last day the report will run 
            RunNowReportTBL.Columns.Add("RNOW", typeof(string));  // last day the report will run 
            RunNowReportTBL.Columns.Add("SREPORT", typeof(string));  // what SQL command does the mail ID's want 
            RunNowReportTBL.Columns.Add("SDULE", typeof(string));  // how offten should the report run  
            RunNowReportTBL.Columns.Add("STAT", typeof(string));  // status of the report 
            RunNowReportTBL.Columns.Add("DCREATED", typeof(string));  // creation date
            RunNowReportTBL.Columns.Add("COMMENTS", typeof(string));  // comments by the code
            RunNowReportTBL.Columns.Add("SYS_COMM", typeof(string));  // comments by the code
            RunNowReportTBL.Columns.Add("DDATE", typeof(string));  // comments by the code 
            return RunNowReportTBL;

        }
     
        private bool mailTemplateImbeddedIcons(string mailSubject, string speedReadText, string WhatDoIneedToKnow, string WhatActionDoINeedToTake, string sendTo, string sendFrom, string ccSent, string attachmentPath, string attachment2, string attachment3)
        {
            bool mailSent = false;
            //sendTo = "tytfyfty";
            //ccSent = "";
            try
            {
                DateTime dateTimeStamp = DateTime.Now;
                string today = dateTimeStamp.ToString(MDYformat);
                string sendmailto = string.Empty;
                using (MailMessage sendMail = new MailMessage())
                {
                    MailAddress fromMail = new MailAddress(sendFrom);
                    string[] sendSplit = Regex.Split(sendTo, "::");
                    foreach (string ID in sendSplit)
                    {
                        if(ID.Contains("@"))
                        {
                            sendMail.To.Add(ID);
                        }
                        else
                        {
                             sendMail.To.Add(ID + "@gsk.com");
                        }
                       

                    }

                    // from the login ID to the tbSendTo.Text
                    if (ccSent != "")
                    {
                        sendMail.CC.Add(ccSent);
                    }
                    System.Net.Mail.Attachment attachment;
                    if (attachmentPath != "")
                    {

                        attachment = new System.Net.Mail.Attachment(attachmentPath);
                        sendMail.Attachments.Add(attachment);

                    }
                    if (attachment2 != "")
                    {

                        attachment = new System.Net.Mail.Attachment(attachment2);
                        sendMail.Attachments.Add(attachment);

                    }
                    if (attachment3 != "")
                    {

                        attachment = new System.Net.Mail.Attachment(attachment3);
                        sendMail.Attachments.Add(attachment);

                    }

                    // make a image link to the logo                    
                    // mailTemplate
                    string bodywithIcon = Properties.Resources.REPORT_Tec_Template; 

                    bodywithIcon = Regex.Replace(bodywithIcon, "!speedReadText!", speedReadText);
                    bodywithIcon = Regex.Replace(bodywithIcon, "!WhatDoIneedToKnow!", WhatDoIneedToKnow);
                    bodywithIcon = Regex.Replace(bodywithIcon, "!WhatActionDoINeedToTake!", WhatActionDoINeedToTake);


                    AlternateView av = AlternateView.CreateAlternateViewFromString(bodywithIcon, null, System.Net.Mime.MediaTypeNames.Text.Html);


                    Bitmap bLogo = new Bitmap(Properties.Resources.GSK_Logo);
                    ImageConverter icLogo = new ImageConverter();
                    Byte[] Logo = (Byte[])icLogo.ConvertTo(bLogo, typeof(Byte[]));
                    MemoryStream LogoImage = new MemoryStream(Logo);
                    LinkedResource LogoImage = new LinkedResource(LogoImage, System.Net.Mime.MediaTypeNames.Image.Jpeg);
                    LogoImage.ContentId = "Logo.jpg";
                    LogoImage.ContentType = new ContentType("image/jpg");
                    av.LinkedResources.Add(LogoImage);

                    Bitmap bYicon = new Bitmap(Properties.Resources.Y_Icon);
                    ImageConverter icYicon = new ImageConverter();
                    Byte[] Yicon = (Byte[])icYicon.ConvertTo(bYicon, typeof(Byte[]));
                    MemoryStream imageY = new MemoryStream(Yicon);
                    LinkedResource YImage = new LinkedResource(imageY, System.Net.Mime.MediaTypeNames.Image.Jpeg);
                    YImage.ContentId = "Y-Icon.jpg";
                    YImage.ContentType = new ContentType("image/jpg");
                    av.LinkedResources.Add(YImage);


                    Bitmap bMail = new Bitmap(Properties.Resources.Mail_Icon);
                    ImageConverter icMail = new ImageConverter();
                    Byte[] mail = (Byte[])icMail.ConvertTo(bMail, typeof(Byte[]));
                    MemoryStream imageMail = new MemoryStream(mail);
                    LinkedResource MaiImage = new LinkedResource(imageMail, System.Net.Mime.MediaTypeNames.Image.Jpeg);
                    MaiImage.ContentId = "Mail-Icon.jpg";
                    MaiImage.ContentType = new ContentType("image/jpg");
                    av.LinkedResources.Add(MaiImage);


                    Bitmap bEX = new Bitmap(Properties.Resources.__icon);
                    ImageConverter icEX = new ImageConverter();
                    Byte[] EX = (Byte[])icMail.ConvertTo(bEX, typeof(Byte[]));
                    MemoryStream imageEX = new MemoryStream(EX);
                    LinkedResource EXImage = new LinkedResource(imageEX, System.Net.Mime.MediaTypeNames.Image.Jpeg);
                    EXImage.ContentId = "!-Icon.jpg";
                    EXImage.ContentType = new ContentType("image/jpg");
                    av.LinkedResources.Add(EXImage);

                    ContentType mimeType = new System.Net.Mime.ContentType("text/html");
                    AlternateView alternate = AlternateView.CreateAlternateViewFromString(bodywithIcon, mimeType);
                    sendMail.AlternateViews.Add(av);

                    sendMail.From = fromMail;
                    sendMail.Subject = mailSubject;
                    sendMail.IsBodyHtml = true;
                    sendMail.Body = bodywithIcon;
                   
                    try
                    {
                    SmtpClient theClient = new SmtpClient("*************");
                    theClient.UseDefaultCredentials = true;
                    theClient.UseDefaultCredentials = true;
                    System.Net.NetworkCredential theCredential = new System.Net.NetworkCredential();
                    theClient.Credentials = theCredential;
                    
                    // reset for production
                    theClient.Send(sendMail);

                       mailSent = true;
                    }
                     catch(SmtpException e)
                    {
                        // try again 
                        //string what = e.Message.ToString();

                        try
                        { // need a new try as if this fails we dont want it dropping out to calling functions catch.
                            SmtpClient theUclient = new SmtpClient("********");
                            theUclient.UseDefaultCredentials = true;
                            System.Net.NetworkCredential theCredential = new System.Net.NetworkCredential();
                            theUclient.Credentials = theCredential;
                            theUclient.Send(sendMail);
                            mailSent = true;
                        }
                        catch (Exception ef)
                        {
                            dateTimeStamp = DateTime.Now;
                            string justDate = dateTimeStamp.ToString(MDYformat);
                            updateLog(pathWay, fileName, justDate, ef.Message.ToString(), false);
                            // ouch
                        }
                    }

                }

            }

            catch (Exception f)
            {

                DateTime datetime = DateTime.Now;
                string MDYformat = "dd/MMM/yyyy HH:mm:ss";
                string justDate = datetime.ToString(MDYformat);
                updateLog(pathWay, fileName, justDate, sendTo + " failed " + f.Message.ToString(), true);

            }

            return mailSent;

        }


    }
}
