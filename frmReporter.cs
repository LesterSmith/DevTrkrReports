﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using DataHelpers;
using BusinessObjects;
using System.Diagnostics;
using System.Runtime.InteropServices;
using AppWrapper;
namespace DevTrkrReports
{
    public partial class frmReporter : Form
    {
        private List<NotableApplication> AppList { get; set; }
        private List<DevProjPath> ProjList { get; set; }
        public frmReporter()
        {
            InitializeComponent();
            Application.DoEvents();
        }

        private void frmReporter_Load(object sender, EventArgs e)
        {
            Application.DoEvents();
            var hlpr = new DHMisc(string.Empty);
            ProjList = hlpr.GetDevProjects();
            lbProjects.Items.Clear();
            lbProjects.Items.Add("All");
            foreach (var item in ProjList)
                lbProjects.Items.Add(item.DevProjectName);

            List<DeveloperNames> developers = hlpr.GetDeveloperNames();
            lbDevelopers.Items.Clear();
            lbDevelopers.Items.Add("All");
            foreach (var developer in developers)
                lbDevelopers.Items.Add(developer.UserName + " - " + developer.UserDisplayName);

            AppList = hlpr.GetNotableApplications();
            lbApplications.Items.Clear();
            lbApplications.Items.Add("All Applications");
            lbApplications.Items.Add("All Listed Applications");

            foreach (var item in AppList)
                lbApplications.Items.Add(item.AppFriendlyName);
        }

        private void btnCreateReport_Click(object sender, EventArgs e)
        {
            if (!ValidateReportParameters())
                return;


            var range = "A1:Z2";
            string dv = chkUseDates.Checked ? $"from {dtStart.Value.Date.ToString("MM/dd/yyyy")} to { dtEnd.Value.Date.ToString("MM/dd/yyyy")}" : string.Empty;
            switch (cbReportType.SelectedItem.ToString())
            {
                case "Project Summary by User":
                    MarkSelectedProjects();
                    List<ColHdr> cols = new List<ColHdr>();
                    cols.Add(new ColHdr { Header = "Project Name", Width = 40 });
                    cols.Add(new ColHdr { Header = "Hours", Width = 10 });
                    cols.Add(new ColHdr { Header = "Minutes", Width = 12 });
                    cols.Add(new ColHdr { Header = "Seconds", Width = 12 });
                    cols.Add(new ColHdr { Header = "Developer", Width = 20 });
                    cols.Add(new ColHdr { Header = "Project Path", Width = 50 });
                    
                    var prHdr = new ReportHdr 
                                            { HdrRange = range, 
                                              Hdrs = cols, 
                                              Title = $"Project Time Report by User {dv}", 
                                              TitleCell = "A1" 
                                            };
                    ProjectReportByUser pr = (ProjectReportByUser) ReportFactory.Factory(
                        new ReporterParms 
                        {
                            Header = prHdr,   
                            FileName = txtFilename.Text, 
                            Type = ReportType.ProjectSummaryByUser, 
                            StartTime = chkUseDates.Checked ? dtStart.Value : (DateTime?)null, 
                            EndTime = chkUseDates.Checked ? dtEnd.Value : (DateTime?)null 
                        });
                    List<DeveloperNames> developers = GetDevelopers(lbDevelopers);

                    if (pr.Process(ProjList, developers)) //, dtStart.Value, dtEnd.Value))
                    {
                        btnOpenReport.Enabled = true;
                        MessageBox.Show("Your report is created, click Open Report Button to view in Excel.", "Report Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        btnCreateReport.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("For some reason your report did not complete, check you database and notify vendor.", "Report Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    pr.Dispose();
                    break;
                case "Project Summary by Project":
                    MarkSelectedProjects();
                    List<ColHdr> colsPP = new List<ColHdr>();
                    colsPP.Add(new ColHdr { Header = "Project Name", Width = 40 });
                    colsPP.Add(new ColHdr { Header = "Hours", Width = 10 });
                    colsPP.Add(new ColHdr { Header = "Minutes", Width = 12 });
                    colsPP.Add(new ColHdr { Header = "Seconds", Width = 12 });
                    colsPP.Add(new ColHdr { Header = "Developer", Width = 20 });
                    colsPP.Add(new ColHdr { Header = "Project Path", Width = 50 });
                    
                    var ppHdr = new ReportHdr
                    {
                        HdrRange = range,
                        Hdrs = colsPP,
                        Title = $"Project Time Report by Project {dv}",
                        TitleCell = "A1"
                    };
                    ProjectReportByProject prp = (ProjectReportByProject)ReportFactory.Factory(
                        new ReporterParms
                        {
                            Header = ppHdr,
                            FileName = txtFilename.Text,
                            Type = ReportType.ProjectSummaryByProject,
                            StartTime = chkUseDates.Checked ? dtStart.Value : (DateTime?)null,
                            EndTime = chkUseDates.Checked ? dtEnd.Value : (DateTime?)null
                        });
                    List<DeveloperNames> devs = GetDevelopers(lbDevelopers);

                    if (prp.Process(ProjList, devs)) //, dtStart.Value, dtEnd.Value))
                    {
                        btnOpenReport.Enabled = true;
                        MessageBox.Show("Your report is created, click Open Report Button to view in Excel.", "Report Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        btnCreateReport.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("For some reason your report did not complete, check you database and notify vendor.", "Report Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    prp.Dispose();
                    break;
                case "Application Usage":
                    MarkSelectedApps();
                    List<ColHdr> colsAU = new List<ColHdr>();
                    colsAU.Add(new ColHdr { Header = "Application", Width = 40 });
                    colsAU.Add(new ColHdr { Header = "Hours", Width = 10 });
                    colsAU.Add(new ColHdr { Header = "Minutes", Width = 12 });
                    colsAU.Add(new ColHdr { Header = "Seconds", Width = 12 });
                    colsAU.Add(new ColHdr { Header = "Developer", Width = 20 });
                    var arHdr = new ReportHdr { HdrRange = range, Hdrs = colsAU, Title = $"Application Usage Report {dv}", TitleCell = "A1" };
                    ApplicationReport ar = (ApplicationReport)ReportFactory.Factory(new ReporterParms { Header = arHdr, FileName = txtFilename.Text, Type = ReportType.ApplicationUsage, StartTime = chkUseDates.Checked ? dtStart.Value : (DateTime?)null, EndTime = chkUseDates.Checked ? dtEnd.Value : (DateTime?)null });
                    List<DeveloperNames> users = GetDevelopers(lbDevelopers);

                    if (ar.Process(AppList, users)) //, dtStart.Value, dtEnd.Value))
                    {
                        btnOpenReport.Enabled = true;
                        MessageBox.Show("Your report is created, click Open Report Button to view in Excel.", "Report Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        btnCreateReport.Enabled = false;
                    }
                    else
                    {
                        MessageBox.Show("For some reason your report did not complete, check you database and notify vendor.", "Report Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    ar.Dispose();
                    break;
                case "Developer Detail":
                    MarkSelectedProjects();
                    List<ColHdr> colsUR = new List<ColHdr>();
                    colsUR.Add(new ColHdr { Header = "Project Name", Width = 40 });
                    colsUR.Add(new ColHdr { Header = "Hours", Width = 10 });
                    colsUR.Add(new ColHdr { Header = "Minutes", Width = 12 });
                    colsUR.Add(new ColHdr { Header = "Seconds", Width = 12 });
                    colsUR.Add(new ColHdr { Header = "Developer", Width = 20 });
                    colsUR.Add(new ColHdr { Header = "Activity", Width = 70 });
                    //string dv = chkUseDates.Checked ? $"from { dtStart.Value.Date.ToString("MM/dd/yyyy")} to { dtEnd.Value.Date.ToString("MM/dd/yyyy")}" : string.Empty; 
                    var urHdr = new ReportHdr { HdrRange = range, Hdrs = colsUR, Title = $"Developer Detail {dv})" , TitleCell = "A1" };
                    UserReport ur = (UserReport)ReportFactory.Factory(new ReporterParms { Header = urHdr, FileName = txtFilename.Text, Type = ReportType.DeveloperDetail, StartTime = chkUseDates.Checked ? dtStart.Value.Date : (DateTime?)null, EndTime = chkUseDates.Checked ? dtEnd.Value.Date : (DateTime?)null }); 
                    List<DeveloperNames> developer = GetDevelopers(lbDevelopers);
                    if (ur.Process(ProjList, developer))
                    {
                        btnOpenReport.Enabled = true;
                        MessageBox.Show("Your report is created, click Open Report Button to view in Excel.", "Report Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        btnCreateReport.Enabled = false;
                    }
                    ur.Dispose();
                    break;
            }
        }

        private void MarkSelectedProjects()
        {
            foreach (var p in ProjList)
                p.Selected = false;

            if (lbProjects.SelectedItems.Count == 1 && lbProjects.SelectedIndices[0] == 0)
            {
                foreach (var item in ProjList)
                {
                    item.Selected = true;
                }
                return;
            }

            for (int i = 0; i < lbProjects.SelectedItems.Count; i++)
            {
                var prj = ProjList.Find(x => x.DevProjectName == lbProjects.SelectedItems[i].ToString());
                if (prj != null)
                    prj.Selected = true;
            }
        }
        private void MarkSelectedApps()
        {
            foreach (var p in AppList)
                p.Selected = false;

            if (lbApplications.SelectedItems.Count == 1 && (string)lbApplications.SelectedItem=="All Listed Applications")
            {
                // if all apps selected, mark none selected
                // the report will not generate a where clause
                // for the applications
                foreach (var item in AppList)
                {
                    item.Selected = true;
                }
                return;
            }
            else if (lbApplications.SelectedItems.Count == 1 && (string)lbApplications.SelectedItem == "All Applications")
            {
                return;
            }

            for (int i = 0; i < lbApplications.SelectedItems.Count; i++)
            {
                var app = AppList.Find(x => x.AppFriendlyName == lbProjects.SelectedItems[i].ToString());
                if (app != null)
                    app.Selected = true;
            }
        }

        
        /// <summary>
        /// Returns 1 item with Username= ALL if all selected
        /// else returns list of selected usernames with selected = true
        /// </summary>
        /// <param name="lb"></param>
        /// <param name="onlyOne"></param>
        /// <returns></returns>
        private List<DeveloperNames> GetDevelopers(ListBox lb, bool onlyOne=false)
        {
            List<DeveloperNames> developers = new List<DeveloperNames>();
            if (lb.SelectedItems.Count.Equals(1) && lb.SelectedItem.ToString().Equals("All"))
            {
                developers.Add(new DeveloperNames { UserName = "All" });
                return developers;
            }
            foreach (var item in lb.SelectedItems)
            {
                var s = Util.GetStringFromFrontOfText(item.ToString());
                var udn = Util.GetStringFromEndOfText(item.ToString());
                developers.Add(new DeveloperNames { UserName = s, Selected=true, UserDisplayName = udn});
            }
            return developers;
        }

        private bool ValidateReportParameters()
        {
            var msg = string.Empty;
            if (string.IsNullOrWhiteSpace(cbReportType.Text))
                msg = "Please select a Report Type." + Environment.NewLine;

            if (lbProjects.SelectedItems.Count.Equals(0) && cbReportType.Text.Equals("Project Report"))
                msg += "You must select one or more or All Projects." + Environment.NewLine;

            //if (lbApplications.SelectedItems.Count.Equals(0) && cbReportType.Text.Equals("Application Report"))
            //    msg += "You must select one or more or All Applications.";

            if (cbReportType.Text.Equals("User Detail") && lbDevelopers.SelectedItems.Count != 1)
                msg += "The User Detail report can only be run for one user at a time.";
            else if (lbDevelopers.SelectedItems.Count.Equals(0))
                msg += "You must select one or more or All developers.";

            if ((dtStart.Value >= dtEnd.Value) && chkUseDates.Checked)
                msg += "End Date and Time must be greater than the Start Date and Time";

            if (string.IsNullOrWhiteSpace(txtFilename.Text))
                msg += "You must select an output file for your report.";
            
            if (cbReportType.Text.Equals("Application Usage") && lbApplications.SelectedItems.Count.Equals(0))
                msg += "For the Application Usage Report, you must select from the Applications List.";
           
            if (!string.IsNullOrWhiteSpace(msg))
            {
                MessageBox.Show(msg, "Invalid Report Selections");
                return false;
            }
            return true;
        }

        private void btnFileBrowse_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Excel File|*.xlsx";
            saveFileDialog1.Title = "Select Output File";
            saveFileDialog1.ShowDialog();
            txtFilename.Text = saveFileDialog1.FileName;
            if (!string.IsNullOrWhiteSpace(txtFilename.Text))
            {
                btnCreateReport.Enabled = true;
                btnOpenReport.Enabled = false;
            }
        }

        private string EnumerateSelections(ListBox lb)
        {
            var pipe = "|"; 
            var s = string.Empty;
            foreach (var li in lb.SelectedItems)
                s += string.IsNullOrWhiteSpace(s) ? li.ToString() : pipe + li.ToString();
            return s;
        }

        private void btnOpenReport_Click(object sender, EventArgs e)
        {
            if (File.Exists(txtFilename.Text))
                Process.Start(txtFilename.Text);
        }

        private void chkUseDates_CheckedChanged(object sender, EventArgs e)
        {
            if (chkUseDates.Checked)
            {
                dtStart.Enabled = true;
                dtEnd.Enabled = true;
            }
            else
            {
                dtStart.Enabled = false;
                dtEnd.Enabled = false;
            }
        }

        private bool RollupReferencedLibraries { get; set; }
        private void mnuRollupReferencedLibraries_Click(object sender, EventArgs e)
        {
            //mnuRollupReferencedLibraries.Checked = !
        }

        private void cbReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
           //switch (cbReportType.Text)
           // {
           //     case "Developer Detail":
           //         lbDevelopers.SelectionMode = SelectionMode.One;
           //         if (lbDevelopers.Items[0].ToString().ToLower().Equals("all"))
           //             lbDevelopers.Items.RemoveAt(0);
           //         break;
           //     default:
           //         lbDevelopers.SelectionMode = SelectionMode.MultiSimple;
           //         if (!lbDevelopers.Items[0].ToString().ToLower().Equals("all"))
           //             lbDevelopers.Items.Insert(0, "All");
           //         break;
           // }

        }

        private void frmReporter_FormClosing(object sender, FormClosingEventArgs e)
        {

        }
    }
}