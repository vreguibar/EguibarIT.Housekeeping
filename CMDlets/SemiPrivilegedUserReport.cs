using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Report on Semi-Privileged Users.</para>
    /// <para type="description">Provides a full report on existing Privileged and Semi-Privileged users. 
    /// For this CMDlet to work properly, ALL privileged and SemiPrivileged accounts. 
    /// must be located within the same OU. Failing to have all users located on the 
    /// mentioned container will show an incorrect report.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdSemiPrivilegedUsersReport "OU=Admin Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet indicating report path</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdSemiPrivilegedUsersReport "OU=Admin Users,OU=Admin,DC=EguibarIT,DC=local" "C:\Reports\"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdSemiPrivilegedUsersReport -AdminUsersDN "OU=Admin Users,OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters including path for the report</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdSemiPrivilegedUsersReport -AdminUsersDN "OU=Admin Users,OU=Admin,DC=EguibarIT,DC=local" -Path "C:\Reports\"</code>
    ///         </para>
    ///      </example>
    /// <remarks>Report on Semi-Privileged Users.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Get, "AdSemiPrivilegedUsersReport", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(System.IO.File))]
    public class AdSemiPrivilegedUsersReport : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[STRING] Admin User Account OU Distinguished Name. (ej. OU=Users,OU=Good,OU=Sites,DC=EguibarIT,DC=local).</para>
        /// <para type="description">Distinguished Name of the container where the Admin users are located.</para>
        /// </summary>
        [Parameter(
            Position = 0,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Distinguished Name of the container where the Admin Accounts are located."
            )]
        [ValidateNotNullOrEmpty]
        public string AdminUsersDN
        {
            get { return _adminusersdn; }
            set { _adminusersdn = value; }
        }

        private string _adminusersdn;

        /// <summary>
        /// <para type="inputType">[STRING] Path where the HTML report will be saved. (ej. C:\Reports\MyReport.html).</para>
        /// <para type="description">Path where the HTML report will be saved.</para>
        /// </summary>
        [Parameter(
            Position = 0,
            Mandatory = false,
            ValueFromPipeline = true,
            ValueFromPipelineByPropertyName = true,
            HelpMessage = "Path where the HTML report will be saved."
            )]
        [ValidateNotNullOrEmpty]
        public string Path
        {
            get { return _path; }
            set { _path = value; }
        }

        private string _path;

        #endregion Parameters definition

        #region Begin()

        /// <summary>
        ///
        /// </summary>
        protected override void BeginProcessing()
        {
            base.BeginProcessing();

            if (this.MyInvocation.BoundParameters.ContainsKey("Verbose"))
            {
                WriteVerbose("|=> ************************************************************************ <=|");
                WriteVerbose(DateTime.Today.ToShortDateString());
                WriteVerbose(string.Format("  Starting: {0}", this.MyInvocation.MyCommand));

                string paramVerbose;
                paramVerbose = "Parameters:\n";
                paramVerbose += string.Format("{0,-12}{1,-30}{2,-30}\n", null, " Key", " Value");
                paramVerbose += string.Format("{0,-12}{1,-30}{2,-30}\n", null, "----------", "----------");

                // display PSBoundparameters formatted nicely for Verbose output
                // var is iDictionary
                var pb = this.MyInvocation.BoundParameters; // | Format - Table - AutoSize | Out - String).TrimEnd()

                foreach (var item in pb)
                {
                    paramVerbose += string.Format("{0,-12}{1,-30}{2,-30}\n", null, item.Key, item.Value);
                }
                WriteVerbose(string.Format("{0}\n", paramVerbose));
            }
        }

        #endregion Begin()

        #region Process()

        /// <summary>
        ///
        /// </summary>
        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            //Define variables to hold Semi-Privileged users found
            List<ExtPrincipal.UserPrincipalEx> T0Users = new List<ExtPrincipal.UserPrincipalEx>();
            List<ExtPrincipal.UserPrincipalEx> T1Users = new List<ExtPrincipal.UserPrincipalEx>();
            List<ExtPrincipal.UserPrincipalEx> T2Users = new List<ExtPrincipal.UserPrincipalEx>();

            //This is the Variable report
            var report = new System.Text.StringBuilder();

            // set up domain context
            using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain))
            {
                // define a "query-by-example" principal - here, we search for a UserPrincipal
                // and with the first name (GivenName) of "Bruce" and a last name (Surname) of "Miller"
                ExtPrincipal.UserPrincipalEx qbeUser = new ExtPrincipal.UserPrincipalEx(ctx);
                //qbeUser.GivenName = "Bruce";
                //qbeUser.Surname = "Miller";
                qbeUser.EmployeeType = "T*";

                // create your principal searcher passing in the QBE principal
                using (PrincipalSearcher srch = new PrincipalSearcher())
                {
                    srch.QueryFilter = qbeUser;

                    // Perform the search
                    PrincipalSearchResult<Principal> results = srch.FindAll();

                    WriteVerbose("Getting the report data.");

                    // find all matches
                    foreach (ExtPrincipal.UserPrincipalEx found in results)
                    {
                        // do whatever here - "found" is of type "Principal" - it could be user, group, computer.....

                        switch (found.EmployeeType.ToLower())
                        {
                            case "t0":
                                T0Users.Add(found);
                                break;

                            case "t1":
                                T1Users.Add(found);
                                break;

                            case "t2":
                                T2Users.Add(found);
                                break;
                        }//end switch
                    }//end foreach

                    WriteVerbose("Building the report.");

                    //Define HTML Headers
                    report.Append(HTML._SemiPrivilegedUsersReport);
                    report.AppendLine(string.Format("<p><b>Report generated: </b>{0}</p>", DateTime.Now.ToString("dd/MMM/yyyy", System.Globalization.CultureInfo.InvariantCulture)));
                    report.AppendLine(string.Format("<p><b>Script executed from: </b>{0}</p>", Environment.MachineName));
                    report.AppendLine(string.Format("<p><b>Script executed by: </b>{0}</p>", System.Security.Principal.WindowsIdentity.GetCurrent().Name));

                    #region Tier0 Table

                    //Define Tier0 Table
                    report.AppendLine("<br/><section><table><caption>Tier 0 Accounts</caption><thead class='Tier0'>");
                    report.Append(HTML._SemiPrivilegedUsersReportTableHeaders);

                    //Fill Tier0 table
                    foreach (ExtPrincipal.UserPrincipalEx T0u in T0Users)
                    {
                        report.AppendLine("<tr>");
                        report.AppendLine(string.Format("<td>{0}</td>", T0u.SamAccountName));
                        report.AppendLine(string.Format("<td>{0}</td>", T0u.DisplayName));
                        report.AppendLine(string.Format("<td>{0}</td>", T0u.Description));
                        report.AppendLine(string.Format("<td>{0}</td>", T0u.LastLogonTimeStamp));
                        report.AppendLine(string.Format("<td>{0}</td>", T0u.LastPasswordSet));
                        report.AppendLine(string.Format("<td>{0}</td>", T0u.PasswordNeverExpires));
                        report.AppendLine(string.Format("<td>{0}</td>", T0u.Enabled));
                        report.AppendLine(string.Format("<td>{0}</td>", T0u.AccountExpirationDate));

                        report.Append("<td>");
                        PrincipalSearchResult<Principal> Groups = T0u.GetGroups();

                        //define table for the groups
                        report.AppendLine("<br/><table>");
                        foreach (GroupPrincipal group in Groups)
                        {
                            report.AppendLine(string.Format("<tr><td>{0}</tr></td>", group.SamAccountName));
                        }
                        //close table for the groups
                        report.AppendLine("</table><br/>");

                        report.AppendLine("</td></tr>");
                    }//end foreach

                    //Close Tier0 Table
                    report.AppendLine("</table></section><br/>");

                    #endregion Tier0 Table

                    #region Tier1 Table

                    //Define Tier0 Table
                    report.AppendLine("<br/><section><table><caption>Tier 1 Accounts</caption><thead class='Tier1'>");
                    report.Append(HTML._SemiPrivilegedUsersReportTableHeaders);

                    //Fill Tier0 table
                    foreach (ExtPrincipal.UserPrincipalEx T1u in T1Users)
                    {
                        report.AppendLine("<tr>");
                        report.AppendLine(string.Format("<td>{0}</td>", T1u.SamAccountName));
                        report.AppendLine(string.Format("<td>{0}</td>", T1u.DisplayName));
                        report.AppendLine(string.Format("<td>{0}</td>", T1u.Description));
                        report.AppendLine(string.Format("<td>{0}</td>", T1u.LastLogonTimeStamp));
                        report.AppendLine(string.Format("<td>{0}</td>", T1u.LastPasswordSet));
                        report.AppendLine(string.Format("<td>{0}</td>", T1u.PasswordNeverExpires));
                        report.AppendLine(string.Format("<td>{0}</td>", T1u.Enabled));
                        report.AppendLine(string.Format("<td>{0}</td>", T1u.AccountExpirationDate));

                        report.Append("<td>");
                        PrincipalSearchResult<Principal> Groups = T1u.GetGroups();

                        //define table for the groups
                        report.AppendLine("<br/><table>");
                        foreach (GroupPrincipal group in Groups)
                        {
                            report.AppendLine(string.Format("<tr><td>{0}</tr></td>", group.SamAccountName));
                        }
                        //close table for the groups
                        report.AppendLine("</table><br/>");

                        report.AppendLine("</td></tr>");
                    }//end foreach

                    //Close Tier0 Table
                    report.AppendLine("</table></section><br/>");

                    #endregion Tier1 Table

                    #region Tier2 Table

                    //Define Tier0 Table
                    report.AppendLine("<br/><section><table><caption>Tier 2 Accounts</caption><thead class='Tier2'>");
                    report.Append(HTML._SemiPrivilegedUsersReportTableHeaders);

                    //Fill Tier0 table
                    foreach (ExtPrincipal.UserPrincipalEx T2u in T2Users)
                    {
                        report.AppendLine("<tr>");
                        report.AppendLine(string.Format("<td>{0}</td>", T2u.SamAccountName));
                        report.AppendLine(string.Format("<td>{0}</td>", T2u.DisplayName));
                        report.AppendLine(string.Format("<td>{0}</td>", T2u.Description));
                        report.AppendLine(string.Format("<td>{0}</td>", T2u.LastLogonTimeStamp));
                        report.AppendLine(string.Format("<td>{0}</td>", T2u.LastPasswordSet));
                        report.AppendLine(string.Format("<td>{0}</td>", T2u.PasswordNeverExpires));
                        report.AppendLine(string.Format("<td>{0}</td>", T2u.Enabled));
                        report.AppendLine(string.Format("<td>{0}</td>", T2u.AccountExpirationDate));

                        report.Append("<td>");
                        PrincipalSearchResult<Principal> Groups = T2u.GetGroups();

                        //define table for the groups
                        report.AppendLine("<br/><table>");
                        foreach (GroupPrincipal group in Groups)
                        {
                            report.AppendLine(string.Format("<tr><td>{0}</tr></td>", group.SamAccountName));
                        }
                        //close table for the groups
                        report.AppendLine("</table><br/>");

                        report.AppendLine("</td></tr>");
                    }//end foreach

                    //Close Tier0 Table
                    report.AppendLine("</table></section><br/>");

                    #endregion Tier2 Table

                }//end using PrincipalSearcher

            }//end using ctx
             //Define the footer
            report.AppendLine("<footer><p>  .::   Report generated by EguibarIT.Housekeeping PowerShell CMDlet. EguibarIT © 2019   ::. </p></footer>");
            //Close the report
            report.AppendLine("</body></html>");

            //Give the name of the report
            string _fileName = string.Format("Semi-PrivilegedUsers_Report_{0}.html", DateTime.Now.ToString("yyyy-MMM-dd_HHmmss"));

            //If no path was given, use %TEMP%
            if (_path == null)
            {
                WriteVerbose("Path not provided. Using %TEMP% as the path.");

                _path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), _fileName);
            }

            if (!File.Exists(_path))
            {
                WriteVerbose("File does not exist. Creating the report.");

                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(_path))
                {
                    sw.Write(report.ToString());
                }
            }
            else
            {
                WriteWarning(string.Format("A file with the same name already exists on {0}", _path));
            }

            WriteObject(string.Format("The report was written to: {0}", _path));
        }

        #endregion Process()

        #region End()

        /// <summary>
        ///
        /// </summary>
        protected override void EndProcessing()
        {
            if (this.MyInvocation.BoundParameters.ContainsKey("Verbose"))
            {
                string paramVerboseEnd;
                paramVerboseEnd = string.Format("\n         Function {0} finished.\n", this.MyInvocation.MyCommand);
                paramVerboseEnd += "-------------------------------------------------------------------------------\n\n";

                WriteVerbose(paramVerboseEnd);
            }
        }

        #endregion End()

        /// <summary>
        ///
        /// </summary>
        protected override void StopProcessing()
        {
        }
    } // end class
}