using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Function for Privileged Computers housekeeping</para>
    /// <para type="description">PAWs and Infrastructure Services must be managed by Tier0 administrators and services.
    /// Each typemust belong to a certain group(SL_PAWS or SL_InfrastructureServices) 
    /// in order to be mantained.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///     <para>-        </para>
    ///         <para>
    ///             <code>Set-PrivilegedComputerHousekeeping "OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///     <para>-        </para>
    ///         <para>
    ///             <code>Set-PrivilegedComputerHousekeeping -SearchRootDN "OU=Admin,DC=EguibarIT,DC=local"</code>
    ///         </para>
    ///     </example>
    /// <remarks>Ensures that PAWs and Infrastructure Services are managed as Tier0 assets by adding those to its corresponding group.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Set, "PrivilegedComputerHousekeeping", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(int))]
    public class PrivilegedComputer : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[STRING] Admin  OU Distinguished Name. (ej. OU=Users,OU=Good,OU=Sites,DC=EguibarIT,DC=local).</para>
        /// <para type="description">Distinguished Name of the container where the computers are located.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = true,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Admin OU Distinguished Name."
            )]
        [ValidateNotNullOrEmpty]
        public string SearchRootDN
        {
            get { return _searchrootdn; }
            set { _searchrootdn = value; }
        }

        private string _searchrootdn;

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

            // Item Counter
            int i = 0;

            // New found Servers counter
            int NewServer = 0;

            // New found PAW counter
            int NewPAW = 0;

            // set up domain context and given OU
            PrincipalContext ctx = new PrincipalContext(ContextType.Domain, EguibarIT.Housekeeping.AdHelper.AdDomain.GetNetbiosDomainName(), _searchrootdn);

            // Declare the list of groups
            List<EguibarIT.Housekeeping.ExtPrincipal.ComputerPrincipalEx> AllPrivComputers = new List<EguibarIT.Housekeeping.ExtPrincipal.ComputerPrincipalEx>();

            // Declare InfraServers group
            GroupPrincipal InfraServers = GroupPrincipal.FindByIdentity(ctx, "SL_InfrastructureServers");

            // Declare PAWs group
            GroupPrincipal PAW = GroupPrincipal.FindByIdentity(ctx, "SL_PAWs");

            // Define the Progress Record (Progress Bar to be displayed)
            int myId = 0;
            string myActivity = "Checking Privileged Computers (PAWs & Infra Servers)";
            string myStatus = "Progress:";
            ProgressRecord pr = new ProgressRecord(myId, myActivity, myStatus);

            try
            {
                //ComputerPrincipal qbeComputer = new ComputerPrincipal(ctx);
                EguibarIT.Housekeeping.ExtPrincipal.ComputerPrincipalEx qbeComputer = new EguibarIT.Housekeeping.ExtPrincipal.ComputerPrincipalEx(ctx);

                using (PrincipalSearcher srch = new PrincipalSearcher(qbeComputer))
                {
                    // iterate the results
                    foreach (EguibarIT.Housekeeping.ExtPrincipal.ComputerPrincipalEx p in srch.FindAll())
                    {
                        // Remove if Service Account
                        if (!p.DistinguishedName.Contains("Service"))
                        {
                            // Add the current group to the list.
                            AllPrivComputers.Add(p);
                        }
                    }//end for
                }//end using
            } //end try
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString();
                E.Message.ToString();
                //WriteObject("ERROR - Something went wrong while adding MSA to the group.");
            }

            // Total Objects Found
            int TotalObjectsFound = AllPrivComputers.Count;

            Console.WriteLine("");
            Console.WriteLine("");
            WriteVerbose(string.Format("Iterate through each computer returned. Total computers found: {0}", TotalObjectsFound));
            Console.WriteLine("");

            try
            {
                // Iterate through all groups returned
                foreach (EguibarIT.Housekeeping.ExtPrincipal.ComputerPrincipalEx computer in AllPrivComputers)
                {
                    i++;

                    int PercentComplete = (i * 100 / TotalObjectsFound);

                    // Progress Record % completed
                    pr.PercentComplete = PercentComplete;

                    // Process Record Current Operation
                    pr.CurrentOperation = string.Format("Procesing object # {0}: NAME: {1}", i, computer.Name);

                    // Process Record Status message
                    pr.StatusDescription = string.Format("Processing {0} objects. Complete %: {1}", TotalObjectsFound, PercentComplete);

                    // Write the Progress Status
                    WriteProgress(pr);

                    // Exclude computer from Housekeeping container
                    if (!computer.DistinguishedName.Contains("Housekeeping"))
                    {
                        if (computer != null)
                        {
                            if (computer.OperatingSystem != null)
                            {
                                if (computer.OperatingSystem.Contains("Server"))
                                {
                                    if (!computer.IsMemberOf(InfraServers))
                                    {
                                        InfraServers.Members.Add(computer);
                                        InfraServers.Save();

                                        WriteVerbose(string.Format("Adding found Server {0} to SL_InfrastructureServers group", computer.Name));

                                        NewServer++;
                                    }//end if
                                }//end if
                                else
                                {
                                    if (!computer.IsMemberOf(PAW))
                                    {
                                        PAW.Members.Add(computer);
                                        PAW.Save();

                                        WriteVerbose(string.Format("Adding found Server {0} to SL_PAWs group", computer.Name));

                                        NewPAW++;
                                    }//end if
                                }
                            } //end if
                        }
                    } //end if
                } //end foreach
            } //end try
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                E.Message.ToString();
            }

            pr.RecordType = ProgressRecordType.Completed;
            WriteProgress(pr);

            Console.WriteLine("");
            WriteVerbose("Any PAW or Infrastructure Server will be patched and managed by Tier0 services");
            WriteVerbose("------------------------------------------------------------------------------");
            WriteVerbose(string.Format("Servers found...: {0}", NewServer));
            WriteVerbose(string.Format("PAWs found......: {0}", NewPAW));
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
    } //end class
} //end namespace