using System;
using Microsoft.GroupPolicy;
using System.Collections;
using System.DirectoryServices;
using System.IO;
using System.Linq;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Find orphaned GPO</para>
    /// <para type="description">An orphaned GPO is an AD GPO object which is missing its corresponding GPT in SYSVOL.
    /// This CMDlet can find all orphaned GPOs, and can delete them.</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdOrphanGPO</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet with parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdOrphanGPO -RemoveOrphanGPO</code>
    ///         </para>
    ///     </example>
    /// <remarks>Find orphaned GPO</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Get, "AdOrphanGPO", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(string))]
    public class AdOrphanGPO : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[SWITCH] parameter (true or false). If present the value becomes TRUE, and the Orphan GPO will be removed</para>
        /// <para type="description">Switch indicator to remove the Orphan GPO</para>
        /// </summary>
        [Parameter(
               Position = 1,
               Mandatory = false,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Switch indicator. If present (TRUE), the Orphan GPO will be removed."
            )]
        public SwitchParameter RemoveOrphanGPO
        {
            get { return _removeorphangpo; }
            set { _removeorphangpo = value; }
        }

        private bool _removeorphangpo;

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
            ArrayList gpos = new ArrayList();
            ArrayList gpts = new ArrayList();

            //SYSVOL path
            string unc = string.Format(@"\\" + "{0}" + @"\SYSVOL\{0}\Policies", EguibarIT.Housekeeping.AdHelper.AdDomain.GetAdFQDN());
            //GPC
            string GPOPoliciesDN = string.Format("CN=Policies,CN=System,{0}", EguibarIT.Housekeeping.AdHelper.AdDomain.GetdefaultNamingContext());

            // Get all GPOs
            using (DirectoryEntry gpoc = new DirectoryEntry("LDAP://" + GPOPoliciesDN))
            {
                DirectoryEntries store = gpoc.Children;

                foreach (DirectoryEntry gpo in store)
                {
                    gpos.Add(gpo.Name.Replace("CN=", ""));
                }
            }

            //Get all GPTs
            string[] dirs = Directory.GetDirectories(unc, "*.*", SearchOption.TopDirectoryOnly);

            foreach (string dir in dirs)
            {
                if (!dir.Contains("PolicyDefinitions"))
                {
                    gpts.Add(Path.GetFileName(dir));
                }
            }

            var OrphanedGPOs = gpos.ToArray().Except(gpts.ToArray());

            WriteVerbose(string.Format("Found {0} Orphaned GPOs", OrphanedGPOs.Count()));

            //Find orphaned GPOs (GPO existing without corresponding GPT)
            WriteObject(OrphanedGPOs);

            if (_removeorphangpo)
            {
                foreach (var gpo in OrphanedGPOs)
                {
                    GPDomain domain = new GPDomain();

                    Gpo currentGpo = domain.GetGpo(new System.Guid(gpo.ToString()));

                    WriteVerbose(string.Format("Deleting {0} Orphaned GPOs", currentGpo.DisplayName));
                    currentGpo.Delete();
                }
            }
        }//end ProcessRecord

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
}//end namespace