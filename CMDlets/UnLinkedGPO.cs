using Microsoft.GroupPolicy;
using System;
using System.Collections;
using System.Management.Automation;

// Microsoft.GroupPolicy Namespace
// https://docs.microsoft.com/en-us/previous-versions/windows/desktop/wmi_v2/class-library/microsoft-grouppolicy-namespace

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Find unlinked GPOs</para>
    /// <para type="description">Find unlinked GPO in the domain. Remove those if switch is present</para>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdUnLinkedGPO</code>
    ///         </para>
    ///     </example>
    ///     <example>
    ///         <para>This example shows how to use this CMDlet using named parameters</para>
    ///         <para>-        </para>
    ///         <para>
    ///             <code>Get-AdUnLinkedGPO -RemoveUnLinkedGpo</code>
    ///         </para>
    ///     </example>
    /// <remarks>Find unlinked GPOs</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    [Cmdlet(VerbsCommon.Get, "AdUnLinkedGPO", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(string))]
    public class AdUnLinkedGPO : PSCmdlet
    {
        #region Parameters definition

        /// <summary>
        /// <para type="inputType">[SWITCH] parameter (true or false). If present the value becomes TRUE, and the UnLinked GPOs will be removed</para>
        /// <para type="description">Switch indicator to remove the UnLinked GPOs</para>
        /// </summary>
        [Parameter(
               Position = 1,
               Mandatory = false,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Switch indicator. If present (TRUE), the UnLinked GPOs will be removed."
            )]
        public SwitchParameter RemoveUnLinkedGpo
        {
            get { return _removeunlinkedgpo; }
            set { _removeunlinkedgpo = value; }
        }

        private bool _removeunlinkedgpo;

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

            // Use current user’s domain
            GPDomain domain = new GPDomain();

            ArrayList unLinkedGpos = new ArrayList();

            // Get all GPOs in the domain
            GPSearchCriteria searchCriteria = new GPSearchCriteria();
            GpoCollection gpos = domain.SearchGpos(searchCriteria);

            WriteVerbose(string.Format("\r\nFound a total of {0} GPOs on domain {1}\r\n", gpos.Count, domain.DomainName));

            foreach (Gpo gpo in gpos)
            {
                // Search for SOMs where this GPO is linked
                searchCriteria = new GPSearchCriteria();
                searchCriteria.Add(SearchProperty.SomLinks, SearchOperator.Contains, gpo);
                SomCollection soms = domain.SearchSoms(gpo);

                // If the SomCollection.Count property is 0, that tells us that no linksexist for the specified GPO
                if (soms.Count == 0)
                {
                    unLinkedGpos.Add(gpo);
                }
            }

            WriteVerbose(string.Format("\r\nFound a total of {0} unlinked GPOs\r\n", unLinkedGpos.Count));

            WriteObject(unLinkedGpos);

            //Remove Unlinked GPOs
            if (_removeunlinkedgpo)
            {
                foreach (Gpo gpo in unLinkedGpos)
                {
                    try
                    {
                        WriteVerbose(string.Format("Deleting {0} UnLinked GPO.", gpo.DisplayName));

                        gpo.Delete();
                    }
                    catch (Exception ex)
                    {
                        //Console.WriteLine("An error occurred: '{0}'", ex.Message);
                        throw new ApplicationException(string.Format("An error occurred while deleting UnLinked GPO: '{0}'. Message is {1}", ex, ex.Message));
                    }
                }
            }
            /*
             //* "GPO Admin 1.0 Type Library" from "GPOAdmin.dll" -> GPMGMTLib

            GPMGMTLib.GPM gpm = new GPMGMTLib.GPM();
            GPMGMTLib.IGPMConstants gpc = gpm.GetConstants();

            GPMGMTLib.IGPMDomain gpd = gpm.GetDomain(EguibarIT.Housekeeping.AdHelper.AdDomain.GetAdFQDN(), "", gpc.UseAnyDC);
            GPMGMTLib.GPMSearchCriteria gps = gpm.CreateSearchCriteria();

            GPMGMTLib.IGPMGPOCollection gpoc = gpd.SearchGPOs(gps);

            string outputString = "";

            foreach (GPMGMTLib.GPMGPO name in gpoc)
            {
                outputString += "ID: " + name.ID + "\tName: " + name.DisplayName + "\r\n";
            }
            WriteVerbose(outputString);
            */
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