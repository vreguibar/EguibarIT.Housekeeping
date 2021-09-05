using EguibarIT.Housekeeping.IP;
using System;
using System.Collections.Generic;
using System.Management.Automation;

namespace EguibarIT.Housekeeping.CMDlets
{
    /// <summary>
    /// <para type="synopsis">Find overlap Subnet objects.</para>
    /// <para type="description">This CMDlet reads all AD Subnet objects, and will look for any IP Range overlap
    /// Active Directory Sites and Services uses this information to optimize traffic
    /// (replication, authentication, authorization, etc) by providing the most
    /// optimal domain controller on the network. IP Layer 3 of the ASI model should
    /// be configured as subnet object, and each of these objects has to be associated
    /// to an existing site.</para>
    /// <example>
    ///     <para>This example shows how to use this CMDlet.</para>
    ///     <para>-        </para>
    ///     <para>
    ///         <code>Get-AdOverlapSubnets</code>
    ///     </para>
    /// </example>
    /// <example>
    ///     <para>This example shows how to use this CMDlet including Site information</para>
    ///     <para>-        </para>
    ///     <para>
    ///         <code>Get-AdOverlapSubnets -IncludeSite</code>
    ///     </para>
    /// </example>
    /// <remarks>Find overlap Subnet objects.</remarks>
    /// </summary>
    /// <para type="link" uri="(http://EguibarIT.eu)">[Eguibar Information Technology S.L. web site]</para>
    /// <list type="alertSet">
    ///     <item>
    ///     <term>Copiright and Version</term>
    ///         <description>
    ///             <para>Version:         1.0</para>
    ///             <para>DateModified:    22/Aug/2019</para>
    ///             <para>LasModifiedBy:   Vicente Rodriguez Eguibar</para>
    ///                 <para>.    vicente@eguibar.com</para>
    ///                 <para>.    Eguibar Information Technology S.L.</para>
    ///                 <para>.    http://www.eguibarit.com</para>
    ///         </description>
    ///     </item>
    /// </list>
    [Cmdlet(VerbsCommon.Get, "AdOverlapSubnets", ConfirmImpact = ConfirmImpact.Medium)]
    [OutputType(typeof(string))]
    public class AdOverlapSubnets : PSCmdlet
    {
        // Variable to hold all results
        private string _finalReport = string.Empty;

        #region Parameters definition

        /// <summary>
        ///     <para type="inputType">[SWITCH] (bool)</para>
        ///     <para type="description">Switch indicator. If present (TRUE), the Site information will be displayed.</para>
        /// </summary>
        [Parameter(
               Position = 0,
               Mandatory = false,
               ValueFromPipeline = true,
               ValueFromPipelineByPropertyName = true,
               HelpMessage = "Switch indicator. If present (TRUE), the Site information will be displayed."
            )]
        public SwitchParameter IncludeSite
        {
            get { return _includeSite; }
            set { _includeSite = value; }
        }

        private bool _includeSite;

        #endregion Parameters definition

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

            // Initial message
            Console.WriteLine("\n\n\n\n\n\n\n\n" +
                              "This CMDlet reads all AD Subnet objects, and will look for any IP Range overlap. \n" +
                              "Active Directory Sites & Services uses this information to optimize traffic \n" +
                              "(replication, authentication, authorization, etc) by providing the most \n" +
                              "optimal domain controller on the network. IP Layer 3 of the OSI model should \n" +
                              "be configured as subnet object, and each of these objects has to be associated \n" +
                              "to an existing site.\n" +
                              "================================================================================\n\n");
        }

        /// <summary>
        ///
        /// </summary>
        protected override void ProcessRecord()
        {
            base.ProcessRecord();

            // Define the Progress Record (Progress Bar to be displayed)
            int myId = 0;
            int progressCount = 0;
            string myActivity = "Checking all AD Subnet objects for overlaping.";
            string myStatus = "Progress:";
            ProgressRecord pr = new ProgressRecord(myId, myActivity, myStatus);

            //Read ALL subnets from AD and store it on var
            Dictionary<IPAddressRange, string> _allSubnets = EguibarIT.Housekeeping.GetFromAd.GetDirectorySubNet();

            // Iterate through all found subnets
            foreach (var _subnet in _allSubnets)
            {
                progressCount++;

                int PercentComplete = (progressCount * 100 / _allSubnets.Count);

                // Progress Record % completed
                pr.PercentComplete = PercentComplete;

                // Process Record Current Operation
                pr.CurrentOperation = string.Format("Procesing Subnet # {0}: NAME: {1}", progressCount, _subnet.Key.ToCidrString());

                // Process Record Status message
                pr.StatusDescription = string.Format("Processing {0} Subnetss. Complete %: {1}", _allSubnets.Count, PercentComplete);

                // Write the Progress Status
                WriteProgress(pr);

                int _overlapSubnetCount = 0;

                // Compare current subnet with all existing subnets
                foreach (var _containedSubnet in _allSubnets)
                {
                    // Exclude current subnet from itself.
                    if (!_subnet.Equals(_containedSubnet))
                    {
                        //Check overlap
                        if (_subnet.Key.Contains(EguibarIT.Housekeeping.IP.IPAddressRange.Parse(_containedSubnet.Key.ToCidrString())))
                        {
                            // Overlap found. Write message
                            //Console.WriteLine(string.Format("Subnet {0} contains {1}", _subnet.Key.ToCidrString(), _containedSubnet.Key.ToCidrString())); // is True.
                            _finalReport += string.Format("Subnet {0} contains {1}\n", _subnet.Key.ToCidrString(), _containedSubnet.Key.ToCidrString());
                            _overlapSubnetCount++;

                            //Check if switch on to display site info
                            if (_includeSite)
                            {
                                // Exclude subnets not assigned
                                if (!_subnet.Value.Equals("Not Assigned"))
                                {
                                    // Check if Subnet belongs to the same site.
                                    //If NOT, display a warning message under it.
                                    if (!_subnet.Value.Equals(_containedSubnet.Value))
                                        /*Console.WriteLine(
                                            string.Format("       These Subnets are assigned to different sites.\n" +
                                                          "               {0} is assigned to site:    {1}\n" +
                                                          "         while {2} is assigned to site:    {3}",
                                                            _subnet.Key.ToCidrString(),
                                                            _subnet.Value,
                                                            _containedSubnet.Key.ToCidrString(),
                                                            _containedSubnet.Value

                                         );*/
                                        _finalReport += string.Format("       These Subnets are assigned to different sites.\n" +
                                                              "               {0} is assigned to site:    {1}\n" +
                                                              "         while {2} is assigned to site:    {3}\n",
                                                                _subnet.Key.ToCidrString(),
                                                                _subnet.Value,
                                                                _containedSubnet.Key.ToCidrString(),
                                                                _containedSubnet.Value
                                                               );
                                }
                            }
                        }
                    }
                }
                if (_overlapSubnetCount > 0)
                {
                    //Console.WriteLine("------------------------------------------------------------");
                    //Console.WriteLine(string.Format("Range {0} contains {1} smaller subnets \n\n\n", _subnet.Key.ToCidrString(), _overlapSubnetCount));

                    _finalReport += "------------------------------------------------------------\n";
                    _finalReport += string.Format("Range {0} contains {1} smaller subnets \n\n\n", _subnet.Key.ToCidrString(), _overlapSubnetCount);
                }
            }//end for
        }//end ProcessRecord()

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

        /// <summary>
        ///
        /// </summary>
        protected override void StopProcessing()
        {
        }
    }
}