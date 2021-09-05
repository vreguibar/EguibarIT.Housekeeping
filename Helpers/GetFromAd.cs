using EguibarIT.Housekeeping.IP;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;

namespace EguibarIT.Housekeeping
{
    /// <summary>
    ///
    /// </summary>
    public class GetFromAd
    {
        /// <summary>
        /// Get the Group Members of a given AD Group
        /// </summary>
        /// <param name="GroupName"></param>
        /// <returns>List(Principal) of users which are members of the group</returns>
        public static List<Principal> GetGroupMembers(string GroupName)
        {
            // Declare the returning object
            List<Principal> users = new List<Principal>();

            try
            {
                // set up domain context
                PrincipalContext ctx = new PrincipalContext(ContextType.Domain);

                // find the group in question
                GroupPrincipal group = GroupPrincipal.FindByIdentity(ctx, GroupName);

                // if found....
                if (group != null)
                {
                    // iterate over members
                    foreach (Principal p in group.GetMembers())
                    {
                        // do whatever you need to do to those members
                        if (p is Principal currentUser)
                        {
                            users.Add(currentUser);
                        }
                    } //foreach
                } //if group found
            }// end try
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString();
                E.Message.ToString();
                //WriteObject("ERROR - Something went wrong while adding MSA to the group.");
            }

            // Return the list
            return users;
        } // end function

        /// <summary>
        /// Get all users contained in a OU
        /// </summary>
        /// <param name="OuDN">Organizational Unit DistinguishedName</param>
        /// <returns>List(Principal) of users which are contained on the OU</returns>
        public static List<ExtPrincipal.UserPrincipalEx> GetUsersFromOU(string OuDN)
        {
            // Declare the returning object
            List<ExtPrincipal.UserPrincipalEx> users = new List<ExtPrincipal.UserPrincipalEx>();

            try
            {
                // set up domain context and given OU
                using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain, EguibarIT.Housekeeping.AdHelper.AdDomain.GetNetbiosDomainName(), OuDN))
                {
                    ExtPrincipal.UserPrincipalEx qbeUser = new ExtPrincipal.UserPrincipalEx(ctx);
                    using (PrincipalSearcher srch = new PrincipalSearcher(qbeUser))
                    {
                        foreach (ExtPrincipal.UserPrincipalEx p in srch.FindAll())
                        {
                            users.Add(p);
                        }
                    }
                }
            }// end try
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString();
                E.Message.ToString();
                //WriteObject("ERROR - Something went wrong while adding MSA to the group.");
            }
            // Return the list
            return users;
        } //end function

        /// <summary>
        /// Get all Managed Service Accounts from a given OU
        /// </summary>
        /// <param name="OuDN">Organizational Unit DistinguishedName</param>
        /// <returns>List(Principal) of Managed Service Accounts which are contained on the OU</returns>
        public static List<Principal> GetMSAsFromOU(string OuDN)
        {
            // Declare the returning object
            List<Principal> MSAs = new List<Principal>();

            try
            {
                // set up domain context and given OU
                using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain, EguibarIT.Housekeeping.AdHelper.AdDomain.GetNetbiosDomainName(), OuDN))
                {
                    ComputerPrincipal MSA = new ComputerPrincipal(ctx);

                    using (PrincipalSearcher srch = new PrincipalSearcher(MSA))
                    {
                        foreach (ComputerPrincipal p in srch.FindAll())
                        {
                            MSAs.Add(p);
                        }
                    }
                }
            }// end try
            catch (System.DirectoryServices.DirectoryServicesCOMException E)
            {
                //doSomething with E.Message.ToString();
                E.Message.ToString();
                //WriteObject("ERROR - Something went wrong while adding MSA to the group.");
            }

            // Return the list
            return MSAs;
        }

        /// <summary>
        /// Reads ALL subnets from AD
        /// </summary>
        /// <returns>List of type IPAddressRange</returns>
        public static Dictionary<IPAddressRange, string> GetDirectorySubNet()
        {
            // Declare the returning object
            Dictionary<IPAddressRange, string> _allSubnets = new Dictionary<IPAddressRange, string>();

            // Retrieve the Configuration Naming Context from RootDSE
            string ConfigNC = EguibarIT.Housekeeping.AdHelper.AdDomain.GetConfigurationNamingContext();

            // Connect to the Configuration Naming Context
            using (DirectoryEntry ConfigSearchRoot = new DirectoryEntry("LDAP://" + ConfigNC))
            {
                // Prepare the search
                using (DirectorySearcher ConfigSearch = new DirectorySearcher(ConfigSearchRoot))
                {
                    ConfigSearch.Filter = ("(objectClass=subnet)");
                    ConfigSearch.PropertiesToLoad.Add("name");
                    ConfigSearch.PropertiesToLoad.Add("siteObject");

                    // Iterate results
                    foreach (SearchResult SubNet in ConfigSearch.FindAll())
                    {
                        // Add the key as the Subnet name, which is unique
                        var _tmpIP = IPAddressRange.Parse(SubNet.Properties["name"][0].ToString());

                        string _tmpSite;

                        if (SubNet.Properties.Count == 3)
                        {
                            _tmpSite = ((SubNet.Properties["siteObject"][0].ToString()).Split(',')[0]).Replace("CN=", "");
                        }
                        else { _tmpSite = "Not Assigned"; }

                        // Add each subnet to the List
                        _allSubnets.Add(_tmpIP, _tmpSite);
                    }
                }
            }
            //Return the List<IPAddressRange>
            return _allSubnets;
        }

        /// <summary>
        ///
        /// </summary>
        public void GetDirectorySites()
        {
            // Retrieve the Configuration Naming Context from RootDSE
            string ConfigNC = EguibarIT.Housekeeping.AdHelper.AdDomain.GetConfigurationNamingContext();

            // Connect to the Configuration Naming Context
            DirectoryEntry ConfigSearchRoot = new DirectoryEntry("LDAP://" + ConfigNC);

            DirectorySearcher ConfigSearch = new DirectorySearcher(ConfigSearchRoot)
            {
                Filter = ("(objectClass=site)")
            };

            ConfigSearch.PropertiesToLoad.Add("name");

            foreach (SearchResult Site in ConfigSearch.FindAll())
            {
                string SiteName = Site.Properties["name"][0].ToString();
                Console.WriteLine("Site name: " + SiteName);
            }

            ConfigSearch.Dispose();
        }

        /// <summary>
        ///
        /// </summary>
        public void GetDirectorySitesAndSubnets()
        {
            foreach (ActiveDirectorySite site in Forest.GetCurrentForest().Sites)
            {
                Console.WriteLine(site.Name);
                Console.WriteLine("----------------------------------------");
                foreach (ActiveDirectorySubnet subnet in site.Subnets)
                {
                    Console.WriteLine("  {0}", subnet.Name);
                }
                Console.WriteLine("\n\n");
            }
        }
    } //end class
} //end namespace