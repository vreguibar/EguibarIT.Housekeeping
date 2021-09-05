using System.DirectoryServices;
using System.Net.NetworkInformation;

namespace EguibarIT.Housekeeping.AdHelper
{
    /// <summary>
    ///
    /// </summary>
    public class AdDomain
    {
        /// <summary>
        /// Gets the AD FQDS from current domain
        /// </summary>
        /// <returns>AD Fully Qualified Domain Name - FQDN</returns>
        public static string GetAdFQDN()
        {
            IPGlobalProperties properties = IPGlobalProperties.GetIPGlobalProperties();

            return properties.DomainName;
        }

        /// <summary>
        /// Gets NetBIOS domain name from DNS Domain Name
        /// </summary>
        /// <param name="dnsDomainName"></param>
        /// <returns>AD NetBIOS domain Name</returns>
        public static string GetNetbiosDomainName(string dnsDomainName)
        {
            string netbiosDomainName = string.Empty;

            using (DirectoryEntry rootDSE = new DirectoryEntry(string.Format("LDAP://{0}/RootDSE", dnsDomainName)))
            {
                string configurationNamingContext = rootDSE.Properties["configurationNamingContext"][0].ToString();

                DirectoryEntry searchRoot = new DirectoryEntry("LDAP://cn=Partitions," + configurationNamingContext);

                using (DirectorySearcher searcher = new DirectorySearcher(searchRoot))
                {
                    searcher.SearchScope = SearchScope.OneLevel;
                    searcher.PropertiesToLoad.Add("netbiosname");
                    searcher.Filter = string.Format("(&(objectcategory=Crossref)(dnsRoot={0})(netBIOSName=*))", dnsDomainName);

                    SearchResult result = searcher.FindOne();

                    if (result != null)
                    {
                        netbiosDomainName = result.Properties["netbiosname"][0].ToString();
                    }
                }
            }
            return netbiosDomainName;
        }

        /// <summary>
        /// Gets NetBIOS domain name finding the current AD FQDN
        /// </summary>
        /// <returns>AD NetBIOS domain Name</returns>
        public static string GetNetbiosDomainName()
        {
            return GetNetbiosDomainName(GetAdFQDN());
        }

        /// <summary>
        /// Get the "Default Naming Context" form the current Domain/Forest
        /// </summary>
        /// <returns>Default Naming Context</returns>
        public static string GetdefaultNamingContext()
        {
            using (DirectoryEntry ldapRoot = new DirectoryEntry(string.Format("LDAP://{0}/rootDSE", AdDomain.GetAdFQDN())))
            {
                return ldapRoot.Properties["defaultNamingContext"][0].ToString();
            }
        }

        /// <summary>
        /// Get the "Configuration Naming Context" form the current Domain/Forest
        /// </summary>
        /// <returns>Configuration Naming Context</returns>
        public static string GetConfigurationNamingContext()
        {
            using (DirectoryEntry ldapRoot = new DirectoryEntry(string.Format("LDAP://{0}/rootDSE", AdDomain.GetAdFQDN())))
            {
                return ldapRoot.Properties["ConfigurationNamingContext"][0].ToString();
            }
        }

        /// <summary>
        /// Get the "Schema Naming Context" form the current Domain/Forest
        /// </summary>
        /// <returns>Schema Naming Context</returns>
        public static string GetschemaNamingContext()
        {
            using (DirectoryEntry ldapRoot = new DirectoryEntry(string.Format("LDAP://{0}/rootDSE", AdDomain.GetAdFQDN())))
            {
                return ldapRoot.Properties["schemaNamingContext"][0].ToString();
            }
        }

        /// <summary>
        /// Get the "Root Domain Naming Context" form the current Domain/Forest
        /// </summary>
        /// <returns>Root Domain Naming Context</returns>
        public static string GetrootDomainNamingContext()
        {
            using (DirectoryEntry ldapRoot = new DirectoryEntry(string.Format("LDAP://{0}/rootDSE", AdDomain.GetAdFQDN())))
            {
                return ldapRoot.Properties["rootDomainNamingContext"][0].ToString();
            }
        }
    }
}