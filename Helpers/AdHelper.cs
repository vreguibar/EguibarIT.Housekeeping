using System;
using System.DirectoryServices;

namespace EguibarIT.Housekeeping.AdHelper
{
    //https://www.c-sharpcorner.com/uploadfile/dhananjaycoder/all-operations-on-active-directory-ad-using-C-Sharp/

    /// <summary>
    ///
    /// </summary>
    public class ActiveDirectoryHelper
    {
        /*
        /// <summary>
        ///
        /// </summary>
        /// <param name="userlogin"></param>
        /// <param name="groupName"></param>
        /// <returns></returns>
        public bool AddUserToGroup(string userlogin, string groupName)
        {
            try
            {
                _directoryEntry = null;
                ADManager admanager = new ADManager(LDAPDomain);
                admanager.AddUserToGroup(userlogin, groupName);
                return true;
            }
            catch (Exception ex)
            {
                ex.ToString();
                return false;
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="userlogin"></param>
        /// <param name="groupName"></param>
        /// <returns></returns>
        public bool RemoveUserToGroup(string userlogin, string groupName)
        {
            try
            {
                _directoryEntry = null;
                ADManager admanager = new ADManager("xxx", LDAPUser, LDAPPassword);
                admanager.RemoveUserFromGroup(userlogin, groupName);
                return true;
            }
            catch (Exception ex)
            {
                ex.ToString();
                return false;
            }
        }

        #endregion

        /*
        /// <summary>
        /// Gets the maxPwdAge property on the domain (exmpl.wdc.com).
        /// This is days until the password expires unless the account has PasswordNeverExpires.
        /// </summary>
        /// <param name="domain">Domain Name (exmpl.wdc.com)</param>
        /// <returns>Days until a password expires</returns>
        public long GetMaxPasswordAgeInDays(string domain)
        {
            long days = 0;

            const long NS_IN_A_DAY = -864000000000;

            try
            {
                _directoryEntry = null;
                _directorySearcher = null;

                _directorySearcher.Filter = "(maxPwdAge=*)";

                var result = _directorySearcher.FindOne();

                if (result != null && result.Properties.Contains("maxPwdAge"))
                {
                    long maxPwdAge = TryGetResult<long>(result, "maxPwdAge");

                    days = maxPwdAge / NS_IN_A_DAY;
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
                return 0;
            }

            return days;
        }

        /// <summary>
        /// Helper function that provides the directory path string to pass onto DirectoryEntry()
        /// constructor.
        /// https://github.com/westerndigitalcorporation/DirectoryLib
        /// </summary>
        /// <param name="baseDistinguishedName">Optional distinguished name of domain to search under.</param>
        /// <returns></returns>
        private string GetDirectoryPath(string baseDistinguishedName = null)
        {
            string path = null;

            if (!string.IsNullOrEmpty(baseDistinguishedName))
            {
                path = LDAPPath;
            }

            return path;
        }
        */

        /// <summary>
        /// Method to convert date to AD format
        /// </summary>
        /// <param name="date">date to convert</param>
        /// <returns>string containing AD formatted date</returns>
        public static string ToADDateString(DateTime date)
        {
            string year = date.Year.ToString();
            int month = date.Month;
            int day = date.Day;
            int hr = date.Hour;
            string sb = string.Empty;
            sb += year;

            if (month < 10)
            {
                sb += "0";
            }

            sb += month.ToString();
            //bv

            if (day < 10)
            {
                sb += "0";
            }

            sb += day.ToString();

            if (hr < 10)
            {
                sb += "0";
            }

            sb += hr.ToString();
            sb += "0000.0Z";

            return sb.ToString();
        }//end ToADDateString

        /// <summary>
        ///
        /// </summary>
        /// <param name="objectCls"></param>
        /// <param name="returnValue"></param>
        /// <param name="objectName"></param>
        /// <param name="LdapDomain"></param>
        /// <returns></returns>
        public string GetObjectDistinguishedName(objectClass objectCls, returnType returnValue, string objectName, string LdapDomain)
        {
            string distinguishedName = string.Empty;
            string connectionPrefix = "LDAP://" + LdapDomain;
            DirectoryEntry entry = new DirectoryEntry(connectionPrefix);
            DirectorySearcher mySearcher = new DirectorySearcher(entry);

            switch (objectCls)
            {
                case objectClass.user:
                    mySearcher.Filter = "(&(objectClass=user)(|(cn=" + objectName + ")(sAMAccountName=" + objectName + ")))";
                    break;

                case objectClass.group:
                    mySearcher.Filter = "(&(objectClass=group)(|(cn=" + objectName + ")(dn=" + objectName + ")))";
                    break;

                case objectClass.computer:
                    mySearcher.Filter = "(&(objectClass=computer)(|(cn=" + objectName + ")(dn=" + objectName + ")))";
                    break;
            }
            SearchResult result = mySearcher.FindOne();

            if (result == null)
            {
                throw new NullReferenceException
                ("unable to locate the distinguishedName for the object " + objectName + " in the " + LdapDomain + " domain");
            }
            DirectoryEntry directoryObject = result.GetDirectoryEntry();
            if (returnValue.Equals(returnType.distinguishedName))
            {
                distinguishedName = "LDAP://" + directoryObject.Properties
                    ["distinguishedName"].Value;
            }
            if (returnValue.Equals(returnType.ObjectGUID))
            {
                distinguishedName = directoryObject.Guid.ToString();
            }
            entry.Close();
            entry.Dispose();
            mySearcher.Dispose();
            return distinguishedName;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="objectDN"></param>
        /// <returns></returns>
        public string ConvertDNtoGUID(string objectDN)
        {
            //Removed logic to check existence first

            DirectoryEntry directoryObject = new DirectoryEntry(objectDN);
            return directoryObject.Guid.ToString();
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="GUID"></param>
        /// <returns></returns>
        public static string ConvertGuidToDn(string GUID)
        {
            DirectoryEntry ent = new DirectoryEntry();
            String ADGuid = ent.NativeGuid;
            DirectoryEntry x = new DirectoryEntry("LDAP://{GUID=" + ADGuid + ">");
            //change the { to <>

            return x.Path.Remove(0, 7); //remove the LDAP prefix from the path
        }
    }//end class
}//end namespace