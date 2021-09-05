using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;

namespace EguibarIT.Housekeeping
{
    /// <summary>
    ///
    /// </summary>
    public class Helpers
    {
        /// <summary>
        /// Get user or group by SamAccountName, clear the attribute and reset the object security inheritance.
        /// </summary>
        /// <param name="SamAccountName"></param>
        /// <returns>String indicating the result</returns>
        public static string ClearAdminCount(string SamAccountName)
        {
            // allows inheritance
            bool isProtected = false;
            // preserves inherited rules
            bool PreserveInheritance = true;
            string result;

            // set up domain context
            using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain))
            {
                // Get the User Principal
                using (ExtPrincipal.UserPrincipalEx user = ExtPrincipal.UserPrincipalEx.FindByIdentity(ctx, SamAccountName))
                {
                    //Remove adminCount attribute
                    user.adminCount = null;
                    user.Save();

                    string _dnPath = string.Format("LDAP://{0}", user.DistinguishedName);

                    using (DirectoryEntry de = new DirectoryEntry(_dnPath))
                    {
                        ActiveDirectorySecurity acl = de.ObjectSecurity;

                        if (acl.AreAccessRulesProtected)
                        {
                            acl.SetAccessRuleProtection(isProtected, PreserveInheritance);
                            de.CommitChanges();
                        }//end if
                    }//end using DirectoryEntry

                    result = String.Format("Object: {0} Updated permissions blocked", user.DistinguishedName);
                }//end using UserPrincipalEx
            }//end using DomainContext

            return result;
        }

        /// <summary>
        /// Find staled users by Last Logon date offset
        /// </summary>
        /// <param name="DaysOffset">Time span of days as Int</param>
        /// <returns>List of UserPrincipalEx objects which has not logon in the DaysOffset time</returns>
        public static List<ExtPrincipal.UserPrincipalEx> StaleUsers(int DaysOffset)
        {
            List<ExtPrincipal.UserPrincipalEx> AllUsers = new List<ExtPrincipal.UserPrincipalEx>();

            using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain))
            {
                try
                {
                    // Define QueryByExample user
                    ExtPrincipal.UserPrincipalEx qbeUser = new ExtPrincipal.UserPrincipalEx(ctx);

                    DateTime timeStamp = DateTime.Now.Subtract(TimeSpan.FromDays(DaysOffset)).ToUniversalTime();

                    // Set the value to search from
                    //qbeUser.LastLogonTimeStamp > timeStamp;
                    PrincipalSearcher srch = new PrincipalSearcher(qbeUser);
                    srch.QueryFilter = qbeUser;

                    foreach (ExtPrincipal.UserPrincipalEx p in srch.FindAll())
                    {
                        if (p.StructuralObjectClass == "user")
                        {
                            if (p.LastLogonTimeStamp <= timeStamp)
                            {
                                AllUsers.Add(p);
                            }
                        }
                    }//end foreach
                }//end try
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    //doSomething with E.Message.ToString();
                    E.Message.ToString();
                    //WriteObject("ERROR - Something went wrong while adding MSA to the group.");
                }
                return AllUsers;
            }//end using
        }

        /// <summary>
        /// Find staled computers by Last Logon date offset
        /// </summary>
        /// <param name="DaysOffset">Time span of days as Int</param>
        /// <returns>List of ComputerPrincipalEx objects which has not logon in the DaysOffset time</returns>
        public static List<ExtPrincipal.ComputerPrincipalEx> StaleComputers(int DaysOffset)
        {
            List<ExtPrincipal.ComputerPrincipalEx> AllComputers = new List<ExtPrincipal.ComputerPrincipalEx>();

            using (PrincipalContext ctx = new PrincipalContext(ContextType.Domain))
            {
                try
                {
                    // Define QueryByExample user
                    ExtPrincipal.ComputerPrincipalEx qbeComputer = new ExtPrincipal.ComputerPrincipalEx(ctx);

                    DateTime timeStamp = DateTime.Now.Subtract(TimeSpan.FromDays(DaysOffset)).ToUniversalTime();

                    // Set the value to search from
                    //qbeUser.LastLogonTimeStamp > timeStamp;
                    PrincipalSearcher srch = new PrincipalSearcher(qbeComputer);
                    srch.QueryFilter = qbeComputer;

                    foreach (ExtPrincipal.ComputerPrincipalEx p in srch.FindAll())
                    {
                        if (p != null)
                        {
                            if (p.LastLogonTimeStamp <= timeStamp)
                            {
                                AllComputers.Add(p);
                            }
                        }
                    }//end foreach
                }//end try
                catch (System.DirectoryServices.DirectoryServicesCOMException E)
                {
                    //doSomething with E.Message.ToString();
                    E.Message.ToString();
                    //WriteObject("ERROR - Something went wrong while adding MSA to the group.");
                }
                return AllComputers;
            }//end using
        }

        /// <summary>
        /// Generates a random password based on the rules passed in the parameters
        /// </summary>
        /// <param name="includeLowercase">Bool to say if lowercase are required</param>
        /// <param name="includeUppercase">Bool to say if uppercase are required</param>
        /// <param name="includeNumeric">Bool to say if numerics are required</param>
        /// <param name="includeSpecial">Bool to say if special characters are required</param>
        /// <param name="includeSpaces">Bool to say if spaces are required</param>
        /// <param name="lengthOfPassword">Length of password required. Should be between 8 and 128</param>
        /// <returns>STRING representing the generated password</returns>
        public static string RandomPassword(bool includeLowercase, bool includeUppercase, bool includeNumeric, bool includeSpecial, bool includeSpaces, int lengthOfPassword)
        {
            const int MAXIMUM_IDENTICAL_CONSECUTIVE_CHARS = 2;
            const string LOWERCASE_CHARACTERS = "abcdefghijklmnopqrstuvwxyz";
            const string UPPERCASE_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            const string NUMERIC_CHARACTERS = "0123456789";
            const string SPECIAL_CHARACTERS = @"!#$%&*@\+=?-_/~^)(<>";
            const string SPACE_CHARACTER = " ";
            const int PASSWORD_LENGTH_MIN = 8;
            const int PASSWORD_LENGTH_MAX = 128;

            if (lengthOfPassword < PASSWORD_LENGTH_MIN || lengthOfPassword > PASSWORD_LENGTH_MAX)
            {
                return "Password length must be between 8 and 128 characters.";
            }

            string characterSet = "";

            if (includeLowercase)
            {
                characterSet += LOWERCASE_CHARACTERS;
            }

            if (includeUppercase)
            {
                characterSet += UPPERCASE_CHARACTERS;
            }

            if (includeNumeric)
            {
                characterSet += NUMERIC_CHARACTERS;
            }

            if (includeSpecial)
            {
                characterSet += SPECIAL_CHARACTERS;
            }

            if (includeSpaces)
            {
                characterSet += SPACE_CHARACTER;
            }

            char[] password = new char[lengthOfPassword];
            int characterSetLength = characterSet.Length;

            System.Random random = new System.Random();
            for (int characterPosition = 0; characterPosition < lengthOfPassword; characterPosition++)
            {
                password[characterPosition] = characterSet[random.Next(characterSetLength - 1)];

                bool moreThanTwoIdenticalInARow =
                    characterPosition > MAXIMUM_IDENTICAL_CONSECUTIVE_CHARS
                    && password[characterPosition] == password[characterPosition - 1]
                    && password[characterPosition - 1] == password[characterPosition - 2];

                if (moreThanTwoIdenticalInARow)
                {
                    characterPosition--;
                }
            }

            return string.Join(null, password);
        }
    }//end class
}//end namespace