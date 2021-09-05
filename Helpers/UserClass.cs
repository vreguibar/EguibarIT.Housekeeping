using System;
using System.Collections.Generic;
using System.DirectoryServices;

namespace EguibarIT.Housekeeping.AdHelper
{
    /// <summary>
    /// Class representing an AD User object
    /// </summary>
    public class AD_UserFull : AD_Object, UserMethods
    {
        #region Properties

        /// <summary>
        /// The AdminCount attribute.
        /// If 0 object is not protected by the AdminSDHolder process
        /// </summary>
        public int adminCount { get; set; }

        /// <summary>
        /// Represents the name of a locality, such as a town or city.
        /// City. (LDAP = l)
        /// </summary>
        public String City { get; set; }

        /// <summary>
        /// The common name that is usually the same as DisplayName.
        /// Example: Darth Vader
        /// </summary>
        public String CommonName { get; set; }

        /// <summary>
        /// Company. (LDAP = company)
        /// </summary>
        public String Company { get; set; }

        /// <summary>
        /// The country/region in the address of the user.
        /// The country/region is represented as a 2-character code based on ISO-3166.
        /// CountryName. (LDAP = co)
        /// </summary>
        public String Country { get; set; }

        /// <summary>
        /// Contains the name for the department in which the user works.
        /// Example: Master Commander
        /// </summary>
        public String Department { get; set; }

        /// <summary>
        /// Contains the description to display for an object.
        /// This value is restricted as single-valued for backward compatibility in some cases.
        /// </summary>
        public String Description { get; set; }

        /// <summary>
        /// The display name for an object. This is usually
        /// the combination of the users first name, middle initial, and last name.
        /// Example: Darth Vader
        /// </summary>
        public String DisplayName { get; set; }

        /// <summary>
        /// Same as the Distinguished Name for an object.
        /// </summary>
        public String DistinguishedName { get; set; }

        /// <summary>
        /// The email address for this contact.
        /// Example: darth.vader@eguibarit.com
        /// </summary>
        public String EmailAddress { get; set; }

        /// <summary>
        /// Telephone Number
        /// </summary>
        public String TelephoneNumber { get; set; }

        /// <summary>
        /// Fax number
        /// </summary>
        public String Fax { get; set; }

        /// <summary>
        /// Contains the given name (first name) of the user.
        /// Example: Darth
        /// </summary>
        public String FirstName { get; set; }

        /// <summary>
        /// Home Telephone number
        /// </summary>
        public String HomePhone { get; set; }

        /// <summary>
        /// This attribute contains the family or last name for a user.
        /// Example: Vader
        /// </summary>
        public String LastName { get; set; }

        /// <summary>
        /// Get the Manager by DistinguishedName
        /// </summary>
        public AD_UserFull Manager
        {
            get
            {
                if (!String.IsNullOrEmpty(_managerName))
                {
                    AD_UserFull ad = new AD_UserFull();
                    return ad.GetUserByDistinguishedName(_managerName);
                }
                return null;
            }
        }

        private String _manager;

        /// <summary>
        /// Manager Name
        /// </summary>
        public String ManagerName { get { return _managerName; } }

        private String _managerName;

        /// <summary>
        /// Middle Name
        /// </summary>
        public String MiddleName { get; set; }

        /// <summary>
        /// The primary mobile phone number.
        /// </summary>
        public String Mobile { get; set; }

        /// <summary>
        /// GUID of object
        /// </summary>
        public override byte[] ObjectGuid { get; set; }

        /// <summary>
        /// SID of object
        /// </summary>
        public override byte[] ObjectSID { get; set; }

        /// <summary>
        /// The primary telephone number.
        /// Example: 949-672-7000
        /// </summary>
        public String Phone { get; set; }

        /// <summary>
        /// Postal Code
        /// </summary>
        public String PostalCode { get; set; }

        /// <summary>
        /// Office number or cubicle number.
        /// Example: 1-2079
        /// </summary>
        public String PhysicalDeliveryOfficeName { get; set; }

        /// <summary>
        /// The logon name used to support clients and servers running earlier versions of the operating system.
        /// This attribute must be 20 characters or less to support earlier clients.
        /// </summary>
        public String SamAccountName { get; set; }

        /// <summary>
        /// The name of a user's state or province.
        /// </summary>
        public String State { get; set; }

        /// <summary>
        /// Street Address
        /// </summary>
        public String StreetAddress { get; set; }

        /// <summary>
        /// Contains the user's job title. This property is commonly used to indicate the formal job title,
        /// such as Senior Programmer, rather than occupational class, such as programmer.
        /// It is not typically used for suffix titles such as Esq. or DDS.
        /// Example: Dark Side Master
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// Used to store an image of a person.
        /// </summary>
        public byte[] thumbnailPhoto { get; set; }

        /// <summary>
        /// User Account Control
        /// </summary>
        public UserAccountControl userAccountControl { get; set; }

        /// <summary>
        /// This attribute contains the UPN that is an Internet-style login name for a user based on the
        /// Internet standard RFC 822. The UPN is shorter than the distinguished name and easier to remember.
        /// By convention, this should map to the user email name, but this is not required.
        /// The value set for this attribute is equal to the length of the user's ID and the domain name.
        /// Typically formatted as lastname_f@domain
        /// </summary>
        public String UserPrincipalName { get; set; }

        /// <summary>
        /// Implement UserAccountControl class
        /// </summary>
        public ADUserAccountControl aDUserAccountControl { get; set; }

        #endregion Properties

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        public AD_UserFull() : base() { }

        /// <summary>
        /// Constructor method with parameters
        /// </summary>
        /// <param name="de">DirectoryEntry for the object to be loaded</param>
        public AD_UserFull(DirectoryEntry de) : base()
        {
            GetUser(de);
        }

        #endregion Constructor

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        #region Members

        /// <summary>
        /// Returns the AD_UserFull object from the given DirectoryEntry
        /// </summary>
        /// <param name="de">DirectoryEntry</param>
        /// <returns>AD_UserFull object</returns>
        public AD_UserFull GetUser(DirectoryEntry de)
        {
            this.City = TryGetResult<string>(de, AdProperties.CITY);
            this.CommonName = TryGetResult<string>(de, AdProperties.COMMONNAME);
            this.Company = TryGetResult<string>(de, AdProperties.COMPANY);
            this.Country = TryGetResult<string>(de, AdProperties.COUNTRY);
            this.Department = TryGetResult<string>(de, AdProperties.DEPARTMENT);
            this.Description = TryGetResult<string>(de, AdProperties.DESCRIPTION);
            this.DisplayName = TryGetResult<string>(de, AdProperties.DISPLAYNAME);
            this.DistinguishedName = TryGetResult<string>(de, AdProperties.DISTINGUISHEDNAME);
            this.EmailAddress = TryGetResult<string>(de, AdProperties.EMAILADDRESS);
            this.TelephoneNumber = TryGetResult<string>(de, AdProperties.EXTENSION);
            this.Fax = TryGetResult<string>(de, AdProperties.FAX);
            this.FirstName = TryGetResult<string>(de, AdProperties.FIRSTNAME);
            this.HomePhone = TryGetResult<string>(de, AdProperties.HOMEPHONE);
            this.LastName = TryGetResult<string>(de, AdProperties.LASTNAME);
            this._manager = TryGetResult<string>(de, AdProperties.MANAGER);
            if (!String.IsNullOrEmpty(_manager))
            {
                String[] managerArray = _manager.Split(',');
                this._managerName = managerArray[0].Replace("CN=", "");
            }

            this.MiddleName = TryGetResult<string>(de, AdProperties.MIDDLENAME);
            this.Mobile = TryGetResult<string>(de, AdProperties.MOBILE);
            this.Phone = TryGetResult<string>(de, AdProperties.TELEPHONENUMBER);
            this.PostalCode = TryGetResult<string>(de, AdProperties.POSTALCODE);
            this.PhysicalDeliveryOfficeName = TryGetResult<string>(de, AdProperties.PHYSICALDELIVERYOFFICENAME);
            this.SamAccountName = TryGetResult<string>(de, AdProperties.SAMACCOUNTNAME);
            this.State = TryGetResult<string>(de, AdProperties.STATE);
            this.StreetAddress = TryGetResult<string>(de, AdProperties.STREETADDRESS);
            this.Title = TryGetResult<string>(de, AdProperties.TITLE);
            this.UserPrincipalName = TryGetResult<string>(de, AdProperties.USERPRINCIPALNAME);
            this.ObjectGuid = TryGetResult<byte[]>(de, AdProperties.OBJECTGUID);
            this.ObjectSID = TryGetResult<byte[]>(de, AdProperties.OBJECTSID);

            //_thumbnailPhoto = GetProperty(de, AdProperties.THUMBNAILPHOTO);

            this.userAccountControl = TryGetResult<UserAccountControl>(de, AdProperties.USERACCOUNTCONTROL);

            aDUserAccountControl = new ADUserAccountControl();

            bool z = (aDUserAccountControl.notDelegated == false);

            return this;
        }

        private void UpdateUser(DirectoryEntry de)
        {
            TrySetProperty(de, AdProperties.CITY, City);
            TrySetProperty(de, AdProperties.COMMONNAME, CommonName);
            TrySetProperty(de, AdProperties.COMPANY, Company);
            TrySetProperty(de, AdProperties.COUNTRY, Country);
            TrySetProperty(de, AdProperties.DEPARTMENT, Department);
            TrySetProperty(de, AdProperties.DESCRIPTION, Description);
            TrySetProperty(de, AdProperties.DISPLAYNAME, DisplayName);
            TrySetProperty(de, AdProperties.EMAILADDRESS, EmailAddress);
            TrySetProperty(de, AdProperties.EXTENSION, TelephoneNumber);
            TrySetProperty(de, AdProperties.FAX, Fax);
            TrySetProperty(de, AdProperties.FIRSTNAME, FirstName);
            TrySetProperty(de, AdProperties.HOMEPHONE, HomePhone);
            TrySetProperty(de, AdProperties.LASTNAME, LastName);
            TrySetProperty(de, AdProperties.MANAGER, ManagerName);
            TrySetProperty(de, AdProperties.MIDDLENAME, MiddleName);
            TrySetProperty(de, AdProperties.MOBILE, Mobile);
            TrySetProperty(de, AdProperties.TELEPHONENUMBER, Phone);
            TrySetProperty(de, AdProperties.POSTALCODE, PostalCode);
            TrySetProperty(de, AdProperties.PHYSICALDELIVERYOFFICENAME, PhysicalDeliveryOfficeName);
            TrySetProperty(de, AdProperties.STATE, State);
            TrySetProperty(de, AdProperties.STREETADDRESS, StreetAddress);
            TrySetProperty(de, AdProperties.TITLE, Title);

            //_thumbnailPhoto = GetProperty(de, AdProperties.THUMBNAILPHOTO);

            //this.userAccountControl = TryGetResult<UserAccountControl>(de, AdProperties.USERACCOUNTCONTROL);

            //Finally, save all changes
            de.CommitChanges();
        }

        /// <summary>
        ///
        /// </summary>
        public void Save()
        {
            try
            {
                using (DirectorySearcher DirSearch = GetDirectorySearcher())
                {
                    DirSearch.Filter = String.Format("(&(objectClass=person)(objectGUID={0}))", GuidToHex(ObjectGuid));

                    SearchResult results = DirSearch.FindOne();

                    if (results != null)
                    {
                        using (DirectoryEntry de = GetDirectoryEntry(results.Path))
                        {
                            UpdateUser(de);
                        }
                    }
                }//end using
            }
            catch (Exception ex)
            {
                throw (new Exception("User cannot be updated" + ex.Message));
            }
        }

        /// <summary>
        /// Get user by Guid
        /// </summary>
        /// <param name="guid">Represents the GUID for the object we are searching for</param>
        /// <returns>AD_UserFull object</returns>
        public AD_UserFull GetUserByGuid(Guid guid)
        {
            try
            {
                using (DirectorySearcher DirSearch = GetDirectorySearcher())
                {
                    DirSearch.Filter = String.Format("(&(objectClass=person)(objectGUID={0}))", GuidToHex(guid));

                    SearchResult results = DirSearch.FindOne();

                    if (results == null)
                    {
                        return null;
                    }
                    else
                    {
                        //DirectoryEntry user = new DirectoryEntry(results.Path);

                        return GetUser(GetDirectoryEntry(results.Path));
                    }
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
                return null;
            }
        }

        /// <summary>
        /// Get User by SamAccountName
        /// </summary>
        /// <param name="samAccountName">Represents the SamAccountName for the object we are searching for</param>
        /// <returns>AD_UserFull object</returns>
        public AD_UserFull GetUserBySamAccountName(string samAccountName)
        {
            try
            {
                using (DirectorySearcher DirSearch = GetDirectorySearcher())
                {
                    DirSearch.Filter = String.Format("(&(objectClass=user)(SAMAccountName={0}))", samAccountName);

                    SearchResult results = DirSearch.FindOne();

                    if (results == null)
                    {
                        return null;
                    }
                    else
                    {
                        return GetUser(GetDirectoryEntry(results.Path));
                    }
                }//end using
            }
            catch (Exception ex)
            {
                ex.ToString();
                return null;
            }
        }

        /// <summary>
        /// Get user by DistinguishedName
        /// </summary>
        /// <param name="distinguishedName">DistinguishedName (CN=dvader,OU=Users,OUSites,DC=Eguibar,DC=local)</param>
        /// <returns>AD_UserFull object</returns>
        public AD_UserFull GetUserByDistinguishedName(string distinguishedName)
        {
            try
            {
                using (DirectorySearcher DirSearch = GetDirectorySearcher())
                {
                    DirSearch.Filter = String.Format("(&(objectClass=person)(distinguishedName={0}))", distinguishedName);

                    SearchResult results = DirSearch.FindOne();

                    if (results == null)
                    {
                        return null;
                    }
                    else
                    {
                        return GetUser(GetDirectoryEntry(results.Path));
                    }
                }
            }
            catch (Exception ex)
            {
                ex.ToString();
                return null;
            }
        }

        /// <summary>
        /// Get user by UserPrincipalName (samaccountname@eguibarIT.local)
        /// <param name="upn">UserPrincipalName (samaccountname@eguibarIT.local)</param>
        /// </summary>
        /// <returns>AD_UserFull object</returns>
        public AD_UserFull GetUserByUpn(string upn)
        {
            try
            {
                using (DirectorySearcher DirSearch = GetDirectorySearcher())
                {
                    DirSearch.Filter = String.Format("(&(objectClass=person)(userPrincipalName={0}))", upn);

                    SearchResult results = DirSearch.FindOne();

                    if (results == null)
                    {
                        return null;
                    }
                    else
                    {
                        return GetUser(new DirectoryEntry(results.Path));
                    }
                }//end using
            }
            catch (Exception ex)
            {
                ex.ToString();
                return null;
            }
        }

        /// <summary>
        /// Get user by email address (user@eguibarIT.com)
        /// <param name="email">Email address (user@eguibarIT.com)</param>
        /// </summary>
        public AD_UserFull GetUserByEmail(string email)
        {
            try
            {
                using (DirectorySearcher DirSearch = GetDirectorySearcher())
                {
                    DirSearch.Filter = String.Format("(&(objectClass=person)(mail={0}))", email);

                    SearchResult results = DirSearch.FindOne();

                    if (results == null)
                    {
                        return null;
                    }
                    else
                    {
                        return GetUser(GetDirectoryEntry(results.Path));
                    }
                }//end using
            }
            catch (Exception ex)
            {
                ex.ToString();
                return null;
            }
        }

        /// <summary>
        /// Get all users in Domain
        /// </summary>
        /// <param name="DomainName">Domain name where to retrieve users from</param>
        /// <returns>List of AD_UserFull object</returns>
        public List<AD_UserFull> GetAllUsersInDomain(string DomainName)
        {
            List<AD_UserFull> userlist = new List<AD_UserFull>();

            try
            {
                using (DirectorySearcher DirSearch = GetDirectorySearcher())
                {
                    DirSearch.Filter = "(&(objectCategory=person)(objectClass=user))";
                    DirSearch.SizeLimit = int.MaxValue;
                    DirSearch.PageSize = int.MaxValue;

                    SearchResultCollection userCollection = DirSearch.FindAll();

                    foreach (SearchResult users in userCollection)
                    {
                        AD_UserFull userInfo = GetUser(GetDirectoryEntry(users.Path));

                        userlist.Add(userInfo);
                    }
                }//end using
            }
            catch (Exception ex)
            {
                ex.ToString();
                return null;
            }

            return userlist;
        }

        /// <summary>
        /// Get all users in Domain
        /// </summary>
        /// <returns>List of AD_UserFull object from current domain</returns>
        public List<AD_UserFull> GetAllUsersInDomain()
        {
            List<AD_UserFull> userlist = GetAllUsersInDomain(AdDomain.GetAdFQDN());
            return userlist;
        }

        #endregion Members
    }//end class

    /// <summary>
    ///
    /// </summary>
    public interface UserMethods
    {
    }
}