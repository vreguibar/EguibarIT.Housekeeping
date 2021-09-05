using EguibarIT.Housekeeping.AdHelper;
using System;
using System.DirectoryServices.AccountManagement;
using System.Reflection;

namespace EguibarIT.Housekeeping.ExtPrincipal
{
    /// <summary>
    /// UserPrincipal extended attributes
    /// https://msdn.microsoft.com/en-us/library/ms683980(v=vs.85).aspx
    /// </summary>
    [DirectoryObjectClass("User")]
    [DirectoryRdnPrefix("CN")]
    public class UserPrincipalEx : System.DirectoryServices.AccountManagement.UserPrincipal
    {
        #region Constructors

        /// <summary>
        /// Inplement the constructor using the base class constructor.
        /// </summary>
        public UserPrincipalEx(PrincipalContext context) : base(context) { }

        /// <summary>
        /// Implement the constructor with initialization parameters.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="samAccountName"></param>
        /// <param name="password"></param>
        /// <param name="enabled"></param>
        public UserPrincipalEx(PrincipalContext context, string samAccountName, string password, bool enabled) : base(context, samAccountName, password, enabled) { }

        #endregion Constructors

        #region ExtendedProperties

        /// <summary>
        /// The AdminCount attribute.
        /// If 0 object is not protected by the AdminSDHolder process
        /// </summary>
        [DirectoryProperty(AdProperties.ADMINCOUNT)]
        public string adminCount
        {
            get
            {
                if (ExtensionGet(AdProperties.ADMINCOUNT).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.ADMINCOUNT)[0];
            }
            set { ExtensionSet(AdProperties.ADMINCOUNT, value); }
        }

        /// <summary>
        /// Represents the name of a locality, such as a town or city.
        /// </summary>
        [DirectoryProperty(AdProperties.CITY)]
        public string City
        {
            get
            {
                if (ExtensionGet(AdProperties.CITY).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.CITY)[0];
            }
            set { ExtensionSet(AdProperties.CITY, value); }
        }

        /// <summary>
        /// The user's company name
        /// </summary>
        [DirectoryProperty(AdProperties.COMPANY)]
        public string Company
        {
            get
            {
                if (ExtensionGet(AdProperties.COMPANY).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.COMPANY)[0];
            }
            set { ExtensionSet(AdProperties.COMPANY, value); }
        }

        /// <summary>
        /// The country/region in the address of the user. The country/region is represented as a 2-character code based on ISO-3166.
        /// </summary>
        [DirectoryProperty(AdProperties.COUNTRYNAME)]
        public string CountryName
        {
            get
            {
                if (ExtensionGet(AdProperties.COUNTRYNAME).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.COUNTRYNAME)[0];
            }
            set { ExtensionSet(AdProperties.COUNTRYNAME, value); }
        }

        /// <summary>
        /// Contains the name for the department in which the user works.
        /// </summary>
        [DirectoryProperty(AdProperties.DEPARTMENT)]
        public string Department
        {
            get
            {
                if (ExtensionGet(AdProperties.DEPARTMENT).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.DEPARTMENT)[0];
            }
            set { ExtensionSet(AdProperties.DEPARTMENT, value); }
        }

        /// <summary>
        /// The user's division.
        /// </summary>
        [DirectoryProperty(AdProperties.DIVISION)]
        public string Division
        {
            get
            {
                if (ExtensionGet(AdProperties.DIVISION).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.DIVISION)[0];
            }
            set { ExtensionSet(AdProperties.DIVISION, value); }
        }

        /// <summary>
        /// The number assigned to an employee other than the ID
        /// </summary>
        [DirectoryProperty(AdProperties.EMPLOYEENUMBER)]
        public string employeeNumber
        {
            get
            {
                if (ExtensionGet(AdProperties.EMPLOYEENUMBER).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.EMPLOYEENUMBER)[0];
            }
            set { ExtensionSet(AdProperties.EMPLOYEENUMBER, value); }
        }

        /// <summary>
        /// The job category for an employee
        /// </summary>
        [DirectoryProperty(AdProperties.EMPLOYEETYPE)]
        public string EmployeeType
        {
            get
            {
                if (ExtensionGet(AdProperties.EMPLOYEETYPE).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.EMPLOYEETYPE)[0];
            }
            set { ExtensionSet(AdProperties.EMPLOYEETYPE, value); }
        }

        /// <summary>
        /// This is the time that the user last logged into the domain.
        /// This value is stored as a large integer that represents the number of 100-nanosecond intervals since January 1, 1601 (UTC).
        /// Whenever a user logs on, the value of this attribute is read from the DC.
        /// If the value is older [ current_time - msDS-LogonTimeSyncInterval ], the value is updated.
        /// The initial update after the raise of the domain functional level is calculated as 14 days minus random percentage of 5 days.
        /// Read Only
        /// </summary>
        [DirectoryProperty(AdProperties.LASTLOGONTIMESTAMP)]
        public DateTime? LastLogonTimeStamp
        {
            get
            {
                if (ExtensionGet(AdProperties.LASTLOGONTIMESTAMP).Length > 0)
                {
                    var lastLogonDate = ExtensionGet(AdProperties.LASTLOGONTIMESTAMP)[0];
                    var lastLogonDateType = lastLogonDate.GetType();

                    var highPart = (Int32)lastLogonDateType.InvokeMember("HighPart", BindingFlags.GetProperty, null, lastLogonDate, null);
                    var lowPart = (Int32)lastLogonDateType.InvokeMember("LowPart", BindingFlags.GetProperty | BindingFlags.Public, null, lastLogonDate, null);

                    var longDate = ((Int64)highPart << 32 | (UInt32)lowPart);

                    return longDate > 0 ? (DateTime?)DateTime.FromFileTime(longDate) : null;
                }

                return null;
            }
        }

        /// <summary>
        /// Indicates whether the account has permission to dial in to the RAS server
        /// </summary>
        [DirectoryProperty(AdProperties.MSNPALLOWDIALIN)]
        public bool AllowDialin
        {
            get
            {
                if (ExtensionGet(AdProperties.MSNPALLOWDIALIN).Length != 1)
                    return false;

                return (bool)ExtensionGet(AdProperties.MSNPALLOWDIALIN)[0];
            }
            set { ExtensionSet(AdProperties.MSNPALLOWDIALIN, value); }
        }

        /// <summary>
        /// The primary mobile phone number
        /// </summary>
        [DirectoryProperty(AdProperties.MOBILE)]
        public string MobilePhone
        {
            get
            {
                if (ExtensionGet(AdProperties.MOBILE).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.MOBILE)[0];
            }
            set { ExtensionSet(AdProperties.MOBILE, value); }
        }

        /// <summary>
        /// Contains the office location in the user's place of business
        /// </summary>
        [DirectoryProperty(AdProperties.PHYSICALDELIVERYOFFICENAME)]
        public string Office
        {
            get
            {
                if (ExtensionGet(AdProperties.PHYSICALDELIVERYOFFICENAME).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.PHYSICALDELIVERYOFFICENAME)[0];
            }
            set { ExtensionSet(AdProperties.PHYSICALDELIVERYOFFICENAME, value); }
        }

        /// <summary>
        /// The name of the company or organization
        /// </summary>
        [DirectoryProperty(AdProperties.ORGANIZATIONNAME)]
        public string Organization
        {
            get
            {
                if (ExtensionGet(AdProperties.ORGANIZATIONNAME).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.ORGANIZATIONNAME)[0];
            }
            set { ExtensionSet(AdProperties.ORGANIZATIONNAME, value); }
        }

        /// <summary>
        /// The postal or zip code for mail delivery
        /// </summary>
        [DirectoryProperty(AdProperties.POSTALCODE)]
        public string PostalCode
        {
            get
            {
                if (ExtensionGet(AdProperties.POSTALCODE).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.POSTALCODE)[0];
            }
            set { ExtensionSet(AdProperties.POSTALCODE, value); }
        }

        /// <summary>
        /// The post office box number for this object
        /// </summary>
        [DirectoryProperty(AdProperties.POBOX)]
        public string POBox
        {
            get
            {
                if (ExtensionGet(AdProperties.POBOX).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.POBOX)[0];
            }
            set { ExtensionSet(AdProperties.POBOX, value); }
        }

        /// <summary>
        /// The date and time that the password for this account was last changed.
        /// This value is stored as a large integer that represents the number of 100 nanosecond intervals since January 1, 1601 (UTC).
        /// If this value is set to 0 and the User-Account-Control attribute does not contain the UF_DONT_EXPIRE_PASSWD flag, then the user must set the password at the next logon.
        /// Read Only
        /// </summary>
        [DirectoryProperty(AdProperties.PWDLASTSET)]
        public DateTime? PwdLastSet
        {
            get
            {
                if (ExtensionGet(AdProperties.PWDLASTSET).Length > 0)
                {
                    var pwdLastSet = ExtensionGet(AdProperties.PWDLASTSET)[0];
                    var pwdLastSetType = pwdLastSet.GetType();

                    var highPart = (Int32)pwdLastSetType.InvokeMember("HighPart", BindingFlags.GetProperty, null, pwdLastSet, null);
                    var lowPart = (Int32)pwdLastSetType.InvokeMember("LowPart", BindingFlags.GetProperty | BindingFlags.Public, null, pwdLastSet, null);

                    var longDate = ((Int64)highPart << 32 | (UInt32)lowPart);

                    return longDate > 0 ? (DateTime?)DateTime.FromFileTime(longDate) : null;
                }

                return null;
            }
        }

        /// <summary>
        /// The name of a user's state or province
        /// </summary>
        [DirectoryProperty(AdProperties.STATE)]
        public string State
        {
            get
            {
                if (ExtensionGet(AdProperties.STATE).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.STATE)[0];
            }
            set { ExtensionSet(AdProperties.STATE, value); }
        }

        /// <summary>
        /// The street address
        /// </summary>
        [DirectoryProperty(AdProperties.STREETADDRESS)]
        public string StreetAddress
        {
            get
            {
                if (ExtensionGet(AdProperties.STREETADDRESS).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.STREETADDRESS)[0];
            }
            set { ExtensionSet(AdProperties.STREETADDRESS, value); }
        }

        /// <summary>
        /// The encryption algorithms supported by user, computer or trust accounts.
        /// NOTE: The KDC uses this information while generating a service ticket for this account.
        /// Services and Computers can automatically update this attribute on their respective accounts in Active Directory, and therefore need write access to this attribute.
        /// </summary>
        [DirectoryProperty(AdProperties.SUPPORTEDENCRYPTIONTYPES)]
        public int SupportedEncryptionTypes
        {
            get
            {
                if (ExtensionGet(AdProperties.SUPPORTEDENCRYPTIONTYPES).Length < 1)
                    return 0;

                return (int)ExtensionGet(AdProperties.SUPPORTEDENCRYPTIONTYPES)[0];
            }
            set { ExtensionSet(AdProperties.SUPPORTEDENCRYPTIONTYPES, value); }
        }

        /// <summary>
        /// The primary telephone number
        /// </summary>
        [DirectoryProperty(AdProperties.TELEPHONENUMBER)]
        public string TelephoneNumber
        {
            get
            {
                if (ExtensionGet(AdProperties.TELEPHONENUMBER).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.TELEPHONENUMBER)[0];
            }
            set { ExtensionSet(AdProperties.TELEPHONENUMBER, value); }
        }

        /// <summary>
        /// Contains the user's job title.
        /// This property is commonly used to indicate the formal job title, such as Senior Programmer, rather than occupational class, such as programmer.
        /// It is not typically used for suffix titles such as Esq. or DDS.
        /// </summary>
        [DirectoryProperty(AdProperties.TITLE)]
        public string Title
        {
            get
            {
                if (ExtensionGet(AdProperties.TITLE).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.TITLE)[0];
            }
            set { ExtensionSet(AdProperties.TITLE, value); }
        }

        /// <summary>
        /// Flags that control the behavior of the user account
        /// </summary>
        [DirectoryProperty(AdProperties.USERACCOUNTCONTROL)]
        public int UserAccountControl
        {
            get
            {
                if (ExtensionGet(AdProperties.USERACCOUNTCONTROL).Length <= 0)
                    return 0;

                return (int)ExtensionGet(AdProperties.USERACCOUNTCONTROL)[0];
            }
            set { ExtensionSet(AdProperties.USERACCOUNTCONTROL, value); }
        }

        #endregion ExtendedProperties

        /// <summary>
        /// Implement the overloaded search method FindByIdentity.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="identityValue"></param>
        /// <returns></returns>
        public static new UserPrincipalEx FindByIdentity(PrincipalContext context, string identityValue)
        {
            return (UserPrincipalEx)FindByIdentityWithType(context, typeof(UserPrincipalEx), identityValue);
        }

        /// <summary>
        /// Implement the overloaded search method FindByIdentity.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="identityType"></param>
        /// <param name="identityValue"></param>
        /// <returns></returns>
        public static new UserPrincipalEx FindByIdentity(PrincipalContext context, IdentityType identityType, string identityValue)
        {
            return (UserPrincipalEx)FindByIdentityWithType(context, typeof(UserPrincipalEx), identityType, identityValue);
        }

        internal static UserPrincipalEx FindByIdentity(PrincipalContext contextManager)
        {
            throw new NotImplementedException();
        }
    } // end class

    /// <summary>
    /// ComputerPrincipal TelephoneNumber to include attributes
    /// https://msdn.microsoft.com/en-us/library/ms680987(v=vs.85).aspx
    /// </summary>
    [DirectoryObjectClass("computer")]
    [DirectoryRdnPrefix("CN")]
    public class ComputerPrincipalEx : System.DirectoryServices.AccountManagement.ComputerPrincipal
    {
        #region Constructor

        /// <summary>
        ///
        /// </summary>
        public ComputerPrincipalEx(PrincipalContext context) : base(context) { }

        //public ComputerPrincipalEx(PrincipalContext context, string samAccountName) : base( (context, samAccountName) {}

        #endregion Constructor

        #region ExtendedProperties

        /// <summary>
        /// This is the time that the computer last logged into the domain.
        /// This value is stored as a large integer that represents the number of 100-nanosecond intervals since January 1, 1601 (UTC).
        /// Whenever a user logs on, the value of this attribute is read from the DC.
        /// If the value is older [ current_time - msDS-LogonTimeSyncInterval ], the value is updated.
        /// The initial update after the raise of the domain functional level is calculated as 14 days minus random percentage of 5 days.
        /// Read Only
        /// </summary>
        [DirectoryProperty(AdProperties.LASTLOGONTIMESTAMP)]
        public DateTime? LastLogonTimeStamp
        {
            get
            {
                if (ExtensionGet(AdProperties.LASTLOGONTIMESTAMP).Length > 0)
                {
                    var lastLogonDate = ExtensionGet(AdProperties.LASTLOGONTIMESTAMP)[0];
                    var lastLogonDateType = lastLogonDate.GetType();

                    var highPart = (Int32)lastLogonDateType.InvokeMember("HighPart", BindingFlags.GetProperty, null, lastLogonDate, null);
                    var lowPart = (Int32)lastLogonDateType.InvokeMember("LowPart", BindingFlags.GetProperty | BindingFlags.Public, null, lastLogonDate, null);

                    var longDate = ((Int64)highPart << 32 | (UInt32)lowPart);

                    return longDate > 0 ? (DateTime?)DateTime.FromFileTime(longDate) : null;
                }

                return null;
            }
        }

        /// <summary>
        /// The Operating System name, for example, Windows Vista Enterprise
        /// </summary>
        [DirectoryProperty(AdProperties.OPERATINGSYSTEM)]
        public string OperatingSystem
        {
            get
            {
                if (ExtensionGet(AdProperties.OPERATINGSYSTEM).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.OPERATINGSYSTEM)[0];
            }
            set { this.ExtensionSet(AdProperties.OPERATINGSYSTEM, value); }
        }

        /// <summary>
        /// The operating system version string, for example, 4.0.
        /// </summary>
        [DirectoryProperty(AdProperties.OPERATINGSYSTEMVERSION)]
        public string OperatingSystemVersion
        {
            get
            {
                if (ExtensionGet(AdProperties.OPERATINGSYSTEMVERSION).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.OPERATINGSYSTEMVERSION)[0];
            }
            set { this.ExtensionSet(AdProperties.OPERATINGSYSTEMVERSION, value); }
        }

        /// <summary>
        /// Name of computer as registered in DNS
        /// </summary>
        [DirectoryProperty(AdProperties.DNSHOSTNAME)]
        public string DnsHostName
        {
            get
            {
                if (ExtensionGet(AdProperties.DNSHOSTNAME).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.DNSHOSTNAME)[0];
            }
            set { this.ExtensionSet(AdProperties.DNSHOSTNAME, value); }
        }

        #endregion ExtendedProperties

        /// <summary>
        /// Implement the overloaded search method FindByIdentity.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="identityValue"></param>
        /// <returns></returns>
        public static new ComputerPrincipalEx FindByIdentity(PrincipalContext context, string identityValue)
        {
            return (ComputerPrincipalEx)FindByIdentityWithType(context, typeof(ComputerPrincipalEx), identityValue);
        }

        /// <summary>
        /// Implement the overloaded search method FindByIdentity.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="identityType"></param>
        /// <param name="identityValue"></param>
        /// <returns></returns>
        public static new ComputerPrincipalEx FindByIdentity(PrincipalContext context, IdentityType identityType, string identityValue)
        {
            return (ComputerPrincipalEx)FindByIdentityWithType(context, typeof(ComputerPrincipalEx), identityType, identityValue);
        }

        internal static ComputerPrincipalEx FindByIdentity(ComputerPrincipalEx contextManager)
        {
            throw new NotImplementedException();
        }
    } // end class

    /// <summary>
    /// UserPrincipal TelephoneNumber to include attributes
    /// </summary>
    [DirectoryObjectClass("group")]
    [DirectoryRdnPrefix("CN")]
    public class GroupPrincipalEx : System.DirectoryServices.AccountManagement.GroupPrincipal
    {
        #region Constructor

        /// <summary>
        /// Inplement the constructor using the base class constructor.
        /// </summary>
        public GroupPrincipalEx(PrincipalContext context) : base(context) { }

        /// <summary>
        /// Implement the constructor with initialization parameters.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="samAccountName"></param>
        /// <param name="password"></param>
        /// <param name="enabled"></param>
        public GroupPrincipalEx(PrincipalContext context, string samAccountName, string password, bool enabled) : base(context, samAccountName) { }

        #endregion Constructor

        #region ExtendedProperties

        /// <summary>
        /// The AdminCount attribute.
        /// If 0 object is not protected by the AdminSDHolder process
        /// </summary>
        [DirectoryProperty(AdProperties.ADMINCOUNT)]
        public string adminCount
        {
            get
            {
                if (ExtensionGet(AdProperties.ADMINCOUNT).Length != 1)
                    return null;

                return (string)ExtensionGet(AdProperties.ADMINCOUNT)[0];
            }
            set { ExtensionSet(AdProperties.ADMINCOUNT, value); }
        }

        #endregion ExtendedProperties

        /// <summary>
        /// Implement the overloaded search method FindByIdentity.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="identityValue"></param>
        /// <returns></returns>
        public static new GroupPrincipalEx FindByIdentity(PrincipalContext context, string identityValue)
        {
            return (GroupPrincipalEx)FindByIdentityWithType(context, typeof(GroupPrincipalEx), identityValue);
        }

        /// <summary>
        /// Implement the overloaded search method FindByIdentity.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="identityType"></param>
        /// <param name="identityValue"></param>
        /// <returns></returns>
        public static new GroupPrincipalEx FindByIdentity(PrincipalContext context, IdentityType identityType, string identityValue)
        {
            return (GroupPrincipalEx)FindByIdentityWithType(context, typeof(GroupPrincipalEx), identityType, identityValue);
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="contextManager"></param>
        /// <returns></returns>
        internal static GroupPrincipalEx FindByIdentity(PrincipalContext contextManager)
        {
            throw new NotImplementedException();
        }
    } // end class

    // Example
    /*
    [DirectoryRdnPrefix("CN")]
    [DirectoryObjectClass("Person")]
    public class ExtendedPrincipalUser : UserPrincipal

    {
        // Inplement the constructor using the base class constructor.

        public ExtendedPrincipalUser(PrincipalContext ctx) : base(ctx) { }
        // Implement the constructor with initialization parameters.

        public ExtendedPrincipalUser(PrincipalContext ctx, string samAccountName, string password, bool enabled) : base(ctx, samAccountName, password, enabled) { }

        [DirectoryProperty("facsimileTelephoneNumber")]
        public string facsimileTelephoneNumber
        {
            get
            {
                if (ExtensionGet("facsimileTelephoneNumber").Length != 1)
                    return null;
                return (string)ExtensionGet("facsimileTelephoneNumber")[0];
            }
            set
            {
                this.ExtensionSet("facsimileTelephoneNumber", value);
            }
        }

        [DirectoryProperty("title")]
        public string jobTitle

        {
            get

            {
                if (ExtensionGet("title").Length != 1)

                    return null;

                return (string)ExtensionGet("title")[0];
            }
            set

            {
                this.ExtensionSet("title", value);
            }
        }

        [DirectoryProperty("manager")]
        public string manager

        {
            get

            {
                if (ExtensionGet("manager").Length != 1)

                    return string.Empty;

                return (string)ExtensionGet("manager")[0];
            }
            set

            {
                this.ExtensionSet("manager", value);
            }
        }

        [DirectoryProperty("facsimileTelephoneNumber")]
        public string fax
        {
            get

            {
                if (ExtensionGet("facsimileTelephoneNumber").Length != 1)

                    return null;

                return (string)ExtensionGet("facsimileTelephoneNumber")[0];
            }
            set

            {
                this.ExtensionSet("facsimileTelephoneNumber", value);
            }
        }

        [DirectoryProperty("department")]
        public string department
        {
            get

            {
                if (ExtensionGet("department").Length != 1)

                    return null;

                return (string)ExtensionGet("department")[0];
            }
            set

            {
                this.ExtensionSet("department", value);
            }
        }

        [DirectoryProperty("division")]
        public string division
        {
            get

            {
                if (ExtensionGet("division").Length != 1)

                    return null;

                return (string)ExtensionGet("division")[0];
            }
            set

            {
                this.ExtensionSet("division", value);
            }
        }

        [DirectoryProperty("mobile")]
        public string mobile
        {
            get

            {
                if (ExtensionGet("mobile").Length != 1)

                    return null;

                return (string)ExtensionGet("mobile")[0];
            }
            set

            {
                this.ExtensionSet("mobile", value);
            }
        }

        [DirectoryProperty("displayname")]
        public string displayName
        {
            get

            {
                if (ExtensionGet("displayname").Length != 1)

                    return null;

                return (string)ExtensionGet("displayname")[0];
            }
            set

            {
                this.ExtensionSet("displayname", value);
            }
        }

        [DirectoryProperty("mail")]
        public string mail
        {
            get

            {
                if (ExtensionGet("mail").Length != 1)

                    return null;

                return (string)ExtensionGet("mail")[0];
            }
            set

            {
                this.ExtensionSet("mail", value);
            }
        }

        // Implement the overloaded search method FindByIdentity.

        public static new ExtendedPrincipalUser FindByIdentity(PrincipalContext context, string identityValue)
        {
            return (ExtendedPrincipalUser)FindByIdentityWithType(context, typeof(ExtendedPrincipalUser), identityValue);
        }

        // Implement the overloaded search method FindByIdentity.
        public static new ExtendedPrincipalUser FindByIdentity(PrincipalContext context, IdentityType identityType, string identityValue)
        {
            return (ExtendedPrincipalUser)FindByIdentityWithType(context, typeof(ExtendedPrincipalUser), identityType, identityValue);
        }

        internal static ExtendedPrincipalUser FindByIdentity(PrincipalContext contextManager)
        {
            throw new NotImplementedException();
        }
    }

    */
}//end namespace