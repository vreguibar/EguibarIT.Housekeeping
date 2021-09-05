namespace EguibarIT.Housekeeping.AdHelper
{
    /// <summary>
    /// Class representing an generic AD object
    /// This class is used to inherit
    /// </summary>
    public abstract class AD_Object : Helpers
    {
        #region Properties

        /// <summary>
        /// The bytes that make up the identifier.
        /// </summary>
        /// <returns>A byte array</returns>
        ///<remarks>This is the format that is stored in the Active Directory repository.</remarks>
        public abstract byte[] ObjectGuid { get; set; }

        /// <summary>
        /// The bytes that make up the sid.
        /// </summary>
        /// <returns>A byte array</returns>
        /// <remarks>This is the format that is stored in the Active Directory repository.</remarks>
        public abstract byte[] ObjectSID { get; set; }

        #endregion Properties
    }//end class
}//end namespace