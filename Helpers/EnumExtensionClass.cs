using System;
using System.Collections.Generic;
using System.Linq;

namespace Enum.Extensions
{
    /// <summary>
    ///
    /// </summary>
    public static class EnumerationExtensions
    {
        /// <summary>
        /// Check if Enum
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="withFlags">Flags to compare against.</param>
        private static void CheckIsEnum<T>(bool withFlags)
        {
            if (!typeof(T).IsEnum)
                throw new ArgumentException(string.Format("Type '{0}' is not an enum", typeof(T).FullName));
            if (withFlags && !Attribute.IsDefined(typeof(T), typeof(FlagsAttribute)))
                throw new ArgumentException(string.Format("Type '{0}' doesn't have the 'Flags' attribute", typeof(T).FullName));
        }

#pragma warning disable CS1570 // XML comment has badly formed XML -- 'Whitespace is not allowed at this location.'

        /// <summary>
        /// Check Is the Flag Set
        /// </summary>
        /// <typeparam name="T">The type of the enum.</typeparam>
        /// <param name="value">Value of the given Type</param>
        /// <param name="flag">Flags to compare against.</param>
        /// <returns>
        ///  <c>true</c> if the specified value has flags; otherwise, <c>false</c>.
        /// </returns>
        /// (a & b) == b
        public static bool IsFlagSet<T>(this T value, T flag) where T : struct, IConvertible, IComparable, IFormattable
#pragma warning restore CS1570 // XML comment has badly formed XML -- 'Whitespace is not allowed at this location.'
        {
            CheckIsEnum<T>(true);
            long lValue = Convert.ToInt64(value);
            long lFlag = Convert.ToInt64(flag);
            return (lValue & lFlag) != 0;
        }

        /// <summary>
        /// Get the Flags
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="value">Value of the given Type</param>
        /// <returns>IEnumerable object containing the Flags</returns>
        public static IEnumerable<T> GetFlags<T>(this T value) where T : struct, IConvertible, IComparable, IFormattable
        {
            CheckIsEnum<T>(true);
            foreach (T flag in System.Enum.GetValues(typeof(T)).Cast<T>())
            {
                if (value.IsFlagSet(flag))
                    yield return flag;
            }
        }

#pragma warning disable CS1570 // XML comment has badly formed XML -- 'Whitespace is not allowed at this location.'

        /// <summary>
        /// Set the Flags
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="value">Value of the given Type</param>
        /// <param name="flags">Flags to set up.</param>
        /// <param name="on">Turn ON the flag</param>
        /// <returns></returns>
        /// a | b  OR  a & (~b)
        public static T SetFlags<T>(this T value, T flags, bool on) where T : struct, IConvertible, IComparable, IFormattable
#pragma warning restore CS1570 // XML comment has badly formed XML -- 'Whitespace is not allowed at this location.'
        {
            CheckIsEnum<T>(true);
            long lValue = Convert.ToInt64(value);
            long lFlag = Convert.ToInt64(flags);
            if (on)
            {
                //Set flag
                lValue |= lFlag;
            }
            else
            {
                //Clear flag
                //a & (~b)
                lValue &= (~lFlag);
            }
            return (T)System.Enum.ToObject(typeof(T), lValue);
        }

        /// <summary>
        /// Overload - Set the Flags
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="value">Value of the given Type</param>
        /// <param name="flags">Flags to set up.</param>
        /// <returns></returns>
        public static T SetFlags<T>(this T value, T flags) where T : struct, IConvertible, IComparable, IFormattable
        {
            return value.SetFlags(flags, true);
        }

#pragma warning disable CS1570 // XML comment has badly formed XML -- 'Whitespace is not allowed at this location.'

        /// <summary>
        /// Clear the flag
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="value">Value of the given Type</param>
        /// <param name="flags">Flags to clear up.</param>
        /// <returns></returns>
        /// a & (~b)
        public static T ClearFlags<T>(this T value, T flags) where T : struct, IConvertible, IComparable, IFormattable
#pragma warning restore CS1570 // XML comment has badly formed XML -- 'Whitespace is not allowed at this location.'
        {
            return value.SetFlags(flags, false);
        }

        /// <summary>
        /// Combine flags.
        /// </summary>
        /// <typeparam name="T">Type</typeparam>
        /// <param name="flags">Flags to combine.</param>
        /// <returns></returns>
        public static T CombineFlags<T>(this IEnumerable<T> flags) where T : struct, IConvertible, IComparable, IFormattable
        {
            CheckIsEnum<T>(true);
            long lValue = 0;
            foreach (T flag in flags)
            {
                long lFlag = Convert.ToInt64(flag);
                lValue |= lFlag;
            }
            return (T)System.Enum.ToObject(typeof(T), lValue);
        }
    }
}//end namespace Enum.Extensions