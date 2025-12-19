using System.Runtime.Serialization;

namespace Skrypton.LegacyParser.Tokens.Basic
{
    [DataContract(Namespace = "http://vbs")]
    public enum OperatorKind
    {
        [EnumMember] Unknown = 0,
        [EnumMember] Plus = 1,
        [EnumMember] Minus = 2,
        [EnumMember] Equal = 3,
        [EnumMember] NotEqual = 4,
        [EnumMember] LessThan = 5,
        [EnumMember] GreaterThan = 6,
        [EnumMember] LessThanOrEqual = 7,
        [EnumMember] GreaterThanOrEqual = 8,
        [EnumMember] IsSameObject = 9,
        [EnumMember] LogicalEquivalence = 10,
        [EnumMember] LogicalImplication = 11,

        [EnumMember] LogicalNot = 12,
        [EnumMember] LogicalAnd = 13,
        [EnumMember] LogicalOr = 14,
        [EnumMember] LogicalXor = 15,

        [EnumMember] Exponentiation = 16,
        [EnumMember] Division = 17,
        [EnumMember] IntegerDivision = 18,
        [EnumMember] Multiplication = 19,
        [EnumMember] Modulus = 20,
        [EnumMember] StringConcatenation = 21,

        [EnumMember] VerySpecial = 22,
    }
}
