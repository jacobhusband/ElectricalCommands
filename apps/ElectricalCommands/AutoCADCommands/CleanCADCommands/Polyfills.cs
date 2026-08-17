#if NET48
namespace System
{
    // This allows you to use the ^1 syntax in .NET 4.8
    public readonly struct Index : IEquatable<Index>
    {
        private readonly int _value;

        public int Value => _value < 0 ? ~_value : _value;
        public bool IsFromEnd => _value < 0;

        public Index(int value, bool fromEnd = false)
        {
            if (value < 0 && !fromEnd) throw new ArgumentOutOfRangeException(nameof(value));
            _value = fromEnd ? ~value : value;
        }

        public static implicit operator Index(int value) => FromStart(value);

        public static Index FromStart(int value) => new Index(value);
        public static Index FromEnd(int value) => new Index(value, true);

        public int GetOffset(int length)
        {
            int offset = _value;
            if (IsFromEnd) offset = length + ~offset + 1;
            return offset;
        }

        public override bool Equals(object value) => value is Index && _value == ((Index)value)._value;
        public bool Equals(Index other) => _value == other._value;
        public override int GetHashCode() => _value;
        public override string ToString() => IsFromEnd ? "^" + Value.ToString() : Value.ToString();
    }

    // This allows you to use the [1..] range syntax in .NET 4.8
    public readonly struct Range : IEquatable<Range>
    {
        public Index Start { get; }
        public Index End { get; }

        public Range(Index start, Index end)
        {
            Start = start;
            End = end;
        }

        public override bool Equals(object value) => value is Range r && r.Start.Equals(Start) && r.End.Equals(End);
        public bool Equals(Range other) => other.Start.Equals(Start) && other.End.Equals(End);
        public override int GetHashCode() => Start.GetHashCode() * 31 + End.GetHashCode();
        public override string ToString() => Start + ".." + End;

        public static Range All => new Range(Index.FromStart(0), Index.FromEnd(0));
    }
}
#endif