using System;
using System.Text;

namespace ElectricalCommands
{
  internal static class SwitchCircuitSequence
  {
    internal const string Alphabet = "abcdefghjkmnpqrstuvwxyz";

    internal static string ToSuffix(int index)
    {
      if (index < 0)
      {
        throw new ArgumentOutOfRangeException("index");
      }

      int value = checked(index + 1);
      StringBuilder result = new StringBuilder();
      while (value > 0)
      {
        value--;
        result.Insert(0, Alphabet[value % Alphabet.Length]);
        value /= Alphabet.Length;
      }

      return result.ToString();
    }

    internal static bool TryParseSuffix(string suffix, out int index)
    {
      index = -1;
      string candidate = (suffix ?? string.Empty).Trim();
      if (candidate.Length == 0)
      {
        return false;
      }

      try
      {
        int value = 0;
        foreach (char character in candidate)
        {
          int position = Alphabet.IndexOf(character);
          if (position < 0)
          {
            return false;
          }

          value = checked(value * Alphabet.Length + position + 1);
        }

        index = checked(value - 1);
        return true;
      }
      catch (OverflowException)
      {
        return false;
      }
    }
  }
}
