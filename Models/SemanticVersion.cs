using System.Text.RegularExpressions;

namespace TaskTool.Models;

public sealed class SemanticVersion : IComparable<SemanticVersion>, IEquatable<SemanticVersion>
{
    private static readonly Regex Pattern = new(@"^[vV]?(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(?:-([0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*))?(?:\+[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?$", RegexOptions.Compiled);
    public int Major { get; } public int Minor { get; } public int Patch { get; } public string Prerelease { get; }
    public bool IsPrerelease => Prerelease.Length > 0;
    private SemanticVersion(int major, int minor, int patch, string prerelease) { Major = major; Minor = minor; Patch = patch; Prerelease = prerelease; }
    public static bool TryParse(string? text, out SemanticVersion version) { var m = Pattern.Match(text?.Trim() ?? ""); if (!m.Success) { version = null!; return false; } version = new(int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value), int.Parse(m.Groups[3].Value), m.Groups[4].Value); return true; }
    public static SemanticVersion Parse(string text) => TryParse(text, out var value) ? value : throw new FormatException($"Ungültige Semantic Version: {text}");
    public int CompareTo(SemanticVersion? other) { if (other is null) return 1; var core = Major.CompareTo(other.Major); if (core == 0) core = Minor.CompareTo(other.Minor); if (core == 0) core = Patch.CompareTo(other.Patch); if (core != 0) return core; if (!IsPrerelease) return other.IsPrerelease ? 1 : 0; if (!other.IsPrerelease) return -1; var a = Prerelease.Split('.'); var b = other.Prerelease.Split('.'); for (var i = 0; i < Math.Max(a.Length, b.Length); i++) { if (i == a.Length) return -1; if (i == b.Length) return 1; var an = int.TryParse(a[i], out var ai); var bn = int.TryParse(b[i], out var bi); var c = an && bn ? ai.CompareTo(bi) : an ? -1 : bn ? 1 : string.CompareOrdinal(a[i], b[i]); if (c != 0) return c; } return 0; }
    public override string ToString() => $"{Major}.{Minor}.{Patch}" + (IsPrerelease ? "-" + Prerelease : "");
    public bool Equals(SemanticVersion? other) => CompareTo(other) == 0; public override bool Equals(object? obj) => obj is SemanticVersion v && Equals(v); public override int GetHashCode() => HashCode.Combine(Major, Minor, Patch, Prerelease);
    public static bool operator >(SemanticVersion a, SemanticVersion b) => a.CompareTo(b) > 0; public static bool operator <(SemanticVersion a, SemanticVersion b) => a.CompareTo(b) < 0; public static bool operator >=(SemanticVersion a, SemanticVersion b) => a.CompareTo(b) >= 0; public static bool operator <=(SemanticVersion a, SemanticVersion b) => a.CompareTo(b) <= 0;
}
